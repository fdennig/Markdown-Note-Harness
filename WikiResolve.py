"""WikiResolve.py -- stamp from_wiki:/to_wiki: onto the note you are saving.

This is the thin, editor-shaped half of the feature. All the logic lives in the
**vault**, at `<vault>/scripts/wikiresolve_core.py`, and is shared with
`<vault>/scripts/resolve_wiki.py`, which owns the corpus pass.

Why the split. The batch pass does not need an editor, and both incidents it
caused were editor-lifecycle incidents rather than logic errors: a plugin that
rewrote 1,745 notes three seconds after launch because a setting defaulted the
wrong way, and a crash on every email note opened because it reverted a view
from inside that view's own load callback. Neither was reachable by the test
suite, because the tested surface (pure text transformation) and the risky
surface (Sublime's plugin lifecycle) barely overlapped. Moving the batch work to
a script you invoke, whose report you read before it writes, removes both
classes of failure and makes the logic testable without an editor.

What is left here is the per-note work, which is where an editor earns its keep:

  - `on_pre_save`   edits the buffer being written, so a note you touch is
                    correct the moment you save it. No second write, nothing to
                    revert.
  - `on_load_async` writes the file you just opened, so notes arriving from the
                    Outlook macro are resolved when you first read them. It
                    **never reverts the view** -- that is what crashed Sublime,
                    and Sublime's own watcher reloads a clean buffer anyway.

Deliberately NOT here any more:
  - the corpus sweep, the report, the cleanup pass  -> scripts/resolve_wiki.py
  - the (mtime, size, gaps, fingerprint) cache      -> gone; a deliberate run
                                                       reads everything in ~2s
  - the self_people setting                         -> now `self: true` in the
                                                       person note, so the fact
                                                       travels with the vault

The index still lives in memory, because it must: on_pre_save runs on the main
thread and has to answer in microseconds. It is the only reason this module
holds any state at all -- if the on-save stamp is ever dropped, the whole
module goes with it.
"""

import sublime
import sublime_plugin

import datetime
import os
import sys

SETTINGS_FILE = "WikiResolve.sublime-settings"

# The head of a buffer read for the class gate. Generous because the gate must
# see 'class:', and an Outlook store_id plus a long recipient list pushes the
# rest of the frontmatter a long way down.
HEAD_CHARS = 16384

# Set once the vault's core module is importable. None means every hook is a
# no-op, which is the correct behaviour when the vault is not open.
core = None
_core_error = None


# ------------------------------------------------------------------
# Settings and logging
# ------------------------------------------------------------------

def settings():
    return sublime.load_settings(SETTINGS_FILE)


def log(message):
    """Append to Data/Logs/WikiResolve.log. Silent if logging itself fails."""
    if not settings().get("verbose", True):
        return
    try:
        log_dir = os.path.join(os.path.dirname(sublime.packages_path()), "Logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(os.path.join(log_dir, "WikiResolve.log"), "a",
                  encoding="utf-8") as f:
            f.write("[{0}] {1}\n".format(stamp, message))
    except Exception as e:
        print("WikiResolve: could not write log: {0}".format(e))


def roots(window):
    """Where person notes may live.

    Defaults to the window's own folders, which is what makes the vault owner's
    note findable even though it sits in a different tree from the vault. Set
    `person_roots` when you need the plugin and scripts/resolve_wiki.py to agree
    regardless of which project happens to be open -- they resolve identically
    only when they walk the same roots.
    """
    configured = settings().get("person_roots") or []
    if configured:
        return [os.path.abspath(os.path.expanduser(os.path.expandvars(p)))
                for p in configured]
    return list(window.folders()) if window else []


# ------------------------------------------------------------------
# Finding the core
# ------------------------------------------------------------------

def load_core(window):
    """Import the vault's wikiresolve_core, or leave `core` as None.

    The logic lives in the vault rather than in this package so that the batch
    script does not have to depend on Sublime's install path -- the coupling the
    split exists to remove. The cost is this lookup, and it is why every hook
    below tolerates `core` being None.
    """
    global core, _core_error
    if core is not None:
        return core
    candidates = []
    configured = (settings().get("core_path") or "").strip()
    if configured:
        candidates.append(os.path.abspath(
            os.path.expanduser(os.path.expandvars(configured))))
    for folder in (window.folders() if window else []):
        candidates.append(os.path.join(folder, "scripts"))
    for cand in candidates:
        directory = cand if os.path.isdir(cand) else os.path.dirname(cand)
        if not os.path.isfile(os.path.join(directory, "wikiresolve_core.py")):
            continue
        try:
            if directory not in sys.path:
                sys.path.insert(0, directory)
            import wikiresolve_core
            core = wikiresolve_core
            log("core loaded from {0}".format(directory))
            return core
        except Exception as e:
            _core_error = "{0}: {1}".format(directory, e)
    if _core_error is None:
        _core_error = "wikiresolve_core.py not found under any of {0}".format(
            candidates)
    log("core NOT loaded -- {0}".format(_core_error))
    return None


# ------------------------------------------------------------------
# The in-memory index
# ------------------------------------------------------------------

_index = None


def current_index():
    return _index


def rebuild(window):
    """Full rebuild: walk the roots, read every person note."""
    global _index
    if not load_core(window):
        return None
    index = core.build_index(roots(window), tuple(
        settings().get("person_classes", ["person"])),
        settings().get("exclude", []))
    _index = index
    log("index: {0}, fingerprint {1}".format(index.summary(),
                                             index.fingerprint[:12]))
    if index.ambiguous:
        # Ambiguity used to be visible only if you went looking for it. A key
        # two notes both claim silently costs you every resolution that would
        # have gone through it.
        sublime.status_message(
            "WikiResolve: {0} ambiguous name(s) -- see Wiki: Show person index"
            .format(len(index.ambiguous)))
    return index


def refresh_person(window, path):
    """One person note changed: re-read just that note and rebuild the tables.

    A full rebuild walks the whole project reading heads; this touches one file
    and then re-derives the four tables from records already in hand. That is
    what makes bulk person-note migration cheap -- the old design re-read all
    336 notes on every single save.
    """
    global _index
    if not load_core(window) or _index is None:
        return rebuild(window)
    record = core.read_person(path)
    name = core.display_name(path)
    people = dict(_index.people)
    if record:
        people[name] = record
    else:
        people.pop(name, None)
    _index = core.build_tables(people)
    log("index refreshed from {0}: {1}".format(name, _index.summary()))
    return _index


# ------------------------------------------------------------------
# The listener
# ------------------------------------------------------------------

def head_class(view):
    """The buffer's frontmatter class, from the buffer already in hand.

    No file read and no index lookup, so saving an unrelated note costs one scan
    over the head of the buffer and nothing else.
    """
    if not core:
        return None
    head = view.substr(sublime.Region(0, min(view.size(), HEAD_CHARS)))
    return core.head_class(head)


class WikiResolveStampCommand(sublime_plugin.TextCommand):
    """Bring the buffer's derived keys up to date.

    Called from on_pre_save, so the edit is folded into the write already in
    flight: no revert, no dirty buffer left behind, and the user's undo history
    still walks back through their own edit. Only the affected lines are
    replaced, so cursor and selection survive.
    """

    def run(self, edit):
        index = current_index()
        if not core or not index or not index.people:
            return
        text = self.view.substr(sublime.Region(0, self.view.size()))
        edits, _ = core.analyse(text, index, core.GENERATIVE)
        for start, end, want in edits:      # already descending
            region = sublime.Region(self.view.text_point(start, 0),
                                    self.view.text_point(end, 0))
            self.view.replace(edit, region, "".join(want))


class WikiResolveListener(sublime_plugin.EventListener):
    """Two hooks, and neither of them writes a file.

    on_pre_save changes the buffer Sublime is about to write, so there is no
    second write and nothing to revert. That is the whole reason this hook is
    safe where on_load_async was not.
    """

    def on_pre_save(self, view):
        """Stamp the buffer before Sublime writes it.

        Synchronous on the main thread, which is affordable because resolution
        is dictionary lookups against an index that is already built. A missing
        core or a cold index does nothing at all; scripts/resolve_wiki.py picks
        the note up later.
        """
        path = view.file_name()
        if not core or not path or not path.lower().endswith(".md"):
            return
        if head_class(view) not in core.LEAF_SPEC:
            return
        index = current_index()
        if not index or not index.people:
            return
        view.run_command("wiki_resolve_stamp")

    def on_load_async(self, view):
        """Stamp a note as it is opened.

        This is the hook that keeps up with notes nobody edits: the Outlook
        macro writes them while Sublime is not running, and you would otherwise
        not see resolved recipients until the next scripts/resolve_wiki.py run.

        **It writes the file and stops there. It must never revert the view.**
        An earlier version called view.run_command("revert") from here so the
        reader saw the stamped text at once, and that crashed Sublime on every
        email note opened -- reverting a view from inside that view's own load
        callback, multiplied across the burst of tabs "Email: Open all to-reply"
        produces. Removing the revert costs nothing: Sublime's own file watcher
        reloads an unmodified buffer when the file changes underneath it, which
        is the same result without re-entering the load cycle.

        Two things now make the write itself far safer than it was. The pass is
        monotone, so the worst it can do is add a link it should not have -- it
        cannot clear a value or shorten a list. And it is announced, because a
        hook that silently rewrites the file you just opened is a hook you
        cannot debug.
        """
        if not settings().get("on_load", True):
            return
        path = view.file_name()
        if not core or not path or not path.lower().endswith(".md"):
            return
        if head_class(view) not in core.LEAF_SPEC:
            return
        index = current_index()
        if not index or not index.people:
            return
        changed, stats, error = core.stamp_file(path, index,
                                                mode=core.GENERATIVE)
        if error:
            log("ERROR on load of {0}: {1}".format(path, error))
        elif changed:
            log("on_load stamped {0}".format(os.path.basename(path)))
            sublime.status_message(
                "WikiResolve: resolved {0} of {1} on open".format(
                    stats["resolved"], stats["total"]))

    def on_post_save_async(self, view):
        """A saved person note changes the index. Nothing is re-resolved here."""
        path = view.file_name()
        window = view.window() or sublime.active_window()
        if not core or not path or not window or not path.lower().endswith(".md"):
            return
        if head_class(view) in settings().get("person_classes", ["person"]):
            refresh_person(window, path)


def _warm_up():
    window = sublime.active_window()
    if window:
        rebuild(window)


def plugin_loaded():
    conf = settings()
    conf.clear_on_change("WikiResolve")
    conf.add_on_change("WikiResolve", _on_settings_change)
    # Deliberately only builds the index. The corpus pass that used to run here
    # is scripts/resolve_wiki.py, invoked by hand.
    sublime.set_timeout_async(_warm_up, 3000)


def _on_settings_change():
    global core
    core = None                 # core_path may have moved
    window = sublime.active_window()
    if window:
        sublime.set_timeout_async(lambda: rebuild(window), 0)


def plugin_unloaded():
    settings().clear_on_change("WikiResolve")


# ------------------------------------------------------------------
# Commands -- diagnostics only. The corpus pass is a script.
# ------------------------------------------------------------------

class WikiResolveRebuildIndexCommand(sublime_plugin.WindowCommand):
    """Re-read every person note. Needed when notes change outside Sublime."""

    def run(self):
        sublime.set_timeout_async(self._run, 0)

    def _run(self):
        index = rebuild(self.window)
        self.window.status_message(
            "WikiResolve: " + (index.summary() if index else
                               "core not found -- {0}".format(_core_error)))


class WikiResolveShowIndexCommand(sublime_plugin.WindowCommand):
    """Which person notes the resolver can see, and which names collide."""

    def run(self):
        sublime.set_timeout_async(self._run, 0)

    def _run(self):
        index = rebuild(self.window)
        self.paths = []
        rows = []
        if index is None:
            rows = [["Core not found", str(_core_error)]]
            self.paths = [None]
        else:
            for k in sorted(index.ambiguous):
                self.paths.append(None)
                rows.append(["(ambiguous) {0}".format(k),
                             "claimed by " + ", ".join(index.ambiguous[k])])
            for name in sorted(index.people):
                record = index.people[name]
                self.paths.append(record["path"])
                rows.append([
                    name + ("  (self)" if name in index.mine else ""),
                    "{0} addresses, {1} aliases  |  {2}".format(
                        len(record["emails"]), len(record["aliases"]),
                        record["title"] or "(no title)")])
            if not rows:
                self.paths = [None]
                rows = [["No person notes found",
                         "checked: " + ", ".join(roots(self.window))]]
        sublime.set_timeout(
            lambda: self.window.show_quick_panel(rows, self.open_at), 0)

    def open_at(self, index):
        if 0 <= index < len(self.paths) and self.paths[index]:
            self.window.open_file(self.paths[index])
