"""EmailBatch -- review a staged Outlook export and file what you keep.

The Sublime third of the email batch workflow:

  1. Outlook: `ExtractEmailsBatch` writes every mail in a date range to a
     staging folder OUTSIDE the vault (one folder per run) and calls
     `email_batch_prepare` with that folder.
  2. Vault:   `<vault>/scripts/propose_links.py prepare` drops what is already
     filed, was discarded before or comes from a muted sender, stamps
     from_wiki:/to_wiki: on the staged copies and writes `proposals.json`.
     `... agent` then asks a headless `claude -p` about the rest, in the
     background, while review is already under way.
  3. Here:    one quick panel per email, the note previewed behind it.
     TAB ADDS the highlighted row to the email -- a link, a flag -- and the
     panel stays; ENTER FILES it, with what was added plus the row it is on
     when that row is a hub. The proposals are the first rows, so Enter alone
     accepts a good one; every hub is listed underneath, so typing a few
     letters finds another. The list opens on activities; "/all" widens it
     to every hub and person. ctrl+d discards the email in hand.

**An email reaches the vault only when you accept it here.** Accepting writes
the staged note into `<vault>/emails/` under the macro's own filename, with
`[[link]]` on the first body line -- the place and the shape the links you add
by hand already have -- and any flag you toggled set in the frontmatter. That
is the same edit you would otherwise make after `ExtractEmail`, which is why a
tool is allowed to make it: the choice of link is still yours.

Why a quick panel and not keys bound in a view. MarkdownEditing loads after
this package and wins every key overlap (see CLAUDE.md on F12 and tab), so a
bare `enter`/`d`/`1` binding in a markdown view would have to out-select it
binding by binding. A panel has no keymap to fight, fuzzy-filters natively,
and makes "replace the proposal" the same gesture as picking any link.

Volume shapes the rest. About 80 mails arrive a day and about 5 are filed, so
19 in 20 staged emails end up discarded and the review is only fast if THAT
is cheap. Hence three things:

  - the SWEEP: everything no hub was found for is shown as one list, sorted by
    sender. Its first row discards them all; selecting an email instead pulls
    it out to be filed. Only discards get a bulk path -- a wrong discard
    leaves an email unfiled, which is what happens to most mail anyway, while
    a wrong accept puts a false link in the graph -- so accepts are always
    one at a time.
  - "discard all from this sender" and "discard the rest of this thread".
  - "discard and mute": the sender's address goes into `muted_senders.txt`
    and their mail is dropped by `prepare` from then on.

Every decision is undoable until the batch is cleared, because nothing is
deleted before that: an accepted note's staged copy moves to `filed/`, a
discarded one to `discarded/`, and "Undo" is a row in every panel.
`decisions.jsonl`, beside the batch folders, records each one with what was
proposed and by which signal -- the record `propose_links.py evaluate` cannot
give, since it only sees what ended up linked.

A link you accept is carried to the rest of its thread: later emails in the
same conversation offer it first.

Commands:
  email_batch_prepare {"dir": ...}   what the macro calls; ApplicationCommand
                                     because `subl --command` lands in
                                     whichever window is active
  email_batch_review                 "Email batch: Review" -- start, or resume
                                     a paused review (Esc pauses)

Nothing here runs on save or on a timer, and one thing runs on load:
`plugin_loaded` picks up a batch the macro staged in the last few minutes that
has not been through `prepare`. That is the macro's own request arriving late,
not a sweep. When Sublime is not running, `subl --command` starts it and the
command is delivered before any plugin exists to receive it -- the first live
batch, 20261114, was dropped exactly that way: 76 notes staged, Sublime in
front, and nothing else. Only a fresh AND unprepared batch qualifies, so an
ordinary start, a plugin reload or a review you paused never opens a panel by
itself; those wait for "Email batch: Review".

Python 3.3 plugin host: no f-strings, no pathlib; ValueError rather than
json.JSONDecodeError.
"""
import io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time

import sublime
import sublime_plugin

SETTINGS = "EmailBatch.sublime-settings"
PROPOSALS = "proposals.json"
AGENT_PROPOSALS = "proposals_agent.json"
DECISIONS = "decisions.jsonl"
MUTED = "muted_senders.txt"
FILED_DIR = "filed"
DISCARDED_DIR = "discarded"

DEFAULT_PYTHON = os.path.join(os.path.expanduser("~"), ".local", "bin",
                              "python3.14.exe")
DEFAULT_ROOT = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "MarkdownNoteHarness", "emailBatch")

# Proposal tiers good enough to put an email in front of you on its own. The
# sender shortlist is not one of them: its first guess is right well under
# half the time, so on its own it only earns a place in the sweep.
CONFIDENT_TIERS = ("thread", "subject", "agent")
BATCH_NAME_RE = re.compile(r"^\d{8}_\d{4,6}$")
FLAGS = (("toreply", "true", "to reply"),
         ("priority", "high", "high priority"),
         ("personal", "true", "personal"))
WAIT_POLL_MS = 800
# Classes offered as a "/name" row that narrows the per-email panel to that
# class, for browsing what there is. Activities only: they are where four in
# five email links go, and the only list long enough to need reminding of.
FILTER_CLASSES = ("activity",)
# What the hub list shows when a panel opens. None would mean everything.
DEFAULT_FILTER = "activity"
# Shown beside the Discard row and in the placeholder; the binding itself is in
# Default.sublime-keymap, and the two have to be changed together.
DISCARD_KEY_LABEL = "ctrl+d"

_session = None
_busy = {"prepare": False}


# ------------------------------------------------------------------
# Pure helpers -- no sublime, exercised by tests/test_emailbatch.py
# ------------------------------------------------------------------

def file_text(text, links, flags):
    """A staged note's text as it should be filed.

    `links` go on the first body line as `[[a]] [[b]]`, replacing the blank
    line the macro leaves under the closing `---` -- which is where the caret
    lands after ExtractEmail and so where every hand-added link already sits.
    `flags` maps toreply/priority/personal to truthy; a set flag is written
    with the exact value the syntax pins (`true`, `high`), into the empty key
    the macro always writes.

    Line terminators and the BOM are the text's own: the caller reads and
    writes with newline="" and plain utf-8, so neither can change in transit.
    Returns None when the text has no frontmatter to anchor on.
    """
    lines = text.splitlines(True)
    if not lines or lines[0].lstrip(u"﻿").strip() != "---":
        return None
    close = None
    for i in range(1, len(lines)):
        if lines[i].strip() in ("---", "..."):
            close = i
            break
    if close is None:
        return None
    eol = "\r\n" if lines[0].endswith("\r\n") else "\n"

    for key, value, _ in FLAGS:
        if not flags.get(key):
            continue
        for i in range(1, close):
            if re.match(r"^{0}[ \t]*:".format(key), lines[i]):
                lines[i] = "{0}: {1}{2}".format(key, value, eol)
                break
        else:
            lines.insert(1, "{0}: {1}{2}".format(key, value, eol))
            close += 1

    if links:
        line = " ".join("[[{0}]]".format(name) for name in links) + eol
        after = close + 1
        if after < len(lines) and not lines[after].strip():
            lines[after] = line
        else:
            if not lines[close].endswith(("\n", "\r")):
                lines[close] += eol
            lines.insert(after, line)
    return "".join(lines)


def unique_target(folder, name, exists=os.path.exists):
    """`folder/name`, or `folder/stem (2).md` and so on if that is taken.

    A basename is a wiki link target and must be unique in the vault; prepare
    has already dropped anything that IS the existing note, so a clash here
    is a different mail with the same minute, sender and subject.
    """
    stem, ext = os.path.splitext(name)
    candidate = name
    n = 1
    while exists(os.path.join(folder, candidate)):
        n += 1
        candidate = "{0} ({1}){2}".format(stem, n, ext)
    return os.path.join(folder, candidate)


def merge_proposals(base, agent, inherited, tier_order):
    """One list, each hub once, best evidence first.

    `inherited` is what you filed the same thread under a moment ago; nothing
    outranks a decision you just made.
    """
    rank = dict((tier, i) for i, tier in enumerate(tier_order))
    rows = [{"hub": hub, "tier": "thread", "why": "where you just filed this thread"}
            for hub in (inherited or [])]
    pool = list(agent or []) + list(base or [])
    pool.sort(key=lambda p: rank.get(p.get("tier"), len(rank)))
    seen = set()
    out = []
    for p in rows + pool:
        if p["hub"] not in seen:
            seen.add(p["hub"])
            out.append(p)
    return out


def classify(proposals, personal_hint, rescued, answered, agent_pending):
    """Where an email goes next: "item", "wait" or "sweep".

    item   it has a proposal worth your attention, or you asked for it
    wait   the agent has not read it yet and is still running
    sweep  nothing confident: listed with the rest for one bulk decision
    """
    if rescued or personal_hint:
        return "item"
    if any(p["tier"] in CONFIDENT_TIERS for p in proposals):
        return "item"
    if answered:                    # the agent read it and found no hub
        return "sweep"
    if agent_pending:
        return "wait"
    return "item" if proposals else "sweep"


def is_batch_dir(path, root):
    """True only for `<root>/<yyyymmdd_hhmmss>` -- the one thing ever rmtree'd."""
    path = os.path.normcase(os.path.abspath(path))
    root = os.path.normcase(os.path.abspath(root))
    return (os.path.dirname(path) == root
            and bool(BATCH_NAME_RE.match(os.path.basename(path))))


def read_json(path):
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (IOError, OSError, ValueError):
        return None


def pending_files(batch, data):
    """Staged notes still awaiting a decision, in review order."""
    return [fn for fn in data.get("order", [])
            if data["notes"].get(fn, {}).get("status") == "new"
            and os.path.isfile(os.path.join(batch, fn))]


def clear_finished(root):
    """Remove batch folders that were fully decided but never cleared.

    Only reached when no review is in memory -- a fresh Sublime session -- so
    there is no undo stack left that their filed/ and discarded/ copies could
    serve. Returns the folders removed.
    """
    removed = []
    if not os.path.isdir(root):
        return removed
    for name in sorted(os.listdir(root)):
        path = os.path.join(root, name)
        if not is_batch_dir(path, root):
            continue
        data = read_json(os.path.join(path, PROPOSALS))
        if data is not None and not pending_files(path, data):
            shutil.rmtree(path, ignore_errors=True)
            removed.append(path)
    return removed


def next_batch(root):
    """(path, data-or-None) of the oldest batch with something left to do.

    data is None when the batch has not been through `prepare` yet. A batch
    with proposals and nothing pending is finished and is skipped.
    """
    if not os.path.isdir(root):
        return None, None
    for name in sorted(os.listdir(root)):
        path = os.path.join(root, name)
        if not is_batch_dir(path, root):
            continue
        data = read_json(os.path.join(path, PROPOSALS))
        if data is None:
            if any(f.lower().endswith(".md") for f in os.listdir(path)):
                return path, None
            continue
        if pending_files(path, data):
            return path, data
    return None, None


# ------------------------------------------------------------------
# Settings, the vault, the script
# ------------------------------------------------------------------

def settings():
    return sublime.load_settings(SETTINGS)


def expand(path):
    return os.path.normpath(os.path.expandvars(os.path.expanduser(path)))


def staging_root():
    return expand(settings().get("staging_root") or DEFAULT_ROOT)


def find_vault():
    """(vault, window): a folder open in some window that holds the script."""
    for window in sublime.windows():
        for folder in window.folders():
            if os.path.isfile(os.path.join(folder, "scripts", "propose_links.py")):
                return folder, window
    return None, None


def run_script(vault, args, timeout=None):
    """(returncode, output) -- never raises."""
    cmd = [expand(settings().get("python") or DEFAULT_PYTHON),
           os.path.join(vault, "scripts", "propose_links.py"),
           "--vault", vault] + args
    startupinfo = None
    if hasattr(subprocess, "STARTUPINFO"):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    try:
        proc = subprocess.Popen(cmd, cwd=vault, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, stdin=subprocess.PIPE,
                                startupinfo=startupinfo)
        if timeout:
            out, _ = proc.communicate(timeout=timeout)
        else:
            out, _ = proc.communicate()
        return proc.returncode, out.decode("utf-8", "replace").strip()
    except Exception as e:      # missing interpreter, timeout, anything
        return -1, "could not run {0}: {1}".format(cmd[0], e)


def say(message):
    sublime.status_message("Email batch: " + message)
    print("EmailBatch: " + message)


def kind(const, char, label):
    return (getattr(sublime, const, 0), char, label)


def panel_item(trigger, details="", annotation="", item_kind=None):
    cls = getattr(sublime, "QuickPanelItem", None)
    if cls is None:             # only under the test stubs
        return [trigger, details or annotation]
    return cls(trigger, details, annotation,
               item_kind or getattr(sublime, "KIND_AMBIGUOUS", (0, "", "")))


KIND_PROPOSED = ("KIND_ID_COLOR_GREENISH", u"→", "Proposed")
KIND_FILE = ("KIND_ID_COLOR_BLUISH", u"✓", "File")
KIND_DISCARD = ("KIND_ID_COLOR_REDISH", u"×", "Discard")
KIND_OPTION = ("KIND_ID_COLOR_YELLOWISH", u"·", "Option")
KIND_UNDO = ("KIND_ID_COLOR_ORANGISH", u"↶", "Undo")


# ------------------------------------------------------------------
# The review
# ------------------------------------------------------------------

class Session(object):
    """One batch under review. Queue state is the folder; this holds the rest."""

    def __init__(self, window, batch, data):
        self.window = window
        self.batch = batch
        self.data = data
        self.notes = data["notes"]
        self.vault = data["vault"]
        self.left = pending_files(batch, data)
        self.total = len(self.left)
        self.agent = {}
        self.agent_stamp = None
        self.agent_running = False
        self.agent_skip = False
        self.inherited = {}         # thread key -> links you filed it under
        self.rescued = set()
        self.undo_stack = []
        self.filed = 0
        self.discarded = 0
        self.item = None
        self.token = 0
        self.dismissed = None       # token of a panel we closed ourselves
        self.panel = None           # the per-email panel while it is open

    # ---- state ---------------------------------------------------

    def load_agent(self):
        path = os.path.join(self.batch, AGENT_PROPOSALS)
        try:
            stamp = os.stat(path).st_mtime
        except OSError:
            return
        if stamp == self.agent_stamp:
            return
        data = read_json(path)
        if data is not None:        # None: caught mid-replace; next poll gets it
            self.agent_stamp = stamp
            self.agent = data.get("notes") or {}

    def proposals(self, fn):
        note = self.notes[fn]
        return merge_proposals(note.get("proposals"), self.agent.get(fn),
                               self.inherited.get(note.get("thread") or None),
                               self.data.get("tier_order", []))

    def kind_of(self, fn):
        note = self.notes[fn]
        return classify(self.proposals(fn), note.get("personal_hint"),
                        fn in self.rescued, fn in self.agent,
                        self.agent_running and not self.agent_skip)

    def same(self, fn, key):
        """Other pending emails sharing this one's sender address or thread."""
        value = self.notes[fn].get(key)
        if not value:
            return [fn]
        return [f for f in self.left if self.notes[f].get(key) == value]

    # ---- flow ----------------------------------------------------

    def later(self, fn):
        # Let the panel that just closed finish closing before the next opens.
        sublime.set_timeout(fn, 0)

    def advance(self):
        self.load_agent()
        waiting = False
        for fn in self.left:
            where = self.kind_of(fn)
            if where == "item":
                return self.show_item(fn)
            if where == "wait":
                waiting = True
        if waiting:
            return self.show_wait()
        if self.left:
            return self.show_sweep()
        self.show_done()

    def show(self, rows, on_select, selected=0, on_highlight=None, placeholder=""):
        self.token += 1
        token = self.token
        self.panel = None           # show_item sets it again for its own panel

        def selected_cb(index):
            # A panel we closed ourselves reports -1 like an Esc does, and may
            # report it after the next panel is already up. Both cases are
            # recognised by token, so neither can be mistaken for a pause and
            # neither leaves a flag behind to swallow a real keypress.
            if token != self.token or token == self.dismissed:
                return
            self.panel = None
            if index < 0:
                return self.pause()
            on_select(index)

        def highlight_cb(index):
            if token == self.token and on_highlight:
                on_highlight(index)

        # KEEP_OPEN_ON_FOCUS_LOST: the preview behind the panel finishes loading
        # after the panel is up, and a panel that closed whenever focus moved
        # would read that as an Esc and pause the review.
        self.window.show_quick_panel(
            rows, selected_cb, getattr(sublime, "KEEP_OPEN_ON_FOCUS_LOST", 0),
            selected, highlight_cb, placeholder)

    def pause(self):
        self.close_preview()
        say("paused with {0} left -- 'Email batch: Review' resumes".format(
            len(self.left)))

    # ---- one email -----------------------------------------------

    def show_item(self, fn, keep=False):
        """The panel for one email. TAB ADDS, ENTER FILES.

        Tab on a hub adds its link, on a flag toggles it, on "/all" widens
        the list, on "Remove" takes a link back -- and the panel comes
        straight back for the same email. Enter writes the note: with what
        has been added so far, plus the row it is on when that row is a hub.
        So Enter alone on a good proposal is still one key, and typing a name
        and pressing Enter cannot file the email with no link by accident.
        After a Tab the selection is on the "File under ..." row, where Enter
        files exactly what was added. On the rows that are not something to
        file under (flags, /activity, Remove) Enter does what Tab does.
        """
        note = self.notes[fn]
        if not keep or not self.item or self.item["fn"] != fn:
            self.item = {"fn": fn, "links": [], "only": DEFAULT_FILTER,
                         "flags": {"personal": bool(note.get("personal_hint"))}}
            self.preview(fn)
        item = self.item
        rows = []
        actions = []

        def add(action, arg, trigger, details="", annotation="", item_kind=None):
            rows.append(panel_item(trigger, details, annotation,
                                   kind(*item_kind) if item_kind else None))
            actions.append((action, arg))

        chosen = item["links"]
        only = item["only"]
        flagged = [label for key, _, label in FLAGS if item["flags"].get(key)]
        suffix = (u" · " + u", ".join(flagged)) if flagged else u""
        nolink = u"File with no link" + suffix

        # What Enter files when nothing more is picked. It leads only once
        # there is something deliberate to file -- a link added with Tab, or a
        # sender whose mail is personal -- so that with nothing added yet the
        # first row is still the proposal.
        if chosen:
            add("file", None, u"File under " + u" ".join(
                u"[[{0}]]".format(h) for h in chosen) + suffix,
                annotation="enter", item_kind=KIND_FILE)
            for hub in chosen:
                add("remove", hub, u"Remove [[{0}]]".format(hub), item_kind=KIND_OPTION)
        elif item["flags"].get("personal"):
            add("file", None, nolink, annotation="enter", item_kind=KIND_FILE)

        labels = dict((h["name"], h.get("label") or "") for h in self.data["hubs"])
        proposed = [p for p in self.proposals(fn) if p["hub"] not in chosen]
        for p in proposed:
            add("pick", p["hub"], u"[[{0}]]".format(p["hub"]),
                details=labels.get(p["hub"], ""),
                annotation=u"{0} · {1}".format(p["tier"], p.get("why", "")),
                item_kind=KIND_PROPOSED)
        if not chosen and not item["flags"].get("personal"):
            add("file", None, nolink, item_kind=KIND_FILE)
        add("discard", [fn], u"Discard", annotation=DISCARD_KEY_LABEL,
            item_kind=KIND_DISCARD)
        by_sender = self.same(fn, "from_email")
        if len(by_sender) > 1:
            add("discard", by_sender, u"Discard all {0} from {1}".format(
                len(by_sender), note.get("from") or note.get("from_email")),
                item_kind=KIND_DISCARD)
        by_thread = self.same(fn, "thread")
        if len(by_thread) > 1:
            add("discard", by_thread, u"Discard the rest of this thread ({0})".format(
                len(by_thread)), item_kind=KIND_DISCARD)
        if note.get("from_email") and not note.get("sent"):
            add("mute", None, u"Discard and mute " + note["from_email"],
                annotation="never staged again", item_kind=KIND_DISCARD)
        for key, _, label in FLAGS:
            add("flag", key, u"Flag: {0}".format(label),
                annotation="on" if item["flags"].get(key) else "tab",
                item_kind=KIND_OPTION)
        # The list below holds one class at a time and opens on DEFAULT_FILTER.
        # Each other view is a row: "/all" for every hub and person, and a
        # "/class" row for each class in FILTER_CLASSES not being shown.
        if only:
            add("only", None, u"/all", annotation="every hub and person",
                item_kind=KIND_OPTION)
        for cls in FILTER_CLASSES:
            if cls != only:
                add("only", cls, u"/" + cls, annotation="only these",
                    item_kind=KIND_OPTION)
        if self.undo_stack:
            add("undo", None, u"Undo: " + self.undo_stack[-1]["label"],
                item_kind=KIND_UNDO)
        # A proposal is listed above whatever its class; below, the filter
        # decides. Under a filter each hub shows its title, since the point of
        # narrowing is to read what there is.
        skip = set(chosen) | set(p["hub"] for p in proposed)
        for hub in self.data["hubs"]:
            if hub["name"] in skip or (only and hub["class"] != only):
                continue
            add("pick", hub["name"], hub["name"],
                details=(hub.get("label") or "") if only else "",
                annotation=hub["class"])

        done = self.total - len(self.left) + 1
        placeholder = u"{0}/{1} · {2} · {3} — {4}   [tab adds · enter files · {5} discards]".format(
            done, self.total, note.get("date", ""), note.get("from", ""),
            note.get("subject", ""), DISCARD_KEY_LABEL)

        def on_select(index):
            action, arg = actions[index]
            if action == "file":
                return self.accept(fn, list(chosen), item["flags"])
            if action == "pick":
                return self.accept(fn, list(chosen) + [arg], item["flags"])
            if action == "discard":
                return self.discard(arg, "item")
            if action == "mute":
                return self.mute(fn)
            if action == "undo":
                return self.undo()
            self.adjust(action, arg)
            self.later(lambda: self.show_item(fn, keep=True))

        def on_highlight(index):
            if self.panel and 0 <= index < len(actions):
                self.panel["index"] = index

        self.show(rows, on_select, 0, on_highlight, placeholder)
        self.panel = {"fn": fn, "actions": actions, "index": 0, "token": self.token}

    def adjust(self, action, arg):
        """Apply a row that changes the email in hand without filing it.
        True when it did."""
        item = self.item
        if action == "pick":
            item["links"].append(arg)
        elif action == "remove":
            item["links"].remove(arg)
        elif action == "flag":
            item["flags"][arg] = not item["flags"].get(arg)
        elif action == "only":
            item["only"] = arg
        else:
            return False
        return True

    def discard_key(self):
        """The discard shortcut in the per-email panel: this email, now."""
        panel = self.panel
        if not panel:
            return
        self.dismissed = panel["token"]
        self.panel = None
        self.window.run_command("hide_overlay")
        self.discard([panel["fn"]], "key")

    def tab(self):
        """Tab in the per-email panel: add the highlighted row and stay."""
        panel = self.panel
        if not panel or not self.item or self.item["fn"] != panel["fn"]:
            return
        action, arg = panel["actions"][panel["index"]]
        if not self.adjust(action, arg):
            return                  # File, Discard, Undo: Enter's business
        fn = panel["fn"]
        self.dismissed = panel["token"]
        self.panel = None
        self.window.run_command("hide_overlay")
        self.later(lambda: self.show_item(fn, keep=True))

    # ---- the sweep -----------------------------------------------

    def show_sweep(self):
        self.item = None
        fns = [fn for fn in self.left if self.kind_of(fn) == "sweep"]
        fns.sort(key=lambda fn: ((self.notes[fn].get("from") or "").lower(),
                                 self.notes[fn].get("date") or ""))
        rows = [panel_item(
            u"Discard all {0}".format(len(fns)),
            u"No hub was found for these. Select one below to file it instead.",
            "", kind(*KIND_DISCARD))]
        for fn in fns:
            note = self.notes[fn]
            rows.append(panel_item(u"{0} — {1}".format(
                note.get("from", ""), note.get("subject", "")),
                "", note.get("date", "")))
        if self.undo_stack:
            rows.append(panel_item(u"Undo: " + self.undo_stack[-1]["label"],
                                   "", "", kind(*KIND_UNDO)))

        def on_select(index):
            if index == 0:
                return self.discard(fns, "sweep")
            if index > len(fns):
                return self.undo()
            self.rescued.add(fns[index - 1])
            self.later(lambda: self.show_item(fns[index - 1]))

        def on_highlight(index):
            if 1 <= index <= len(fns):
                self.preview(fns[index - 1])

        # The first EMAIL is preselected, not the bulk row: an Enter still in
        # flight from the last per-email panel then opens one email instead of
        # discarding several hundred.
        self.show(rows, on_select, 1 if fns else 0, on_highlight,
                  u"{0} left with no hub — up + enter discards them all".format(len(fns)))

    # ---- waiting on the agent ------------------------------------

    def show_wait(self):
        self.item = None
        unread = len([fn for fn in self.left if self.kind_of(fn) == "wait"])
        rows = [panel_item(u"Waiting for the agent",
                           u"{0} emails it has not read yet".format(unread)),
                panel_item(u"Go on without it",
                           u"Whatever it has not read goes to the sweep")]

        def on_select(index):
            if index == 1:
                self.agent_skip = True
            self.later(self.advance)

        self.show(rows, on_select, 0, None, u"the agent is still reading")
        token = self.token

        def poll():
            if token != self.token:
                return
            self.load_agent()
            moved = (not self.agent_running or
                     any(self.kind_of(fn) == "item" for fn in self.left))
            if moved:
                self.dismissed = token
                self.window.run_command("hide_overlay")
                self.later(self.advance)
            else:
                sublime.set_timeout(poll, WAIT_POLL_MS)

        sublime.set_timeout(poll, WAIT_POLL_MS)

    # ---- the end -------------------------------------------------

    def show_done(self):
        self.item = None
        self.close_preview()
        rows = [panel_item(u"Clear this batch",
                           u"{0} filed, {1} discarded".format(self.filed, self.discarded),
                           "", kind(*KIND_FILE))]
        if self.undo_stack:
            rows.append(panel_item(u"Undo: " + self.undo_stack[-1]["label"],
                                   "", "", kind(*KIND_UNDO)))

        def on_select(index):
            if index == 1:
                return self.undo()
            self.clear()

        self.show(rows, on_select, 0, None, u"nothing left in this batch")

    def clear(self):
        global _session
        root = os.path.dirname(self.batch)
        if is_batch_dir(self.batch, root):
            shutil.rmtree(self.batch, ignore_errors=True)
        _session = None
        say("batch cleared: {0} filed, {1} discarded".format(
            self.filed, self.discarded))
        if next_batch(root)[0]:
            sublime.set_timeout(lambda: start(window=self.window), 0)

    # ---- actions -------------------------------------------------

    def log(self, action, fn, **extra):
        note = self.notes[fn]
        row = {"ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
               "batch": os.path.basename(self.batch), "action": action,
               "file": fn, "message_id": note.get("message_id", ""),
               "from": note.get("from_email", ""), "thread": note.get("thread", "")}
        row.update(extra)
        path = os.path.join(os.path.dirname(self.batch), DECISIONS)
        try:
            with io.open(path, "a", encoding="utf-8") as f:
                f.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + u"\n")
        except (IOError, OSError) as e:
            print("EmailBatch: could not write {0}: {1}".format(path, e))

    def shelve(self, fn, folder):
        """Move a staged note out of the queue, into filed/ or discarded/."""
        target = os.path.join(self.batch, folder)
        if not os.path.isdir(target):
            os.makedirs(target)
        shutil.move(os.path.join(self.batch, fn), os.path.join(target, fn))
        if fn in self.left:
            self.left.remove(fn)

    def unshelve(self, fn, folder):
        shutil.move(os.path.join(self.batch, folder, fn),
                    os.path.join(self.batch, fn))
        back = set(self.left) | set([fn])
        self.left = [f for f in self.data["order"] if f in back]

    def accept(self, fn, links, flags):
        note = self.notes[fn]
        src = os.path.join(self.batch, fn)
        try:
            with io.open(src, "r", encoding="utf-8", newline="") as f:
                text = f.read()
            new = file_text(text, links, flags)
            if new is None:
                raise ValueError("no frontmatter")
            folder = os.path.join(self.vault, settings().get("emails_dir") or "emails")
            dest = unique_target(folder, fn)
            with io.open(dest, "w", encoding="utf-8", newline="") as f:
                f.write(new)
            self.shelve(fn, FILED_DIR)
        except Exception as e:
            say("could NOT file {0}: {1}".format(fn, e))
            return
        proposed = [{"hub": p["hub"], "tier": p["tier"]} for p in self.proposals(fn)]
        self.log("accept", fn, links=list(links), proposed=proposed,
                 flags=sorted(k for k in flags if flags[k]))
        if links and note.get("thread"):
            self.inherited[note["thread"]] = list(links)
        self.filed += 1
        self.undo_stack.append({"kind": "accept", "fns": [fn], "dest": dest,
                                "label": u"filing of " + note.get("subject", fn)})
        say(u"filed {0}{1}".format(
            u" ".join(u"[[{0}]]".format(h) for h in links) or u"with no link",
            u" · " + u", ".join(sorted(k for k in flags if flags[k]))
            if any(flags.values()) else u""))
        self.later(self.advance)

    def discard(self, fns, via):
        fns = [fn for fn in fns if fn in self.left]
        done = []
        for fn in fns:
            try:
                self.shelve(fn, DISCARDED_DIR)
            except Exception as e:
                say("could not discard {0}: {1}".format(fn, e))
                continue
            proposed = [{"hub": p["hub"], "tier": p["tier"]} for p in self.proposals(fn)]
            self.log("discard", fn, via=via, proposed=proposed)
            done.append(fn)
        if done:
            self.discarded += len(done)
            label = (u"discard of " + self.notes[done[0]].get("subject", done[0])
                     if len(done) == 1 else u"discard of {0} emails".format(len(done)))
            self.undo_stack.append({"kind": "discard", "fns": done, "label": label})
            say(label.replace("discard of", "discarded", 1))
        self.later(self.advance)

    def mute(self, fn):
        addr = self.notes[fn].get("from_email", "")
        path = os.path.join(os.path.dirname(self.batch), MUTED)
        try:
            with io.open(path, "a", encoding="utf-8") as f:
                f.write(u"{0}   # muted {1}\n".format(addr, time.strftime("%Y%m%d")))
        except (IOError, OSError) as e:
            say("could not write {0}: {1}".format(path, e))
        self.discard(self.same(fn, "from_email"), "mute")

    def undo(self):
        if not self.undo_stack:
            return self.later(self.advance)
        entry = self.undo_stack.pop()
        folder = FILED_DIR if entry["kind"] == "accept" else DISCARDED_DIR
        for fn in entry["fns"]:
            try:
                self.unshelve(fn, folder)
            except Exception as e:
                say("could not undo {0}: {1}".format(fn, e))
                continue
            self.log("undo", fn, undid=entry["kind"])
            # Bring it straight back in front of you rather than wherever its
            # proposals would have queued it.
            self.rescued.add(fn)
        if entry["kind"] == "accept":
            try:
                os.remove(entry["dest"])
            except OSError as e:
                say("could not remove {0}: {1}".format(entry["dest"], e))
            self.filed -= 1
        else:
            self.discarded -= len(entry["fns"])
        say(u"undid the " + entry["label"])
        self.later(self.advance)

    # ---- preview -------------------------------------------------

    def preview(self, fn):
        path = os.path.join(self.batch, fn)
        if not os.path.isfile(path):
            return
        view = self.window.open_file(path, sublime.TRANSIENT)
        # Read-only: a key pressed in the gap between two panels would
        # otherwise be typed into the staged note.
        view.set_read_only(True)
        fold = francis_helper("fold_frontmatter")
        if fold:
            fold(self.window, view)

    def close_preview(self):
        try:
            view = self.window.transient_view_in_group(self.window.active_group())
        except Exception:
            view = None
        name = view.file_name() if view else None
        if name and os.path.normcase(name).startswith(os.path.normcase(self.batch)):
            view.set_scratch(True)
            view.close()


def francis_helper(name):
    """A presentation helper from francis.py, or None.

    Looked up lazily in sys.modules, never imported: the package directory
    has hyphens in its name, and a module-level import would run before
    Sublime has necessarily loaded the sibling. Same approach as WikiLinks.py.
    """
    for key, module in list(sys.modules.items()):
        if key.endswith(".francis") and hasattr(module, name):
            return getattr(module, name)
    return None


# ------------------------------------------------------------------
# Starting and resuming
# ------------------------------------------------------------------

def start(batch=None, window=None):
    """Prepare a batch if it needs it, then review it. Safe to call again."""
    global _session
    vault, vault_window = find_vault()
    if not vault:
        return say("open the vault in Sublime first -- no folder here holds "
                   "scripts/propose_links.py")
    window = vault_window or window or sublime.active_window()
    root = staging_root()

    if batch:
        batch = os.path.normpath(batch)
        data = read_json(os.path.join(batch, PROPOSALS))
    elif _session and os.path.isdir(_session.batch):
        # Paused, or finished but not yet cleared: either way the session
        # knows what to show next, and still holds the undo stack.
        batch, data = _session.batch, _session.data
    else:
        clear_finished(root)
        batch, data = next_batch(root)
    if not batch or not os.path.isdir(batch):
        return say("nothing staged under {0}".format(root))

    if _session and _session.batch == batch:
        _session.window = window
        window.bring_to_front()
        return _session.advance()

    if data is not None:
        return begin(window, vault, batch, data)

    if _busy["prepare"]:
        return say("still preparing the last batch")
    _busy["prepare"] = True
    say("reading the vault… (seconds normally; minutes the first time after a restart)")

    def work():
        code, out = run_script(vault, ["prepare", batch],
                               settings().get("prepare_timeout", 600))
        _busy["prepare"] = False

        def done():
            data = read_json(os.path.join(batch, PROPOSALS))
            if code != 0 or data is None:
                print(out)
                return say("prepare failed -- see the console")
            say(out.splitlines()[-1] if out else "prepared")
            begin(window, vault, batch, data)

        sublime.set_timeout(done, 0)

    threading.Thread(target=work).start()


def begin(window, vault, batch, data):
    global _session
    session = Session(window, batch, data)
    if not session.left:
        if is_batch_dir(batch, os.path.dirname(batch)):
            shutil.rmtree(batch, ignore_errors=True)
        return say("nothing new in this batch -- cleared")
    _session = session
    window.bring_to_front()

    agent_file = read_json(os.path.join(batch, AGENT_PROPOSALS))
    if settings().get("agent", True) and not (agent_file and agent_file.get("done")):
        session.agent_running = True
        args = ["agent", batch, "--model", settings().get("agent_model") or "sonnet"]
        if settings().get("claude"):
            args += ["--claude", expand(settings().get("claude"))]

        def work():
            code, out = run_script(vault, args)
            session.agent_running = False
            if code != 0:
                print("EmailBatch agent: " + out)
                sublime.set_timeout(
                    lambda: say("the agent stopped early -- see the console; "
                                "review carries on without it"), 0)

        threading.Thread(target=work).start()
    session.advance()


# A batch older than this was not staged by the macro run that just started
# Sublime, so it is left for the palette command.
HANDOFF_WINDOW_SECONDS = 600
# Sublime restores its windows and their folders after plugins load; the vault
# may not be findable on the first look.
HANDOFF_FIRST_MS = 3000
HANDOFF_RETRY_MS = 3000
HANDOFF_TRIES = 6


def fresh_unprepared_batch(root, now=None):
    """The batch a just-run macro is waiting on, or None.

    Unprepared (no proposals.json) and created within HANDOFF_WINDOW_SECONDS.
    The age comes from the folder's own name, which the macro stamps, rather
    than from an mtime that OneDrive or a virus scanner may have touched.
    """
    batch, data = next_batch(root)
    if not batch or data is not None:
        return None
    try:
        made = time.mktime(time.strptime(os.path.basename(batch)[:15],
                                         "%Y%m%d_%H%M%S"))
    except ValueError:
        return None
    age = (now if now is not None else time.time()) - made
    return batch if -60 <= age <= HANDOFF_WINDOW_SECONDS else None


def pick_up_handoff(tries=HANDOFF_TRIES):
    if _session or _busy["prepare"]:
        return                      # the command did arrive after all
    batch = fresh_unprepared_batch(staging_root())
    if not batch:
        return
    if not find_vault()[0] and tries > 1:
        sublime.set_timeout(lambda: pick_up_handoff(tries - 1), HANDOFF_RETRY_MS)
        return
    start(batch)


def plugin_loaded():
    sublime.set_timeout(pick_up_handoff, HANDOFF_FIRST_MS)


class EmailBatchTabCommand(sublime_plugin.WindowCommand):
    """Tab while the per-email panel is open. Bound in Default.sublime-keymap
    on the `email_batch_panel` context below, so it exists nowhere else."""

    def run(self):
        if _session:
            _session.tab()


class EmailBatchDiscardCommand(sublime_plugin.WindowCommand):
    """The discard shortcut while the per-email panel is open. Same gate as
    email_batch_tab, so it too exists nowhere else."""

    def run(self):
        if _session:
            _session.discard_key()


class EmailBatchContext(sublime_plugin.EventListener):
    """`email_batch_panel`: true only while a per-email panel is up.

    A quick panel has no scope a keymap could select on, so the binding asks
    this instead. It is what keeps Tab untouched in every other overlay --
    the command palette, Goto Anything, the wiki link picker, and this
    module's own sweep and wait panels.
    """

    def on_query_context(self, view, key, operator, operand, match_all):
        if key != "email_batch_panel":
            return None
        is_open = bool(_session and _session.panel)
        if operator == getattr(sublime, "OP_NOT_EQUAL", 1):
            return is_open != bool(operand)
        return is_open == bool(operand)


class EmailBatchPrepareCommand(sublime_plugin.ApplicationCommand):
    """What ExtractEmailsBatch calls: subl --command 'email_batch_prepare {"dir": ...}'"""

    def run(self, dir=None):
        start(dir)


class EmailBatchReviewCommand(sublime_plugin.WindowCommand):
    """Start on the oldest staged batch, or resume the one that was paused."""

    def run(self, dir=None):
        start(dir, self.window)
