"""WikiLinks.py -- backlinks and rename, answered from Sublime's symbol index.

MarkdownEditing's `mde_list_back_links` walks every project folder and READS
every .md it finds, substring-searching for "[[<basename>]]". On this vault that
is ~2,400 files per invocation, over OneDrive, with no cache. Sublime already
holds the answer: wiki link targets are scoped -- by the built-in Markdown
syntax in the body, by YAMLEmailFrontmatter.sublime-syntax in the frontmatter --
and indexed as REFERENCES by the two "Reference List - Wiki *.tmPreferences"
files. "Who links to X?" is then one native lookup.

What makes this work at all is that a wiki link target IS a filename: the
vault's rule is a globally unique basename with no path. So the name to query is
never something that must be discovered; it is sitting in the directory listing.

Three commands:

  wiki_backlinks     every reference to the note you are in, in a quick panel
                     with previews. Where the caret sits does not change the
                     answer; under_caret=True asks about the link it is in.

  wiki_rename_note   rename a note and rewrite every link to it. Locations come
                     from the index, so nothing is searched for.

  wiki_pick_link     the same observation run backwards -- not "who links to
                     X?" but "what is there to link to?". Tab inside [] doubles
                     the bracket and offers every CLASSIFIED note, narrowable by
                     /person, /activity, /email ... Those class tokens are
                     already in the index because the frontmatter "class: X" is
                     scoped, so the candidate list costs one symbol_locations
                     call per class and no file is opened. MarkdownEditing's
                     mde_make_page_reference answers this by stat-ing every file
                     in every project folder, which is the design this module
                     exists to stop repeating.

Deliberately NOT here: the whole-vault adjacency graph. The index answers
per-name queries and cannot be enumerated -- every window-level API takes a
symbol string, and only View.indexed_references() lists anything, which would
mean opening all ~2,400 notes as views. The graph belongs to a vault-side pass,
for the same reason wikiresolve_core.py does.

THREE THINGS THAT LOOK LIKE PARANOIA AND ARE NOT:

  * Only rows the index reported are ever rewritten. A raw 'to:' line can hold
    a bare "[[francis]]" too (2,286 of them do), and those keys are the Outlook
    macro's audit trail, never edited from this side. Rewriting by line scan
    rather than by reported row would quietly break that contract.

  * The row base is measured, not assumed. SymbolLocation.row feeds
    path_encoded_position(), which is 1-based, but the API reference does not
    promise it and a silent off-by-one would rewrite the line above every link
    in the vault.

  * A rename refuses to run while any affected file has unsaved changes,
    because the rewrite goes to disk underneath open buffers.

ONE BLIND SPOT, AND IT IS STATED OUT LOUD ON EVERY RENAME: a wiki link written
inside a markdown HEADING or inside an open TODO is not indexed as a reference,
so it is neither listed by backlinks nor rewritten by a rename. That is not an
oversight; it is the only way to keep those two things in the symbol lists they
belong to. At any point only the highest scoring .tmPreferences entry applies,
and the wiki entry outscores both containers, so letting it win cut every such
heading and todo in half (the reported symptom was a project symbol called
"]]"). The mechanism, the measurement (219 headings / 477 links, 304 todos /
468 links, about 700 of them pointing at a note that exists) and the experiment
that would buy them back are written up in
"Reference List - Wiki Body.tmPreferences". SKIPPED_NOTE below is what carries
the fact into the dialog and the log.
"""

import sublime
import sublime_plugin

import datetime
import fnmatch
import io
import os
import sys

# Scopes carrying a wiki link target. The first comes free from the built-in
# Markdown syntax; the second is this package's, matched by prefix so it covers
# all five relations (from / to / organizer / attendees / responsible).
#
# BODY_SELECTOR is the BARE meta scope, deliberately unlike the selector in
# "Reference List - Wiki Body.tmPreferences", which subtracts both brackets and
# the two containers. Here the subtractions would be wrong twice over: this is
# what expand_to_scope() widens by, so it has to cover the whole link rather than
# stop at a bracket, and the caret should still find its name inside a heading or
# a todo even though the index has no row for it. The two brackets that come with
# that are stripped in name_under_caret(), which is the only place they matter.
BODY_SELECTOR = "meta.link.reference.wiki.description.markdown"
FRONTMATTER_SELECTOR = "entity.name.wikilink"

# Above this many files the confirmation says so loudly. Renaming the vault
# owner's note is a ~1,700 file rewrite; that should never be a surprise.
RENAME_WARN_FILES = 200

# Said on every rename, in the dialog and in the log. A blind spot nobody is
# told about is the failure this package treats as worse than an error, and a
# rename is the one moment where staying quiet about it could cost a link.
SKIPPED_NOTE = ("Wiki links inside markdown headings and inside open todos are "
                "not indexed, so they are NOT rewritten.")

SKIP_DIRS = (".git", "__pycache__", "node_modules")


# ------------------------------------------------------------------
# Logging and module lookup
# ------------------------------------------------------------------

def log(message):
    """Append to Data/Logs/WikiLinks.log.

    A rename touching four figures of files should leave a record that outlives
    the status bar.
    """
    try:
        log_dir = os.path.join(os.path.dirname(sublime.packages_path()), "Logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with io.open(os.path.join(log_dir, "WikiLinks.log"), "a",
                     encoding="utf-8") as f:
            f.write("[{0}] {1}\n".format(stamp, message))
    except Exception as e:
        print("WikiLinks: could not write log: {0}".format(e))


def sibling(name):
    """Another plugin module from this package, or None.

    The package directory is 'Markdown-Note-Harness'; the hyphens mean it cannot
    be imported by name, and a relative import at module level would run before
    Sublime has necessarily loaded the sibling. Looked up lazily instead: by the
    time a command runs, every root .py in the package is in sys.modules.
    Everything reached this way is presentation only, so None is survivable.
    """
    if not __package__:
        return None
    return sys.modules.get("{0}.{1}".format(__package__, name))


# ------------------------------------------------------------------
# Pure helpers -- all of this is covered by tests/test_wikilinks.py
# ------------------------------------------------------------------

def note_name(path):
    """The wiki name of a note: its basename with the extension dropped."""
    return os.path.splitext(os.path.basename(path))[0]


def link_text(name):
    """The literal a link to `name` is written as.

    Brackets included on purpose: replacing "[[francis]]" then cannot touch
    "[[francisDennig]]", which a bare-name replacement would corrupt.
    """
    return "[[{0}]]".format(name)


def read_text(path):
    """newline="" so the file's own terminators arrive intact, and plain utf-8
    so a BOM stays as the first character of line 1 rather than being decoded
    away. Both matter here: these files are CRLF with a BOM, and a rename must
    not be what rewrites 1,700 of them end to end."""
    with io.open(path, "r", encoding="utf-8", newline="") as f:
        return f.read()


def write_text(path, text):
    """utf-8, never utf-8-sig: the BOM travels as part of the text, so encoding
    one again here would add it to files that never had it."""
    with io.open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def detect_row_base(rows, lines, old):
    """Whether SymbolLocation.row counts from 1 or from 0, measured on the file.

    path_encoded_position() feeds row to ENCODED_POSITION, which is 1-based, so
    1 is overwhelmingly likely. But the API reference does not say so, and the
    cost of being wrong is rewriting the line above every link in the vault, so
    it is checked rather than trusted: under the correct base, every reported
    row lands on a line that actually holds the link.

    Returns 1, 0, or None when neither fits -- which means the file has moved on
    since it was indexed, and the caller must skip rather than guess.
    """
    needle = link_text(old)
    present = set(i for i, line in enumerate(lines) if needle in line)
    for base in (1, 0):
        if all((row - base) in present for row in rows):
            return base
    return None


def rewrite_lines(lines, rows, old, new, base):
    """Replace the link on each reported row, in place. Returns (applied, skipped).

    Only reported rows are touched: a line holding "[[old]]" that the index did
    not report is a raw macro-owned key, and rewriting it would break the
    contract that says those are never edited from this side.

    Note there is no descending-order dance. Replacing inside a line never moves
    another line, so the offset shifting that character-level edits have to sort
    around does not arise at all.
    """
    old_link, new_link = link_text(old), link_text(new)
    applied, skipped = 0, []
    for row in sorted(set(rows)):
        i = row - base
        if i < 0 or i >= len(lines) or old_link not in lines[i]:
            skipped.append(row)
            continue
        applied += lines[i].count(old_link)
        lines[i] = lines[i].replace(old_link, new_link)
    return applied, skipped


def rewrite_file(path, rows, old, new):
    """Apply a rename to one file. Returns (applied, skipped_rows, error)."""
    try:
        lines = read_text(path).split("\n")
    except (OSError, UnicodeDecodeError) as e:
        return 0, list(rows), "unreadable: {0}".format(e)

    base = detect_row_base(rows, lines, old)
    if base is None:
        return 0, list(rows), "no reported row holds {0}".format(link_text(old))

    applied, skipped = rewrite_lines(lines, rows, old, new, base)
    if applied:
        try:
            write_text(path, "\n".join(lines))
        except OSError as e:
            return 0, list(rows), "unwritable: {0}".format(e)
    return applied, skipped, None


def basename_index(roots):
    """{wiki name: [paths]} across the roots. Names only; no file is read.

    This is why a rename can be validated cheaply: the vault's invariant is one
    file per basename, and checking it costs a directory walk.
    """
    index = {}
    for root in roots:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for filename in filenames:
                if filename.lower().endswith(".md"):
                    index.setdefault(note_name(filename), []).append(
                        os.path.join(dirpath, filename))
    return index


def group_rows(locations):
    """{path: [rows]} from a list of SymbolLocation."""
    grouped = {}
    for loc in locations:
        grouped.setdefault(loc.path, []).append(loc.row)
    return grouped


def bracket_span(line_text, col):
    """(start, end, seed) of the bracket pair around `col`, or None.

    start and end are inclusive of the brackets and seed is whatever sits
    between them, so "[]" at col 1 gives (0, 2, "") and "[[x]]" gives
    (0, 5, "x"). One function for both cases on purpose: widening by one
    character on each side when the neighbour is another bracket is the ONLY
    thing that separates a shorthand reference from a wiki link, and the picker
    treats them the same -- it rewrites the pair either way.

    None when the caret is not inside a pair on this line, or when the text
    between the brackets holds a bracket of its own. That second case is a
    nested or unbalanced construct, and a picker that guessed at it would
    rewrite the wrong span of a line it did not understand.
    """
    open_i = line_text.rfind("[", 0, col)
    close_i = line_text.find("]", col)
    if open_i < 0 or close_i < 0:
        return None

    seed = line_text[open_i + 1:close_i]
    if "[" in seed or "]" in seed:
        return None

    start = open_i - 1 if open_i > 0 and line_text[open_i - 1] == "[" else open_i
    end = close_i + 1
    if end < len(line_text) and line_text[end] == "]":
        end += 1
    return (start, end, seed)


# ------------------------------------------------------------------
# Index queries
# ------------------------------------------------------------------

def wiki_references(window, name):
    """Every indexed reference to `name`.

    SYMBOL_SOURCE_ANY rather than SYMBOL_SOURCE_INDEX, so a note that is open
    and freshly edited counts alongside what is on disk.
    """
    return window.symbol_locations(
        name, sublime.SYMBOL_SOURCE_ANY, sublime.SYMBOL_TYPE_REFERENCE)


def name_under_caret(view):
    """The wiki name the caret sits in, or None."""
    for sel in view.sel():
        pt = sel.begin()
        for selector in (FRONTMATTER_SELECTOR, BODY_SELECTOR):
            if not view.match_selector(pt, selector):
                continue
            region = view.expand_to_scope(pt, selector)
            if region is None:
                continue
            text = view.substr(region).strip()
            # BOTH brackets fall inside the body meta scope: it is a meta_scope,
            # so it covers the text of the pushing match ('[[') as well as that
            # of the popping one (']]'). BODY_SELECTOR above is the bare meta
            # scope with no subtraction, so expand_to_scope() hands back the
            # whole '[[name]]'. Stripping only one end is the bug this replaces:
            # it left "[[francis", which happened to agree with the equally
            # wrong name the reference index was storing, so the two masked each
            # other and only a query built any other way (the note's own
            # basename) showed that body links were unreachable. See the header
            # comment in 'Reference List - Wiki Body.tmPreferences'.
            if text.startswith("[["):
                text = text[2:]
            if text.endswith("]]"):
                text = text[:-2]
            return text.strip()
    return None


def target_name(view, under_caret=False):
    """What a backlink query is about.

    By default: the note in view, whatever the caret happens to be sitting in.
    "Who links here?" is a question about the note being read, and answering it
    from the caret made the answer depend on where the cursor was left -- a
    [[link]] mentioned in passing silently redirected the panel to a note that
    was never asked about, and the two cases are indistinguishable once the
    panel is open, because it lists paths rather than naming its subject.

    under_caret=True restores the other reading: the link under the caret, i.e.
    "who else links to the note this link points at?", falling back to the note
    in view when the caret is not in a link. It is a separate palette entry
    rather than the default for the reason above.
    """
    if under_caret:
        under = name_under_caret(view)
        if under:
            return under
    return note_name(view.file_name()) if view.file_name() else None


# ------------------------------------------------------------------
# Command: show backlinks
# ------------------------------------------------------------------

class WikiBacklinksCommand(sublime_plugin.TextCommand):
    """Every note linking to this one, from the index.

    The subject is the note in view. Pass under_caret=True to ask about the
    link the caret is in instead -- see target_name() for why that is the arg
    and not the default.
    """

    def is_enabled(self):
        return bool(self.view.file_name())

    def run(self, edit, under_caret=False):
        window = self.view.window() or sublime.active_window()
        if not window:
            return

        name = target_name(self.view, under_caret)
        if not name:
            sublime.status_message("WikiLinks: save the note first")
            return

        locations = wiki_references(window, name)
        if not locations:
            window.status_message(
                "No links to [[{0}]] (is the folder in this project and "
                "indexed?)".format(name))
            return

        self.name = name
        self.window = window
        self.locations = sorted(locations, key=lambda loc: (loc.path, loc.row))

        # Where Esc has to put you back. Sublime restores the view behind its
        # OWN Goto panels; a plugin's show_quick_panel does not, so whatever
        # the last on_highlight previewed is simply left in front of you. The
        # caret is saved too, because a row can point AT the note you started
        # in (a self-link, or a note that appears in its own backlinks) and
        # arrowing over it moves that caret with ENCODED_POSITION.
        self.origin = self.view
        self.origin_id = self.view.id()
        self.origin_sel = [r for r in self.view.sel()]
        self.origin_viewport = self.view.viewport_position()

        francis = sibling("francis")
        if francis:
            francis._PREVIEW_FOLDS.clear()

        window.show_quick_panel(
            self.items(francis), self.on_done, on_highlight=self.on_highlight)

    def items(self, francis):
        items = []
        for loc in self.locations:
            trigger, detail = self.row_text(loc, francis)
            items.append(sublime.QuickPanelItem(
                trigger,
                details=detail,
                annotation="line {0}".format(loc.row),
                kind=self.kind_of(loc),
            ))
        return items

    def row_text(self, loc, francis):
        """An email's filename already carries its date, sender and subject, so
        a backlink from one reads as correspondence rather than as a
        100-character path. Nothing is opened to find that out."""
        name = note_name(loc.path)
        if francis:
            fields = francis.parse_email_filename(loc.path)
            if fields:
                return (
                    "{0}  ({1})".format(fields["sender"] or "unknown sender",
                                        fields["date"]),
                    fields["subject"] or name,
                )
        folder = os.path.basename(os.path.dirname(loc.path))
        return name, os.path.join(folder, os.path.basename(loc.path))

    def kind_of(self, loc):
        """Label the row by kind when the index carries one.

        Whether a .tmPreferences 'kind' survives into an indexed REFERENCE is
        unverified: no shipped package sets that key. So this reads whatever is
        there and falls back, rather than depending on the answer.
        """
        kind = getattr(loc, "kind", None)
        if isinstance(kind, tuple) and len(kind) == 3:
            return kind
        return sublime.KIND_AMBIGUOUS

    def on_highlight(self, index):
        if index < 0:
            return
        view = self.window.open_file(
            self.locations[index].path_encoded_position(),
            sublime.ENCODED_POSITION | sublime.TRANSIENT)
        francis = sibling("francis")
        if francis:
            francis.fold_frontmatter(self.window, view)

    def on_done(self, index):
        selected = None
        if index >= 0:
            # Same file the last on_highlight previewed, so this promotes that
            # transient sheet to a real tab rather than opening a second one.
            selected = self.window.open_file(
                self.locations[index].path_encoded_position(),
                sublime.ENCODED_POSITION)
        francis = sibling("francis")
        if francis:
            francis.restore_preview_folds(selected)
        if selected is None:
            self.restore_origin()

    def restore_origin(self):
        """Undo the previews on Esc: close what we opened, go back where we were.

        A transient sheet is only ever the preview this panel asked for -- a
        view cannot be both transient and edited, since typing into a preview
        promotes it -- so closing it loses nothing. The identity test is for the
        case where the note you started in was itself opened as a preview by
        something else; closing that would be closing the thing we are trying to
        return to. Folds are restored before this runs, by restore_preview_folds
        in francis.py, which only ever recorded PRE-EXISTING tabs.
        """
        for group in range(self.window.num_groups()):
            transient = self.window.transient_view_in_group(group)
            if transient and transient.id() != self.origin_id:
                transient.close()

        if self.origin is None or not self.origin.is_valid():
            return
        self.window.focus_view(self.origin)
        self.origin.sel().clear()
        self.origin.sel().add_all(self.origin_sel)
        self.origin.set_viewport_position(self.origin_viewport, False)


# ------------------------------------------------------------------
# Command: rename a note and every link to it
# ------------------------------------------------------------------

class WikiRenameNoteCommand(sublime_plugin.WindowCommand):
    """Rename the note in the active view and rewrite every link to it.

    Locations come from the index, so no file is searched. Each write is checked
    against the line the index pointed at: if it no longer holds the link, that
    location is reported and skipped, never guessed at.
    """

    def is_enabled(self):
        view = self.window.active_view()
        return bool(view and view.file_name()
                    and view.file_name().lower().endswith(".md"))

    def run(self):
        view = self.window.active_view()
        if not view or not view.file_name():
            return
        self.path = view.file_name()
        self.old = note_name(self.path)
        self.window.show_input_panel(
            "Rename note (and every link to it):", self.old,
            self.on_done, None, None)

    # --- validation ---------------------------------------------------

    def on_done(self, raw):
        new = raw.strip()
        if new.lower().endswith(".md"):
            new = new[:-3]
        if not new or new == self.old:
            return
        if os.path.sep in new or (os.path.altsep and os.path.altsep in new):
            self.fail("a note name cannot contain a path separator")
            return

        index = basename_index(self.window.folders())
        if new in index:
            self.fail("[[{0}]] already exists: {1}".format(new, index[new][0]))
            return

        grouped = group_rows(wiki_references(self.window, self.old))

        dirty = [p for p in sorted(grouped) if self.is_dirty(p)]
        if self.is_dirty(self.path):
            dirty.append(self.path)
        if dirty:
            self.fail("save these first ({0}): {1}".format(
                len(dirty), ", ".join(note_name(p) for p in dirty[:5])))
            return

        if not self.confirm(new, grouped):
            return
        self.apply(new, grouped)

    def is_dirty(self, path):
        view = self.window.find_open_file(path)
        return bool(view and view.is_dirty())

    def fail(self, message):
        sublime.error_message("Rename: {0}".format(message))

    def confirm(self, new, grouped):
        """The count IS the completeness check.

        Whether the index has seen everything is the one property a lookup
        cannot report on itself, so what stands in for it is you recognising the
        number before it is written. 'Wiki: Show backlinks' shows the same set
        row by row if it looks wrong.
        """
        links = sum(len(rows) for rows in grouped.values())
        summary = "Rename [[{0}]] to [[{1}]]\n\n{2} link(s) in {3} file(s).".format(
            self.old, new, links, len(grouped))
        if len(grouped) > RENAME_WARN_FILES:
            summary += ("\n\nThat is a large rewrite. If the number looks low "
                        "the index may still be catching up: cancel and check "
                        "the backlinks panel first.")
        summary += "\n\n" + SKIPPED_NOTE
        return sublime.ok_cancel_dialog(summary, "Rename")

    # --- the write ----------------------------------------------------

    def apply(self, new, grouped):
        log("rename [[{0}]] -> [[{1}]]: {2} file(s) reported".format(
            self.old, new, len(grouped)))
        log("  note: {0}".format(SKIPPED_NOTE))

        applied = files = 0
        problems = []
        for path in sorted(grouped):
            count, skipped, error = rewrite_file(
                path, grouped[path], self.old, new)
            if error:
                problems.append("{0}: {1}".format(path, error))
            elif skipped:
                problems.append("{0}: rows {1} no longer hold the link".format(
                    path, ", ".join(str(r) for r in skipped)))
            if count:
                applied += count
                files += 1

        target = os.path.join(os.path.dirname(self.path), new + ".md")
        try:
            os.rename(self.path, target)
        except OSError as e:
            self.fail("links were rewritten, but the file itself could not be "
                      "renamed: {0}".format(e))
            log("FAILED to rename the file: {0}".format(e))
            return

        view = self.window.find_open_file(self.path)
        if view:
            view.retarget(target)

        for problem in problems:
            log("  skipped {0}".format(problem))
        log("rename done: {0} link(s) in {1} file(s), {2} problem(s)".format(
            applied, files, len(problems)))

        message = "Renamed to [[{0}]]: {1} link(s) in {2} file(s)".format(
            new, applied, files)
        if problems:
            message += "; {0} skipped (see Data/Logs/WikiLinks.log)".format(
                len(problems))
        self.window.status_message(message)
        print("WikiLinks: {0}".format(message))


# ------------------------------------------------------------------
# Command: pick a note and write [[its basename]]
# ------------------------------------------------------------------

SETTINGS_FILE = "WikiLinks.sublime-settings"

# Only reached if the settings file is missing outright. Kept in step with
# WikiLinks.sublime-settings, which is where the reasoning lives.
DEFAULT_LINK_CLASSES = ["topic", "activity", "programme", "person",
                        "organisation", "meeting", "memo", "document", "email"]
DEFAULT_HIDDEN_CLASSES = ["email"]
DEFAULT_EXCLUDE = ["*_template.md"]

KIND_CLASS_FILTER = (sublime.KIND_ID_NAVIGATION, "/", "Filter")

# A row badge says which kind of note it is, which is the one thing a bare
# basename does not tell you -- "energyForGrowth" could be a programme or a
# topic. An unlisted class falls back rather than raising, so a class added to
# the settings file works with no edit here.
CLASS_KINDS = {
    "topic":        (sublime.KIND_ID_COLOR_PURPLISH, "T", "Topic"),
    "activity":     (sublime.KIND_ID_COLOR_ORANGISH, "A", "Activity"),
    "programme":    (sublime.KIND_ID_COLOR_REDISH, "P", "Programme"),
    "person":       (sublime.KIND_ID_COLOR_GREENISH, "@", "Person"),
    "organisation": (sublime.KIND_ID_COLOR_BLUISH, "O", "Organisation"),
    "meeting":      (sublime.KIND_ID_COLOR_CYANISH, "M", "Meeting"),
    "memo":         (sublime.KIND_ID_COLOR_YELLOWISH, "m", "Memo"),
    "document":     (sublime.KIND_ID_COLOR_YELLOWISH, "D", "Document"),
    "email":        (sublime.KIND_ID_COLOR_CYANISH, "E", "Email"),
}


def settings():
    return sublime.load_settings(SETTINGS_FILE)


def link_classes():
    return list(settings().get("link_classes", DEFAULT_LINK_CLASSES))


def hidden_classes():
    return list(settings().get("hidden_classes", DEFAULT_HIDDEN_CLASSES))


def kind_for_class(cls):
    return CLASS_KINDS.get(cls, sublime.KIND_AMBIGUOUS)


def excluded(path, patterns):
    """fnmatch against the file name and the full path.

    Same convention as TodoSync and WikiResolve, so one "*_template.md" means
    the same thing in all three settings files.
    """
    name = os.path.basename(path)
    for pattern in patterns:
        if fnmatch.fnmatch(name, pattern) or fnmatch.fnmatch(path, pattern):
            return True
    return False


def class_notes(window, classes, exclude=None):
    """[(name, path, cls)] for every indexed note in `classes`.

    One symbol_locations call per class -- the same lookup TodoSync makes to
    resolve its scope and francis.symbol_paths makes for the email flags. No
    file is opened and no directory is walked, so the cost is flat in vault
    size, and a class the index has not seen contributes nothing rather than
    failing.

    Ordered by the class order given, then by name within each class, so the
    four topics and eight programmes sit above 375 people while the query is
    still empty. Fuzzy scoring takes over the moment anything is typed.

    The consequence of being index-only is worth naming: a note with no
    "class:" key produces no symbol and is therefore not a candidate. In this
    vault that is 124 files, all of #countries/. Classifying one is what puts
    it in the list.
    """
    if exclude is None:
        exclude = settings().get("exclude", DEFAULT_EXCLUDE)

    notes, seen = [], set()
    for cls in classes:
        rows = []
        for loc in window.symbol_locations(cls + "_note",
                                           sublime.SYMBOL_SOURCE_INDEX):
            if loc.path in seen or excluded(loc.path, exclude):
                continue
            seen.add(loc.path)
            rows.append((note_name(loc.path), loc.path, cls))
        rows.sort(key=lambda row: row[0].lower())
        notes.extend(rows)
    return notes


class WikiNoteListInput(sublime_plugin.ListInputHandler):
    """One step of the note picker.

    Same shape as EmailListInput in francis.py, and for the same reason: the
    "/class" rows live in the SAME list as the notes, so typing "/" simply
    fuzzy-filters down to them. There is no callback for typing, so a keystroke
    cannot switch modes -- choosing a row returns a fresh handler, and that is
    what makes Sublime draw the chip and re-populate the list.

    Classes UNION rather than narrow. A note carries exactly one "class:", so
    intersecting /person with /organisation the way the email filters intersect
    would always give the empty set. No chips means every class except the
    hidden ones; "/email" is how the 1,746 emails get in.

    No preview(). EmailListInput can open a transient view because its command
    only opens a file afterwards; this one is a TextCommand that still has to
    write into self.view, and a transient preview would move the active view
    out from under it.
    """

    def __init__(self, window, applied=(), depth=0, seed=""):
        self.window = window
        self.applied = tuple(applied)
        self.depth = depth
        self.seed = seed

    def name(self):
        # One command argument per step. Unlike EmailListInput the depth is
        # exactly the number of filters, since there is no field flip here --
        # but the naming is kept identical so the two read the same.
        return "step{0}".format(self.depth)

    def placeholder(self):
        return "Link to a note  ( / to filter by class )"

    def initial_text(self):
        # Whatever was already between the brackets, so "[fran" + tab opens the
        # picker already filtered to fran. First step only: a later step
        # rebuilt the list because the FILTER changed, and re-seeding the query
        # there would fight the chip just added.
        return self.seed if self.depth == 0 else ""

    def description(self, value, text):
        return text.lstrip("/")

    def classes(self):
        """Which classes this step lists."""
        if self.applied:
            return [c for c in link_classes() if c in self.applied]
        hidden = hidden_classes()
        return [c for c in link_classes() if c not in hidden]

    def list_items(self):
        items = [
            sublime.ListInputItem(
                "/" + cls,
                "cls:" + cls,
                details="Show {0} notes".format(cls),
                kind=KIND_CLASS_FILTER,
            )
            for cls in link_classes() if cls not in self.applied
        ]

        for name, path, cls in class_notes(self.window, self.classes()):
            # The trigger IS the text that gets inserted, and it is also the
            # only text matched against typed input (see ListInputItem.text).
            # The folder drops to the detail line so it can be read without
            # ever being searched.
            items.append(sublime.ListInputItem(
                name,
                name,
                details=os.path.basename(os.path.dirname(path)),
                annotation=cls,
                kind=kind_for_class(cls),
            ))
        return items

    def next_input(self, args):
        value = args.get(self.name(), "")
        if isinstance(value, str) and value.startswith("cls:"):
            return WikiNoteListInput(self.window, self.applied + (value[4:],),
                                     self.depth + 1, self.seed)
        return None   # a note was chosen; the command runs


class WikiPickLinkCommand(sublime_plugin.TextCommand):
    """Pick a note from the symbol index and write [[its basename]].

    TWO PHASES, ONE COMMAND, and the split is what the keystroke wants:

      Phase 1 -- no `region` argument, i.e. the tab keystroke or the palette
      entry. Turns the bracket pair around the caret into "[[seed]]" and THEN
      opens the picker, so tab visibly does something before the overlay
      appears.

      Phase 2 -- `region` present, plus the step values the input chain
      collected. Replaces the text inside those brackets with the chosen name.

    Opening the command palette by hand is not a trick: it is exactly what
    sublime_plugin.TextCommand.run_ does for itself (Lib/python33/
    sublime_plugin.py, the "required positional argument" branch) when a
    command carrying an input handler is invoked without one. Relying on that
    branch instead would mean run() never executes before the overlay, so the
    bracket could not be doubled first -- and there would be nowhere to put
    `region`.

    `region` is passed rather than re-derived so a seed containing a stray
    bracket cannot mis-target the write. Nothing can move underneath it while
    it is in flight: the overlay is modal.
    """

    def is_enabled(self):
        if not self.view.window():
            return False
        pt = self.view.sel()[0].begin() if len(self.view.sel()) else 0
        return self.view.match_selector(pt, "text.html.markdown")

    def input_description(self):
        return "Link to"

    def input(self, args):
        applied, depth = [], 0
        while True:
            value = args.get("step{0}".format(depth))
            if not isinstance(value, str):
                break
            if not value.startswith("cls:"):
                return None      # a note is already chosen; nothing to ask
            applied.append(value[4:])
            depth += 1

        seed = ""
        region = args.get("region")
        if depth == 0 and region:
            seed = self.view.substr(sublime.Region(region[0], region[1]))
        return WikiNoteListInput(self.view.window(), tuple(applied), depth, seed)

    def run(self, edit, region=None, **steps):
        if region is None:
            self.start(edit)
            return

        name = None
        for key in sorted((k for k in steps if k.startswith("step")),
                          key=lambda k: int(k[4:])):
            value = steps[key]
            if isinstance(value, str) and not value.startswith("cls:"):
                name = value
        if not name:
            return

        target = sublime.Region(region[0], region[1])
        self.view.replace(edit, target, name)

        # Past the closing ']]', which replace() did not touch.
        end = target.begin() + len(name) + 2
        self.view.sel().clear()
        self.view.sel().add(sublime.Region(end, end))

    def start(self, edit):
        window = self.view.window()
        if not window:
            return

        pt = self.view.sel()[0].begin() if len(self.view.sel()) else 0
        line = self.view.line(pt)
        span = bracket_span(self.view.substr(line), pt - line.begin())
        if span is None:
            # No pair around the caret -- the palette entry invoked from open
            # prose. Insert a fresh empty link rather than refusing.
            begin, end, seed = pt, pt, ""
        else:
            begin = line.begin() + span[0]
            end = line.begin() + span[1]
            seed = span[2]

        self.view.replace(edit, sublime.Region(begin, end),
                          "[[{0}]]".format(seed))
        inner = [begin + 2, begin + 2 + len(seed)]
        self.view.sel().clear()
        self.view.sel().add(sublime.Region(inner[1], inner[1]))

        # Deferred so this edit is committed before the overlay takes focus.
        sublime.set_timeout(lambda: window.run_command("show_overlay", {
            "overlay": "command_palette",
            "command": "wiki_pick_link",
            "args": {"region": inner},
        }), 0)
