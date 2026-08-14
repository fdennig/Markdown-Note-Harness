"""Headless tests for WikiLinks.py.

    "%USERPROFILE%\\.local\\bin\\python3.14.exe" tests/test_wikilinks.py

Covers the half of the feature that does not need Sublime: turning a set of
reported (path, row) locations into correct bytes on disk. The other half --
whether the index reports the right locations in the first place -- is a manual
step, because no stub reproduces Sublime's indexer. See CLAUDE.md for the
counts to check it against.

What is worth testing here is narrow but load-bearing:

  * the row base is measured rather than assumed, so an off-by-one in
    SymbolLocation.row cannot rewrite the line above every link in the vault;
  * only reported rows are touched, which is what keeps the macro-owned raw
    'to:' keys out of a rename;
  * BOM and CRLF survive, because these files are both and a rename must not be
    what rewrites 1,700 of them end to end;
  * the picker refuses a bracket span it cannot read rather than rewriting the
    wrong part of the line, and its candidate list excludes the templates --
    which carry a real class: and would otherwise look like linkable notes.

The picker's Settings stub reads the values shipped in
WikiLinks.sublime-settings rather than inventing its own, so the file and the
module fallbacks cannot drift apart unnoticed.

Must stay in tests/ and never move to the package root: it assigns a fake into
sys.modules["sublime"] at import, so Sublime auto-loading it would break every
plugin in the process.
"""

import io
import os
import shutil
import sys
import tempfile
import types

TMP = tempfile.mkdtemp(prefix="wikilinks_")
# Derived, not hardcoded: the package directory has been renamed once already,
# and "python tests/test_wikilinks.py" puts tests/ on sys.path, not the package
# root, so a wrong constant here is a hard import failure.
PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ---------------- stubs ----------------

class Region(object):
    def __init__(self, a, b=None):
        self.a = a
        self.b = a if b is None else b

    def begin(self):
        return min(self.a, self.b)

    def end(self):
        return max(self.a, self.b)


class View(object):
    """Enough of a view for name_under_caret.

    `spans` is the scope map as (begin, end, selector) triples, which is exactly
    what the syntax produces and what expand_to_scope reads back.
    """

    def __init__(self, text, spans, caret):
        self.text = text
        self.spans = spans
        self.caret = caret
        self._file_name = "note.md"

    def sel(self):
        return [Region(self.caret)]

    def file_name(self):
        return self._file_name

    def _span(self, pt, selector):
        for begin, end, scope in self.spans:
            if begin <= pt < end and scope.startswith(selector):
                return Region(begin, end)
        return None

    def match_selector(self, pt, selector):
        return self._span(pt, selector) is not None

    def expand_to_scope(self, pt, selector):
        return self._span(pt, selector)

    def substr(self, region):
        return self.text[region.begin():region.end()]


class Loc(object):
    def __init__(self, path, row=1):
        self.path = path
        self.row = row


class Settings(object):
    """The shipped WikiLinks.sublime-settings, read for real.

    Not a dict of made-up values: the point of these tests is partly that the
    defaults in the file and the fallbacks in the module have not drifted, and
    a hand-written stub would hide exactly that.
    """

    VALUES = {
        "link_classes": ["topic", "activity", "programme", "person",
                         "organisation", "meeting", "memo", "document",
                         "email"],
        "hidden_classes": ["email"],
        "exclude": ["*_template.md"],
    }

    def get(self, key, default=None):
        return self.VALUES.get(key, default)


class Window(object):
    """Enough of a window for class_notes.

    `index` is {symbol: [paths]}, which is what window.symbol_locations
    answers from -- a class Sublime has not indexed simply is not a key, and
    the empty list back is the behaviour that must not raise.
    """

    def __init__(self, index):
        self.index = index

    def symbol_locations(self, symbol, source=None, type=None):
        return [Loc(p) for p in self.index.get(symbol, [])]


class _ListInputHandler(object):
    pass


sublime = types.ModuleType("sublime")
sublime.Region = Region
sublime.SYMBOL_SOURCE_ANY = 0
sublime.SYMBOL_SOURCE_INDEX = 1
sublime.SYMBOL_TYPE_REFERENCE = 2
sublime.KIND_AMBIGUOUS = (0, "", "")
sublime.ENCODED_POSITION = 1
sublime.TRANSIENT = 2
sublime.packages_path = lambda: os.path.join(TMP, "Packages")
sublime.status_message = lambda m: None
sublime.error_message = lambda m: None
sublime.ok_cancel_dialog = lambda *a, **k: True
sublime.QuickPanelItem = lambda *a, **k: None
sublime.ListInputItem = lambda *a, **k: None
sublime.load_settings = lambda name: Settings()
sublime.set_timeout = lambda fn, ms=0: None
for _kind in ("NAVIGATION", "COLOR_PURPLISH", "COLOR_ORANGISH", "COLOR_REDISH",
              "COLOR_GREENISH", "COLOR_BLUISH", "COLOR_CYANISH",
              "COLOR_YELLOWISH"):
    setattr(sublime, "KIND_ID_" + _kind, _kind)
sys.modules["sublime"] = sublime

sublime_plugin = types.ModuleType("sublime_plugin")


class _TextCommand(object):
    def __init__(self, view):
        self.view = view


class _WindowCommand(object):
    def __init__(self, window):
        self.window = window


# Split in two on purpose: TextCommand must set self.view and WindowCommand
# self.window. A single shared base hides the difference and then every
# is_enabled that reads the wrong one passes for the wrong reason.
sublime_plugin.TextCommand = _TextCommand
sublime_plugin.WindowCommand = _WindowCommand
sublime_plugin.ListInputHandler = _ListInputHandler
sys.modules["sublime_plugin"] = sublime_plugin

sys.path.insert(0, PKG)
import WikiLinks as W          # noqa: E402
from WikiLinks import WikiNoteListInput   # noqa: E402


# ---------------- harness ----------------

FAILURES = []


def check(label, got, want):
    if got == want:
        print("  ok   {0}".format(label))
    else:
        print("  FAIL {0}\n         got:  {1!r}\n         want: {2!r}".format(
            label, got, want))
        FAILURES.append(label)


def write(path, text):
    d = os.path.dirname(path)
    if not os.path.exists(d):
        os.makedirs(d)
    with io.open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    return path


def raw(path):
    with io.open(path, "rb") as f:
        return f.read()


# ---------------- tests ----------------

def test_names():
    print("names")
    check("note_name strips dir and extension",
          W.note_name(os.path.join("a", "b", "silviaKuehl.md")), "silviaKuehl")
    check("link_text wraps", W.link_text("francis"), "[[francis]]")


def test_row_base():
    print("row base is measured, not assumed")
    lines = ["---", 'from_wiki: "[[francis]]"', "---", "body [[francis]]"]
    # rows as Sublime reports them if 1-based
    check("1-based detected", W.detect_row_base([2, 4], lines, "francis"), 1)
    # the same lines, rows as they would be if 0-based
    check("0-based detected", W.detect_row_base([1, 3], lines, "francis"), 0)
    # 1 wins a tie, matching what ENCODED_POSITION implies
    tie = ["[[a]]", "[[a]]", "[[a]]"]
    check("tie prefers 1-based", W.detect_row_base([1], tie, "a"), 1)
    check("no base fits -> None",
          W.detect_row_base([99], lines, "francis"), None)
    check("no rows is vacuously fine",
          W.detect_row_base([], lines, "francis"), 1)


def test_only_reported_rows():
    print("only reported rows are rewritten")
    lines = [
        "---",
        'to:',                      # raw macro key, holds a bare basename
        '  - "[[francis]]"',        # row 3, NOT reported by the index
        'to_wiki:',
        '  - "[[francis]]"',        # row 5, reported
        "---",
        "and [[francis]] in the body",   # row 7, reported
    ]
    applied, skipped = W.rewrite_lines(lines, [5, 7], "francis", "fd", base=1)
    check("applied count", applied, 2)
    check("skipped none", skipped, [])
    check("raw to: untouched", lines[2], '  - "[[francis]]"')
    check("to_wiki rewritten", lines[4], '  - "[[fd]]"')
    check("body rewritten", lines[6], "and [[fd]] in the body")


def test_prefix_safety():
    print("prefix safety")
    lines = ["[[francis]] and [[francisDennig]] and [[francis]]"]
    applied, skipped = W.rewrite_lines(lines, [1], "francis", "fd", base=1)
    check("both hits on the line", applied, 2)
    check("longer name untouched",
          lines[0], "[[fd]] and [[francisDennig]] and [[fd]]")


def test_stale_row_is_skipped():
    print("a stale row is reported, never guessed")
    lines = ["nothing here", "[[francis]]"]
    applied, skipped = W.rewrite_lines(lines, [1], "francis", "fd", base=1)
    check("nothing applied", applied, 0)
    check("row reported", skipped, [1])
    check("file left alone", lines, ["nothing here", "[[francis]]"])


def test_bom_and_crlf():
    print("BOM and CRLF survive a rewrite")
    body = ('﻿---\r\n'
            'class: email\r\n'
            'from_wiki: "[[francis]]"\r\n'
            'to_wiki:\r\n'
            '  - "[[ago]]"\r\n'
            '  - "[[francis]]"\r\n'
            '---\r\n'
            '\r\n'
            'See [[francis]] on this.\r\n')
    p = write(os.path.join(TMP, "vault", "emails", "e1.md"), body)
    before = raw(p)
    check("fixture really has a BOM", before[:3], b"\xef\xbb\xbf")

    applied, skipped, error = W.rewrite_file(p, [3, 6, 9], "francis", "fd")
    check("error", error, None)
    check("applied", applied, 3)
    check("skipped", skipped, [])

    after = raw(p)
    check("BOM still there", after[:3], b"\xef\xbb\xbf")
    check("exactly one BOM", after.count(b"\xef\xbb\xbf"), 1)
    check("no LF was orphaned", after.count(b"\r\n"), before.count(b"\r\n"))
    check("no bare LF introduced",
          after.replace(b"\r\n", b"").count(b"\n"), 0)
    text = io.open(p, encoding="utf-8", newline="").read()
    check("ago untouched", '  - "[[ago]]"\r\n' in text, True)
    check("all three rewritten", text.count("[[fd]]"), 3)
    check("no old links left", text.count("[[francis]]"), 0)


def test_lf_file_stays_lf():
    print("an LF file stays LF")
    p = write(os.path.join(TMP, "vault", "lf.md"), "a\n[[x]]\nb\n")
    W.rewrite_file(p, [2], "x", "y")
    check("no CR introduced", b"\r" in raw(p), False)
    check("trailing newline kept", raw(p), b"a\n[[y]]\nb\n")


def test_no_trailing_newline():
    print("a file with no trailing newline keeps none")
    p = write(os.path.join(TMP, "vault", "bare.md"), "[[x]]")
    W.rewrite_file(p, [1], "x", "y")
    check("content", raw(p), b"[[y]]")


def test_basename_index():
    print("basename index")
    root = os.path.join(TMP, "vault")
    write(os.path.join(root, "people", "francis.md"), "x")
    write(os.path.join(root, ".git", "francis.md"), "x")
    write(os.path.join(root, "notes.txt"), "x")
    index = W.basename_index([root])
    check("finds the note", "francis" in index, True)
    check(".git skipped, so no false collision",
          len(index["francis"]), 1)
    check("non-markdown ignored", "notes" in index, False)


def test_group_rows():
    print("grouping")
    grouped = W.group_rows([Loc("a.md", 3), Loc("b.md", 1), Loc("a.md", 9)])
    check("by path", sorted(grouped), ["a.md", "b.md"])
    check("rows kept", sorted(grouped["a.md"]), [3, 9])


def test_name_under_caret():
    print("name under caret")
    text = 'see [[silviaKuehl]] here'
    # The body meta scope covers BOTH brackets -- it is a meta_scope, so it
    # spans the pushing '[[' as well as the popping ']]'. The span below is the
    # whole '[[silviaKuehl]]' for that reason, and getting it wrong here is what
    # let the one-sided strip ship: the old stub started the span after '[[',
    # so a leading bracket the real editor hands over never reached the code.
    body = [(4, 19, "meta.link.reference.wiki.description.markdown")]
    check("body link, both brackets stripped",
          W.name_under_caret(View(text, body, 8)), "silviaKuehl")

    check("caret on the opening bracket",
          W.name_under_caret(View(text, body, 4)), "silviaKuehl")

    fm = '  - "[[ago]]"'
    spans = [(7, 10, "entity.name.wikilink.to.yaml")]
    check("frontmatter link is already bare",
          W.name_under_caret(View(fm, spans, 8)), "ago")

    check("caret outside a link",
          W.name_under_caret(View(text, body, 0)), None)

    # target_name's DEFAULT is the note in view, caret or no caret: "who links
    # here?" is a question about the note being read, and answering it from the
    # caret made the panel's subject depend on where the cursor was left.
    check("default ignores the link under the caret",
          W.target_name(View(text, body, 8)), "note")

    check("under_caret=True asks about the link instead",
          W.target_name(View(text, body, 8), under_caret=True), "silviaKuehl")

    check("under_caret=True falls back outside a link",
          W.target_name(View(text, body, 0), under_caret=True), "note")


def test_bracket_span():
    print("bracket span")
    check("empty pair, caret between",
          W.bracket_span("[]", 1), (0, 2, ""))
    check("seeded pair",
          W.bracket_span("see [fran] here", 8), (4, 10, "fran"))
    check("caret against the closing bracket",
          W.bracket_span("[fran]", 5), (0, 6, "fran"))

    # The whole point of widening: a wiki link is the same shape one character
    # further out, so one function serves both tab bindings.
    check("wiki link widens to all four brackets",
          W.bracket_span("[[x]]", 3), (0, 5, "x"))
    check("empty wiki link",
          W.bracket_span("[[]]", 2), (0, 4, ""))

    check("no brackets on the line",
          W.bracket_span("just prose", 4), None)
    check("caret before any opening bracket",
          W.bracket_span("[fran]", 0), None)
    # Caret sits between ']' and 'b]' -- the text the naive span would cover
    # holds a bracket of its own, so there is nothing safe to rewrite.
    check("unbalanced span is refused",
          W.bracket_span("[a]b]", 4), None)


def vault_window():
    """A window whose index looks like the real one, templates included."""
    return Window({
        "topic_note": [os.path.join("notes", "energyForGrowth.md")],
        "programme_note": [os.path.join("notes", "climatePromise.md"),
                           os.path.join("newSystem", "programme_template.md")],
        "person_note": [os.path.join("people", "silviaKuehl.md"),
                        os.path.join("people", "agostinoInguscio.md"),
                        os.path.join("newSystem", "person_template.md")],
        "organisation_note": [os.path.join("organisations", "undp.md")],
        "email_note": [os.path.join("emails", "Email 2026-01-02_1030 @X - Y.md")],
    })


def test_class_notes():
    print("class notes")
    w = vault_window()

    check("name is the basename, class rides along",
          W.class_notes(w, ["topic"]),
          [("energyForGrowth", os.path.join("notes", "energyForGrowth.md"),
            "topic")])

    # The templates carry a real class: and would otherwise be linkable notes.
    check("templates are excluded",
          [n for n, _, _ in W.class_notes(w, ["programme", "person"])],
          ["climatePromise", "agostinoInguscio", "silviaKuehl"])

    # Class order is the caller's; names sort within each class. That is what
    # puts eight programmes above 375 people while the query is empty.
    check("class order outranks name order",
          [c for _, _, c in W.class_notes(w, ["programme", "person"])],
          ["programme", "person", "person"])

    check("a class the index has never seen is silent, not fatal",
          W.class_notes(w, ["nosuchclass"]), [])

    check("exclude is overridable",
          [n for n, _, _ in W.class_notes(w, ["person"], exclude=[])],
          ["agostinoInguscio", "person_template", "silviaKuehl"])


def test_class_filters():
    print("class filters")
    w = vault_window()

    # No chips: every class except the hidden ones. The 1,746 emails staying
    # out until asked for is the whole reason hidden_classes exists.
    default = WikiNoteListInput(w).classes()
    check("email is hidden by default", "email" in default, False)
    check("person is not", "person" in default, True)

    check("one chip selects that class only",
          WikiNoteListInput(w, ("person",)).classes(), ["person"])

    # Union, not intersection: a note has exactly one class, so intersecting
    # would always give nothing.
    check("two chips union",
          WikiNoteListInput(w, ("person", "organisation")).classes(),
          ["person", "organisation"])

    check("/email is how emails get in",
          [n for n, _, _ in W.class_notes(
              w, WikiNoteListInput(w, ("email",)).classes())],
          ["Email 2026-01-02_1030 @X - Y"])

    # The chip text, and the argument name that carries it forward.
    step = WikiNoteListInput(w, (), 0, "fran")
    check("seed pre-filters the first step only", step.initial_text(), "fran")
    check("and not the next one",
          step.next_input({"step0": "cls:person"}).initial_text(), "")
    check("each step is its own command argument",
          step.next_input({"step0": "cls:person"}).name(), "step1")
    check("choosing a note ends the chain",
          step.next_input({"step0": "silviaKuehl"}), None)
    check("chip drops the slash", step.description("cls:person", "/person"),
          "person")


def main():
    for test in (test_names, test_row_base, test_only_reported_rows,
                 test_prefix_safety, test_stale_row_is_skipped,
                 test_bom_and_crlf, test_lf_file_stays_lf,
                 test_no_trailing_newline, test_basename_index,
                 test_group_rows, test_name_under_caret,
                 test_bracket_span, test_class_notes, test_class_filters):
        test()
    print("")
    if FAILURES:
        print("{0} FAILURE(S): {1}".format(len(FAILURES), ", ".join(FAILURES)))
    else:
        print("all passed")
    shutil.rmtree(TMP, ignore_errors=True)
    return 1 if FAILURES else 0


if __name__ == "__main__":
    sys.exit(main())
