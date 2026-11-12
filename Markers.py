"""Markers -- the vault's two-way channel between the user and the agent.

    - [>] 20261112 — an instruction from the user to the agent
    - [?] 20261112 — the agent's question or proposal, left for the user
    - [-] 20261112 — a [?] the user dismissed as stale

The vault's CLAUDE.md is the contract; this module only makes the three lines
cheap to type, always with the date raised and the em-dash separator that
`scripts/questions.py` parses.

  marker_newline          Enter on a marker line continues with a fresh dated
                          marker. Bound in Packages/User/Default.sublime-keymap,
                          not here: MarkdownEditing's generic-bullet Enter loads
                          after this package ('-' sorts before 'E') and would
                          win; User loads last.
  marker_open_directive   "." then tab on an otherwise empty line opens a [>].
  marker_toggle_dismiss   alt+z toggles [?] <-> [-] on every line with a caret.

Rebuilt 20261112 from the vault's description after the original was lost with
the old machine; behaviour beyond what that description fixes is a choice made
then, documented on the command that makes it.

Python 3.3 plugin host: no f-strings, nothing from the 3.4+ standard library.
"""
import re
import time

import sublime
import sublime_plugin

SEP = u" \u2014 "
# indent (incl. blockquote '>'), bullet, box, optional date and separator, rest
MARKER_RE = re.compile(
    u"^([\\s>]*)([-+*])(\\s+)\\[([?>-])\\](?:[ \\t]+(\\d{8}))?"
    u"(?:[ \\t]*(?:\u2014|\u2013|--|-|:))?[ \\t]*(.*)$")


def today():
    return time.strftime("%Y%m%d")


def stamp(indent, bullet, gap, box):
    return u"{0}{1}{2}[{3}] {4}{5}".format(indent, bullet, gap, box, today(), SEP)


class MarkerNewlineCommand(sublime_plugin.TextCommand):
    """Enter on a marker line.

    - A marker with no text after its stamp ends the run: the line is cleared,
      as MarkdownEditing does for an empty bullet.
    - Otherwise the text after the caret moves to a new line under a fresh
      stamp of the same kind. [-] continues as [?]: a dismissal is something
      done to an existing question, never a new line typed fresh.
    """

    def run(self, edit):
        view = self.view
        for sel in reversed(list(view.sel())):
            line = view.line(sel.b)
            text = view.substr(line)
            m = MARKER_RE.match(text)
            if not m:
                view.replace(edit, sublime.Region(sel.a, sel.b), "\n")
                continue
            indent, bullet, gap, box, _date, rest = m.groups()
            if not rest.strip():
                view.replace(edit, line, "")
                continue
            if box == "-":
                box = "?"
            region = sublime.Region(min(sel.a, sel.b), max(sel.a, sel.b))
            view.replace(edit, region, "\n" + stamp(indent, bullet, gap, box))


class MarkerOpenDirectiveCommand(sublime_plugin.TextCommand):
    """'.' + tab on an otherwise empty line becomes '- [>] YYYYMMDD — '."""

    def run(self, edit):
        view = self.view
        for sel in reversed(list(view.sel())):
            line = view.line(sel.b)
            text = view.substr(line)
            m = re.match(r"^(\s*)\.\s*$", text)
            if not m:
                view.insert(edit, sel.b, "\t")
                continue
            view.replace(edit, line, stamp(m.group(1), "-", " ", ">"))


class MarkerToggleDismissCommand(sublime_plugin.TextCommand):
    """alt+z: '- [?]' <-> '- [-]' on every line holding a caret.

    Never '- [x]': TodoSync reads that as a completed todo and copies it into
    the work log.
    """

    def run(self, edit):
        view = self.view
        rows = set()
        for sel in view.sel():
            for line in view.lines(sel):
                rows.add(line.begin())
        for begin in sorted(rows, reverse=True):
            line = view.line(begin)
            text = view.substr(line)
            m = re.match(r"^([\s>]*[-+*]\s+)\[([?-])\]", text)
            if not m:
                continue
            box_at = line.begin() + len(m.group(1)) + 1
            view.replace(edit, sublime.Region(box_at, box_at + 1),
                         "-" if m.group(2) == "?" else "?")

    def is_enabled(self):
        return self.view.match_selector(0, "text.html.markdown")
