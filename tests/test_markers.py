"""Headless checks for Markers.py, Hygiene.py and Reminders.py.

    python tests/test_markers.py

Covers the parts that need no editor: the marker regex, the hygiene summary,
and finding the vault's window. The commands' buffer edits are a manual check
in Sublime. Stubs `sublime` / `sublime_plugin`, so it must stay in tests/.
"""
import os
import sys
import types
import unittest

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PKG)

_windows = []
sublime = types.ModuleType("sublime")
sublime.windows = lambda: _windows
sublime.Region = lambda a, b=None: (a, b)
sublime.load_settings = lambda name: {}
sublime.set_timeout = sublime.set_timeout_async = lambda f, d=0: None
sublime.status_message = lambda m: None
sublime_plugin = types.ModuleType("sublime_plugin")
for name in ("TextCommand", "WindowCommand", "ApplicationCommand", "EventListener"):
    setattr(sublime_plugin, name, type(name, (object,), {}))
sys.modules["sublime"] = sublime
sys.modules["sublime_plugin"] = sublime_plugin

import Hygiene  # noqa: E402
import Markers  # noqa: E402
import Reminders  # noqa: E402


class Window(object):
    def __init__(self, folders):
        self._f = folders

    def folders(self):
        return self._f


class MarkerRe(unittest.TestCase):

    def test_dated_marker(self):
        m = Markers.MARKER_RE.match(u"- [?] 20261112 \u2014 confirm the venue")
        self.assertEqual(m.group(4), "?")
        self.assertEqual(m.group(5), "20261112")
        self.assertEqual(m.group(6), "confirm the venue")

    def test_empty_marker_has_no_text(self):
        m = Markers.MARKER_RE.match(u"  - [>] 20261112 \u2014 ")
        self.assertEqual(m.group(1), "  ")
        self.assertEqual(m.group(6), "")

    def test_todo_is_not_a_marker(self):
        self.assertIsNone(Markers.MARKER_RE.match("- [ ] a todo"))
        self.assertIsNone(Markers.MARKER_RE.match("- [x] done"))

    def test_stamp_shape(self):
        s = Markers.stamp("", "-", " ", ">")
        self.assertRegex(s, u"^- \\[>\\] \\d{8} \u2014 $")


class HygieneSummary(unittest.TestCase):

    def test_nothing_owed(self):
        self.assertEqual(Hygiene.summarise(""), "")

    def test_both_due(self):
        out = ("review: 3 of 30 activity hubs carry unreviewed leaves; oldest "
               "[[x]] watermark 20261001, 40d, 2 waiting. DUE\n"
               "triage: [[adHocActivities]] holds 4 unsorted; watermark "
               "20261009, 34d. DUE\n")
        self.assertEqual(Hygiene.summarise(out),
                         u"hygiene: review DUE \u00b7 triage DUE")

    def test_all_listing_rows_are_not_headlines(self):
        out = "review: 2 activity hubs carry unreviewed leaves, none past threshold. ok\n"
        self.assertEqual(Hygiene.summarise(out), "")


class RemindersWindow(unittest.TestCase):

    def test_finds_window_holding_vault(self):
        vault = os.path.join("C:" + os.sep, "v", "UNDPnotes")
        other, mine = Window([os.path.join("C:" + os.sep, "elsewhere")]), Window([vault])
        _windows[:] = [other, mine]
        self.assertIs(Reminders.window_holding(vault.replace(os.sep, "/")), mine)

    def test_no_match(self):
        _windows[:] = [Window([os.path.join("C:" + os.sep, "elsewhere")])]
        self.assertIsNone(Reminders.window_holding("C:/v/UNDPnotes"))


if __name__ == "__main__":
    unittest.main()
