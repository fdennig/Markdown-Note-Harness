"""Reminders -- the Sublime end of `scripts/questions.py --hook`.

On a day with reminders due, the vault's SessionStart hook writes them to
%LOCALAPPDATA%\\ClaudeLogs\\remindersDue.md and, with Sublime already running,
calls

    subl --command 'open_reminders_due {"vault": "C:/.../UNDPnotes"}'

`subl --command` reaches whichever window happens to be active, which may not
be the vault's, and only a command running inside Sublime can see which window
holds which folders. So this one finds the window with the vault among its
folders, opens the file there and brings it forward; failing that, the active
window. (With Sublime closed, the hook launches it on the vault's project
itself and never reaches this.)

Rebuilt 20261112 from questions.py's half of the contract, after the original
was lost with the old machine.

Python 3.3 plugin host: no f-strings.
"""
import os

import sublime
import sublime_plugin


def due_file():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ClaudeLogs", "remindersDue.md")


def norm(path):
    return os.path.normcase(os.path.normpath(path or ""))


def window_holding(vault):
    want = norm(vault)
    if not want:
        return None
    for window in sublime.windows():
        for folder in window.folders():
            f = norm(folder)
            # the vault itself, or a folder inside it
            if f == want or f.startswith(want + os.sep):
                return window
    return None


class OpenRemindersDueCommand(sublime_plugin.ApplicationCommand):

    def run(self, vault=None):
        path = due_file()
        if not os.path.exists(path):
            sublime.status_message("No reminders due file at " + path)
            return
        window = window_holding(vault) or sublime.active_window()
        if window is None:
            sublime.run_command("new_window")
            window = sublime.active_window()
        view = window.open_file(path)
        window.focus_view(view)
        window.bring_to_front()
