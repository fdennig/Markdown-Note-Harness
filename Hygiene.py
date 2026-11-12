"""Hygiene -- a persistent status-bar item when a `review` or `triage` is owed.

The daily counterpart to the vault's SessionStart hook. That hook only fires
when a Claude Code session starts, which is no use for "chase me when I have
gone quiet"; this runs `scripts/hygiene.py` once a day from inside the editor,
and while anything is DUE every view's status bar says so:

    hygiene: review DUE · triage DUE

It clears itself on the first run that finds nothing owed. The status is
per-view in Sublime, so it is set on every open view and again on each view
as it is activated -- that is what makes it persistent.

The vault is found the way WikiResolve finds its core: a folder open in some
window that contains `scripts/hygiene.py`. The interpreter is the uv one the
hooks use, overridable in Hygiene.sublime-settings. `hygiene.py` never
writes, so running it is always safe.

Commands:
  hygiene_check   run now, whatever the date
  hygiene_show    run now and show the full picture (`--all`) in a panel

Rebuilt 20261112 from the vault's description after the original was lost with
the old machine.

Python 3.3 plugin host: no f-strings; subprocess.communicate(timeout=...) and
threading are fine.
"""
import os
import subprocess
import threading
import time

import sublime
import sublime_plugin

KEY = "hygiene"
SETTINGS = "Hygiene.sublime-settings"
DEFAULT_PYTHON = os.path.join(os.path.expanduser("~"), ".local", "bin",
                              "python3.14.exe")

_state = {"text": "", "ran_on": None, "running": False}
_lock = threading.Lock()


def settings():
    return sublime.load_settings(SETTINGS)


def python_path():
    p = settings().get("python") or DEFAULT_PYTHON
    return os.path.expandvars(os.path.expanduser(p))


def find_vault():
    for window in sublime.windows():
        for folder in window.folders():
            if os.path.isfile(os.path.join(folder, "scripts", "hygiene.py")):
                return folder
    return None


def run_hygiene(vault, show_all=False):
    """(returncode, stdout) -- never raises."""
    args = [python_path(), os.path.join(vault, "scripts", "hygiene.py")]
    if show_all:
        args.append("--all")
    startupinfo = None
    if hasattr(subprocess, "STARTUPINFO"):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    try:
        proc = subprocess.Popen(args, cwd=vault, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT,
                                stdin=subprocess.PIPE, startupinfo=startupinfo)
        out, _ = proc.communicate(timeout=settings().get("timeout", 120))
        return proc.returncode, out.decode("utf-8", "replace")
    except Exception as e:  # missing interpreter, timeout, anything
        return -1, "hygiene: could not run {0}: {1}".format(args[0], e)


def summarise(output):
    """'review DUE · triage DUE' from the script's DUE lines, or ''."""
    due = []
    for line in output.splitlines():
        line = line.strip()
        if line.endswith("DUE") and ":" in line:
            due.append(line.split(":", 1)[0] + " DUE")
    return "hygiene: " + u" · ".join(due) if due else ""


def paint(text):
    for window in sublime.windows():
        for view in window.views():
            if text:
                view.set_status(KEY, text)
            else:
                view.erase_status(KEY)


def check(force=False, show_all=False, window=None):
    today = time.strftime("%Y%m%d")
    with _lock:
        if _state["running"] or (not force and _state["ran_on"] == today):
            return
        _state["running"] = True
    try:
        vault = find_vault()
        if not vault:
            return  # no vault open yet; try again at the next tick
        code, out = run_hygiene(vault, show_all)
        if code != 0:
            print(out)
            sublime.status_message("Hygiene: run failed -- see console")
            return
        _state["ran_on"] = today
        _state["text"] = summarise(out)
        sublime.set_timeout(lambda: paint(_state["text"]), 0)
        if show_all and window is not None:
            sublime.set_timeout(lambda: show_panel(window, out), 0)
    finally:
        _state["running"] = False


def show_panel(window, text):
    panel = window.create_output_panel("hygiene")
    panel.run_command("append", {"characters": text or "nothing owed\n"})
    window.run_command("show_panel", {"panel": "output.hygiene"})


def tick():
    sublime.set_timeout_async(check, 0)
    minutes = settings().get("check_every_minutes", 60)
    sublime.set_timeout(tick, int(minutes) * 60 * 1000)


def plugin_loaded():
    # let the windows open their folders first
    sublime.set_timeout(tick, 15000)


def plugin_unloaded():
    paint("")


class HygieneListener(sublime_plugin.EventListener):

    def on_activated(self, view):
        if _state["text"]:
            view.set_status(KEY, _state["text"])
        else:
            view.erase_status(KEY)


class HygieneCheckCommand(sublime_plugin.WindowCommand):

    def run(self):
        sublime.set_timeout_async(lambda: check(force=True), 0)


class HygieneShowCommand(sublime_plugin.WindowCommand):

    def run(self):
        window = self.window
        sublime.set_timeout_async(
            lambda: check(force=True, show_all=True, window=window), 0)
