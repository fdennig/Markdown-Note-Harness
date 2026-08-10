"""Plugin-layer tests for WikiResolve.py. The logic lives in the vault now.

    python tests/test_wikiresolve.py

This file covers only what is left in the editor: finding the vault's core,
gating on the frontmatter class, the on_pre_save buffer edit, and the
on_load_async disk write -- including, in both cases, that the view is never
reverted, which is the crash this suite exists to keep from coming back. The
resolution logic itself is tested where it lives, without any stubbing at all:

    <vault>/scripts/tests/test_wikiresolve_core.py

It is worth being explicit that this suite is thin on purpose. The two real
incidents -- a 1,745-note surprise backfill and a crash on every note opened --
both lived in Sublime's plugin lifecycle, which no stub reproduces faithfully.
Moving the batch pass to a script is what addressed them; more tests here would
not have.

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

TMP = tempfile.mkdtemp(prefix="wikiresolve_plugin_")
PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# The plugin imports the core from the vault. Point at the real one: locating and
# importing it is exactly the integration this file exists to check.
VAULT_SCRIPTS = os.path.join(
    os.path.expanduser("~"),
    "OneDrive - United Nations Development Programme", "UNDPnotes", "scripts")

VAULT = os.path.join(TMP, "vault")
os.makedirs(os.path.join(VAULT, "people"))
os.makedirs(os.path.join(VAULT, "emails"))

SETTINGS = {
    "core_path": VAULT_SCRIPTS,
    "person_classes": ["person"],
    "person_roots": [VAULT],
    "exclude": ["*_template.md"],
    "verbose": False,
}


class Settings(object):
    def get(self, k, d=None):
        return SETTINGS.get(k, d)

    def add_on_change(self, *a):
        pass

    def clear_on_change(self, *a):
        pass


class Region(object):
    def __init__(self, a, b):
        self.a, self.b = a, b


class Window(object):
    def __init__(self, root):
        self.root = root
        self.messages = []
        self.panel = None

    def folders(self):
        return [self.root]

    def status_message(self, m):
        self.messages.append(m)

    def show_quick_panel(self, items, cb, *a, **kw):
        self.panel = items

    def open_file(self, p):
        pass

    def find_open_file(self, p):
        return None


class View(object):
    """A Sublime buffer holds LF text with no BOM whatever the file looks like,
    and translates both back on write. The stub does the same, so the encoding
    assertions test WikiResolve rather than the stub."""

    def __init__(self, path, window):
        self.path = path
        self._window = window
        raw = io.open(path, encoding="utf-8", newline="").read()
        self.bom = raw.startswith("﻿")
        self.crlf = "\r\n" in raw
        self.text = raw.lstrip("﻿").replace("\r\n", "\n")
        self.dirty = False
        self.reverted = 0

    def file_name(self):
        return self.path

    def window(self):
        return self._window

    def size(self):
        return len(self.text)

    def substr(self, region):
        return self.text[region.a:region.b]

    def is_dirty(self):
        return self.dirty

    def text_point(self, row, col):
        lines = self.text.splitlines(True)
        return sum(len(l) for l in lines[:row]) + col

    def replace(self, edit, region, s):
        self.text = self.text[:region.a] + s + self.text[region.b:]

    def run_command(self, name, args=None):
        if name == "wiki_resolve_stamp":
            W.WikiResolveStampCommand(self).run(object())
        elif name == "revert":
            self.reverted += 1

    def save(self):
        """What Ctrl+S does: pre-save hooks, then the write."""
        LISTENER.on_pre_save(self)
        out = self.text
        if self.crlf:
            out = out.replace("\n", "\r\n")
        if self.bom:
            out = "﻿" + out
        with io.open(self.path, "w", encoding="utf-8", newline="") as f:
            f.write(out)
        self.dirty = False
        LISTENER.on_post_save_async(self)


sublime = types.ModuleType("sublime")
sublime.Region = Region
sublime.load_settings = lambda name: Settings()
sublime.packages_path = lambda: os.path.join(TMP, "Data", "Packages")
sublime.cache_path = lambda: os.path.join(TMP, "Data", "Cache")
sublime.set_timeout = lambda fn, ms=0: fn()
sublime.set_timeout_async = lambda fn, ms=0: fn()
sublime.active_window = lambda: WINDOW
sublime.status_message = lambda m: WINDOW.messages.append(m)
sys.modules["sublime"] = sublime

sublime_plugin = types.ModuleType("sublime_plugin")


class _WindowBase(object):
    def __init__(self, window=None):
        self.window = window


class _ViewBase(object):
    def __init__(self, view=None):
        self.view = view


sublime_plugin.EventListener = _WindowBase
sublime_plugin.WindowCommand = _WindowBase
sublime_plugin.TextCommand = _ViewBase
sys.modules["sublime_plugin"] = sublime_plugin

sys.path.insert(0, PKG)
import WikiResolve as W  # noqa: E402

FAILS = []


def check(label, cond, extra=""):
    print(("  ok   " if cond else "  FAIL ") + label +
          (("  " + extra) if extra and not cond else ""))
    if not cond:
        FAILS.append(label)


def write(path, text, bom=True, crlf=True):
    if crlf:
        text = text.replace("\n", "\r\n")
    with io.open(path, "w", encoding="utf-8", newline="") as f:
        f.write(("﻿" if bom else "") + text)
    return path


def raw(path):
    with open(path, "rb") as f:
        return f.read()


joan = write(os.path.join(VAULT, "people", "joanChahenza.md"), u"""---
class: person
email:
  - joan.chahenza@undp.org
aliases:
  - "@Joan Chahenza"
---

Joan Chahenza
=============
""")

email1 = write(os.path.join(VAULT, "emails", "Email_2026-01-01_0900 @x - y.md"),
               u"""---
class: email
subject: "y"
from: "[[@Joan Chahenza]]"
from_email: "joan.chahenza@undp.org"
to:
  - "[[@Joan Chahenza]]"
tags:
---

body
""")

topic = write(os.path.join(VAULT, "recoverableGrants.md"), u"""---
class: topic
---

- [ ] a todo
""")

WINDOW = Window(VAULT)
LISTENER = W.WikiResolveListener()

# ==================================================================
print("\n1. finding the vault's core")
if not os.path.isdir(VAULT_SCRIPTS):
    print("  SKIP -- vault scripts dir not present: " + VAULT_SCRIPTS)
    print("\nALL PASS (skipped)")
    shutil.rmtree(TMP, ignore_errors=True)
    sys.exit(0)

check("core is None before anything runs", W.core is None)
loaded = W.load_core(WINDOW)
check("core loaded from core_path", loaded is not None and W.core is not None)
check("it is the real module",
      hasattr(W.core, "analyse") and hasattr(W.core, "LEAF_SPEC"))

print("\n2. the index")
index = W.rebuild(WINDOW)
check("one person found under person_roots",
      sorted(index.people) == ["joanChahenza"], repr(sorted(index.people)))
check("no ambiguity, so no status message",
      not any("ambiguous" in m for m in WINDOW.messages), repr(WINDOW.messages))

print("\n3. the class gate")
check("an email note passes the gate",
      W.head_class(View(email1, WINDOW)) == "email")
check("a topic note does not",
      W.head_class(View(topic, WINDOW)) not in W.core.LEAF_SPEC)
check("a person note is recognised as such",
      W.head_class(View(joan, WINDOW)) == "person")

print("\n4. on_pre_save stamps the buffer")
before = raw(topic)
v = View(topic, WINDOW)
v.save()
check("saving a topic note changes nothing", raw(topic) == before)
check("...and never reverts", v.reverted == 0)

v = View(email1, WINDOW)
v.save()
after = io.open(email1, encoding="utf-8-sig", newline="").read()
check("the sender was resolved", 'from_wiki: "[[joanChahenza]]"' in after)
check("the recipient list was resolved", '  - "[[joanChahenza]]"' in after)
check("the BOM survived the buffer path", raw(email1).startswith(b"\xef\xbb\xbf"))
check("CRLF survived the buffer path",
      raw(email1).count(b"\n") == raw(email1).count(b"\r\n"))
check("the buffer is clean afterwards -- no second save needed", not v.dirty)
check("no revert on the save path (the crash regression)", v.reverted == 0)

second = raw(email1)
View(email1, WINDOW).save()
check("saving again writes the same bytes", raw(email1) == second)

print("\n5. on_load_async stamps a note nobody edits -- and never reverts")
fresh = write(os.path.join(VAULT, "emails", "Email_2026-02-02_0800 @j - hi.md"),
              u"""---
class: email
subject: "hi"
from: "[[@Joan Chahenza]]"
from_email: "joan.chahenza@undp.org"
to:
  - "[[@Joan Chahenza]]"
tags:
---

body
""")
v = View(fresh, WINDOW)
del WINDOW.messages[:]
LISTENER.on_load_async(v)
on_disk = io.open(fresh, encoding="utf-8-sig", newline="").read()
check("opening a new note resolves it on disk",
      'from_wiki: "[[joanChahenza]]"' in on_disk)
# The regression that crashed Sublime on every email note opened: reverting a
# view from inside that view's own load callback. The write is enough.
check("the view was NOT reverted (the crash regression)", v.reverted == 0,
      repr(v.reverted))
check("...and it says what it did, rather than writing silently",
      any("WikiResolve" in m for m in WINDOW.messages), repr(WINDOW.messages))
check("...and the BOM and CRLF survive",
      raw(fresh).startswith(b"\xef\xbb\xbf") and
      raw(fresh).count(b"\n") == raw(fresh).count(b"\r\n"))

settled = raw(fresh)
del WINDOW.messages[:]
v2 = View(fresh, WINDOW)
LISTENER.on_load_async(v2)
check("opening it again writes nothing", raw(fresh) == settled)
check("...and stays quiet when it changed nothing",
      not any("WikiResolve" in m for m in WINDOW.messages), repr(WINDOW.messages))

v3 = View(topic, WINDOW)
before_topic = raw(topic)
LISTENER.on_load_async(v3)
check("opening a topic note does nothing at all",
      raw(topic) == before_topic and v3.reverted == 0)

SETTINGS["on_load"] = False
bumped = write(os.path.join(VAULT, "emails", "Email_2026-03-03_0800 @j - off.md"),
               u"""---
class: email
from: "[[@Joan Chahenza]]"
from_email: "joan.chahenza@undp.org"
to:
tags:
---

body
""")
before_off = raw(bumped)
LISTENER.on_load_async(View(bumped, WINDOW))
check("on_load: false turns the hook off entirely", raw(bumped) == before_off)
SETTINGS["on_load"] = True

print("\n6. a person note save refreshes the index without a full walk")
write(joan, u"""---
class: person
email:
  - joan.chahenza@undp.org
aliases:
  - "@Joan Chahenza"
  - "Joan C"
---

Joan Chahenza
=============
""")
fp_before = W.current_index().fingerprint
View(joan, WINDOW).save()
check("the fingerprint moved", W.current_index().fingerprint != fp_before)
check("the new alias resolves",
      W.current_index().resolve('"[[Joan C]]"') == ("joanChahenza", "t1"),
      repr(W.current_index().resolve('"[[Joan C]]"')))

print("\n7. with no core, every hook is a no-op")
saved_core, W.core = W.core, None
before = raw(email1)
v = View(email1, WINDOW)
v.save()
check("on_pre_save does nothing", raw(email1) == before)
check("head_class returns None", W.head_class(v) is None)
check("rebuild reports the failure rather than raising",
      W.rebuild(WINDOW) is None or W.core is not None)
W.core = saved_core

print("\n" + ("ALL PASS" if not FAILS else "FAILURES: " + ", ".join(FAILS)))
shutil.rmtree(TMP, ignore_errors=True)
sys.exit(1 if FAILS else 0)
