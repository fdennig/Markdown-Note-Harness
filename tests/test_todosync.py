"""Headless harness for TodoSync.py: stubs sublime/sublime_plugin, builds a
throwaway vault, and drives the round trip. Not part of the package."""

import io
import os
import shutil
import sys
import tempfile
import types

TMP = tempfile.mkdtemp(prefix="todosync_")
# Derived, not hardcoded: the package directory has been renamed once already,
# and running "python tests/test_todosync.py" puts tests/ on sys.path, not the
# package root, so a wrong constant here is a hard import failure.
PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------- stubs ----------------

SETTINGS = {
    "classes": ["topic", "activity", "programme"],
    "always_include": [],
    "exclude": ["*_template.md"],
    "overview": "",
    "guard_empty_scope": True,
    "verbose": False,
}


class Settings(object):
    def get(self, k, d=None):
        return SETTINGS.get(k, d)

    def add_on_change(self, *a):
        pass

    def clear_on_change(self, *a):
        pass


class Loc(object):
    def __init__(self, path):
        self.path = path


class Window(object):
    def __init__(self, root, index):
        self.root = root
        self.index = index
        self.messages = []
        self.lookups = 0

    def folders(self):
        return [self.root]

    def symbol_locations(self, sym, *a, **kw):
        self.lookups += 1
        return [Loc(p) for p in self.index.get(sym, [])]

    def find_open_file(self, path):
        return None

    def status_message(self, m):
        self.messages.append(m)

    def show_quick_panel(self, items, cb):
        self.panel = items

    def open_file(self, p):
        pass


class Region(object):
    def __init__(self, a, b):
        self.a, self.b = a, b


class View(object):
    """Just enough of a view for the save listener."""

    def __init__(self, path, window):
        self.path = path
        self._window = window
        with io.open(path, encoding="utf-8", newline="") as f:
            self.text = f.read()

    def file_name(self):
        return self.path

    def window(self):
        return self._window

    def size(self):
        return len(self.text)

    def substr(self, region):
        return self.text[region.a:region.b]

    def is_dirty(self):
        return False


sublime = types.ModuleType("sublime")
sublime.Region = Region
sublime.load_settings = lambda name: Settings()
sublime.packages_path = lambda: os.path.join(TMP, "Data", "Packages")
sublime.set_timeout = lambda fn, ms=0: fn()
sublime.set_timeout_async = lambda fn, ms=0: fn()
sublime.active_window = lambda: WINDOW
sys.modules["sublime"] = sublime

sublime_plugin = types.ModuleType("sublime_plugin")


class _Base(object):
    def __init__(self, window=None):
        self.window = window


sublime_plugin.EventListener = _Base
sublime_plugin.WindowCommand = _Base
sublime_plugin.TextCommand = _Base
sys.modules["sublime_plugin"] = sublime_plugin

sys.path.insert(0, PKG)
import TodoSync as T  # noqa: E402

# ---------------- vault ----------------

VAULT = os.path.join(TMP, "UNDPnotes")
os.makedirs(VAULT)


def write(name, text):
    path = os.path.join(VAULT, name)
    with io.open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    return path


def read(name):
    with io.open(os.path.join(VAULT, name), encoding="utf-8", newline="") as f:
        return f.read()


noteA = write("EnergyForGrowth.md", u"""---
class: programme
---
# Energy for Growth

- [ ] confirm the co-financing letter
- [ ] chase [the annex](https://example.com/a?b=1&c=2) from [[cristina]]
- [x] draft the terms of reference
""")

noteB = write("recoverableGrants.md", u"""---
class: topic
---
- [ ] ! decide the waiver threshold
""")

write("programme_template.md", u"""---
class: programme
---
- [ ] placeholder
""")

write("Email_2026-01-26 @someone - subject.md", u"""---
class: email
---
- [ ] leaked from an email
""")

write("todoUNDP.md", u"""# Capture

- [ ] call the auditor

## Completed (History)

- [x] something old [[francis]] <!-- resolved: 2026-08-01 -->
""")

INDEX = {
    "programme_note": [noteA, os.path.join(VAULT, "programme_template.md")],
    "topic_note": [noteB],
    "activity_note": [],
    "email_note": [os.path.join(VAULT, "Email_2026-01-26 @someone - subject.md")],
}
WINDOW = Window(VAULT, INDEX)

FAILS = []


def check(label, cond, extra=""):
    print(("  ok   " if cond else "  FAIL ") + label + (("  " + extra) if extra and not cond else ""))
    if not cond:
        FAILS.append(label)


def generated():
    lines = T.read_lines(os.path.join(VAULT, "todoUNDP.md"))
    _c, g, _t = T.split_zones(lines)
    return "".join(g)


def tail():
    lines = T.read_lines(os.path.join(VAULT, "todoUNDP.md"))
    _c, _g, t = T.split_zones(lines)
    return "".join(t)


print("\n1. scope")
scope = T.in_scope(WINDOW, os.path.join(VAULT, "todoUNDP.md"))
names = sorted(T.display_name(p) for p in scope.values())
check("3 files in scope", len(scope) == 3, str(names))
check("template excluded", "programme_template" not in names)
check("email excluded", not any(n.startswith("Email_") for n in names))
check("overview always included", "todoUNDP" in names)

print("\n2. first sweep + render")
T.sync(WINDOW)
g = generated()
check("urgent section present", "## Urgent" in g)
check("urgent todo rendered", "- [ ] ! decide the waiver threshold [[recoverableGrants]]" in g)
check("inbox lists external todo",
      "- [ ] confirm the co-financing letter [[EnergyForGrowth]]" in g)
check("by-file groups it", "## EnergyForGrowth" in g and "- [ ] confirm the co-financing letter\n" in g)
check("capture todo absent from inbox", "call the auditor" not in g)
check("checked todo not listed", "terms of reference" not in g)
check("template placeholder absent", "placeholder" not in g)
check("email todo absent", "leaked" not in g)
check("url survives verbatim",
      "chase [the annex](https://example.com/a?b=1&c=2) from [[cristina]] [[EnergyForGrowth]]" in g)
check("history preserved", "something old" in tail())

print("\n3. tick in the overview -> writes into the note")
path = os.path.join(VAULT, "todoUNDP.md")
text = T.read_lines(path)
out = []
for ln in text:
    if ln.startswith("- [ ] confirm the co-financing letter [[EnergyForGrowth]]"):
        ln = ln.replace("- [ ]", "- [x]", 1)
    out.append(ln)
T.write_lines(path, out)
T.apply_overview_edits(WINDOW)
check("source note ticked", "- [x] confirm the co-financing letter" in read("EnergyForGrowth.md"))
check("gone from generated zone", "confirm the co-financing letter" not in generated())
check("landed in history", "- [x] confirm the co-financing letter [[EnergyForGrowth]] <!-- resolved:" in tail())

print("\n4. tick in the note -> reaches history with no matching")
src = read("recoverableGrants.md").replace("- [ ] ! decide", "- [x] ! decide")
write("recoverableGrants.md", src)
T.sync(WINDOW)
check("gone from generated zone", "waiver threshold" not in generated())
check("landed in history", "- [x] ! decide the waiver threshold [[recoverableGrants]]" in tail())

print("\n5. urgency toggled from the overview")
path = os.path.join(VAULT, "todoUNDP.md")
out = []
for ln in T.read_lines(path):
    if ln.startswith("- [ ] chase [the annex]"):
        ln = ln.replace("- [ ] chase", "- [ ] ! chase", 1)
    out.append(ln)
T.write_lines(path, out)
T.apply_overview_edits(WINDOW)
check("! written into the note", "- [ ] ! chase [the annex]" in read("EnergyForGrowth.md"))
check("rendered as urgent", "## Urgent" in generated() and "! chase [the annex]" in generated())

print("\n6. urgency removed again")
out = []
for ln in T.read_lines(path):
    if ln.strip().startswith("- [ ] ! chase [the annex]"):
        ln = ln.replace("- [ ] ! chase", "- [ ] chase", 1)
    out.append(ln)
T.write_lines(path, out)
T.apply_overview_edits(WINDOW)
check("! removed from the note", "- [ ] chase [the annex]" in read("EnergyForGrowth.md"))

print("\n7. deleting a line in a note removes it silently")
before_tail = tail()
src = "".join(l for l in io.open(os.path.join(VAULT, "EnergyForGrowth.md"), encoding="utf-8")
              if "chase [the annex]" not in l)
write("EnergyForGrowth.md", src)
T.sync(WINDOW)
check("gone from overview", "the annex" not in generated())
check("not added to history", tail().count("the annex") == before_tail.count("the annex"))

print("\n8. capture zone round trip")
path = os.path.join(VAULT, "todoUNDP.md")
out = []
for ln in T.read_lines(path):
    if ln.startswith("- [ ] call the auditor"):
        ln = "- [ ] ! call the auditor\n"
    out.append(ln)
T.write_lines(path, out)
T.sync(WINDOW)
check("capture urgent reaches Urgent", "- [ ] ! call the auditor\n" in generated())
check("still not in By File", "## todoUNDP" not in generated())

out = []
for ln in T.read_lines(path):
    if ln.startswith("- [ ] ! call the auditor") and "[[" not in ln:
        ln = "- [x] ! call the auditor\n"
    out.append(ln)
T.write_lines(path, out)
T.apply_overview_edits(WINDOW)
check("capture tick reaches history", "call the auditor" in tail())

print("\n9. a cold index does not blank the overview")
saved = dict(INDEX)
INDEX["programme_note"] = []
INDEX["topic_note"] = []
INDEX["activity_note"] = []
before = generated()
T.sync(WINDOW)
check("todos survive the index going dark", generated() == before,
      "cache-recorded classes should hold scope open")

print("\n9b. cold index AND no cache: refuses rather than rendering nothing")
os.remove(os.path.join(VAULT, "todoUNDP.cache.json"))
T.load_cache(os.path.join(VAULT, "todoUNDP.md"))  # confirm it is gone
T.sync(WINDOW)          # nothing cached, nothing indexed -> empty generated zone
check("empty render is allowed when nothing was ever known",
      "EnergyForGrowth" not in generated())
INDEX.update(saved)
T.sync(WINDOW)          # index back: everything returns
check("recovers once the index returns", generated() == before)

INDEX["programme_note"] = []
INDEX["topic_note"] = []
INDEX["activity_note"] = []
T.sync(WINDOW)
check("cache-recorded classes alone hold scope open", generated() == before)

print("\n9c. last resort: neither the index nor the cache can vouch for a note")
import json as _json
CACHE = os.path.join(VAULT, "todoUNDP.cache.json")
with io.open(CACHE, encoding="utf-8") as f:
    blob = _json.load(f)
for rec in blob["files"].values():          # a cache from before classes were recorded
    rec["class"] = None
with io.open(CACHE, "w", encoding="utf-8") as f:
    f.write(_json.dumps(blob, ensure_ascii=False))
WINDOW.messages = []
T.sync(WINDOW)
check("overview left intact", generated() == before)
check("and says so", any("cold" in m for m in WINDOW.messages), str(WINDOW.messages))

INDEX.update(saved)
T.sync(WINDOW)

print("\n10. stat gating")
cache, completed, aborted = T.sweep(WINDOW, os.path.join(VAULT, "todoUNDP.md"))
check("second sweep reads nothing new", not completed and not aborted)

print("\n11. ambiguity is reported, not guessed")
sub = os.path.join(VAULT, "sub")
os.makedirs(sub)
dupe = os.path.join(sub, "EnergyForGrowth.md")
with io.open(dupe, "w", encoding="utf-8", newline="") as f:
    f.write(u"---\nclass: programme\n---\n- [ ] shared text\n")
with io.open(os.path.join(VAULT, "EnergyForGrowth.md"), "a", encoding="utf-8", newline="") as f:
    f.write(u"- [ ] shared text\n")
INDEX["programme_note"] = [os.path.join(VAULT, "EnergyForGrowth.md"), dupe]
T.sync(WINDOW)
out = []
for ln in T.read_lines(path):
    if ln.startswith("- [ ] shared text [[EnergyForGrowth]]"):
        ln = ln.replace("- [ ]", "- [x]", 1)
    out.append(ln)
T.write_lines(path, out)
T.apply_overview_edits(WINDOW)
check("neither note was ticked",
      "- [x] shared text" not in read("EnergyForGrowth.md")
      and "- [x] shared text" not in io.open(dupe, encoding="utf-8").read())
check("reported to the user", any("not found" in m for m in WINDOW.messages), str(WINDOW.messages[-1:]))

print("\n12. CRLF notes stay CRLF")
crlf = os.path.join(VAULT, "crlf.md")
with io.open(crlf, "w", encoding="utf-8", newline="") as f:
    f.write(u"---\r\nclass: topic\r\n---\r\n- [ ] windows line endings\r\n")
INDEX["topic_note"] = [crlf]
T.sync(WINDOW)
T.edit_source(crlf, 3, "windows line endings", check=True)
raw = io.open(crlf, "rb").read()
check("no LF-only lines introduced", raw.count(b"\n") == raw.count(b"\r\n"))
check("box ticked", b"- [x] windows line endings" in raw)

print("\n13. repeated sweeps change nothing")
import hashlib, re as _re


def fingerprint():
    s = io.open(os.path.join(VAULT, "todoUNDP.md"), encoding="utf-8", newline="").read()
    return hashlib.sha256(_re.sub(r"^_Updated .*$", "", s, flags=_re.M).encode("utf-8")).hexdigest()


T.sync(WINDOW)          # settle: step 12 ticked a box, which is a real change
a = fingerprint()
T.sync(WINDOW)
b = fingerprint()
T.sync(WINDOW)
c = fingerprint()
check("sweep is idempotent", a == b == c)

print("\n14. a save outside the swept classes does nothing at all")
listener = T.TodoSyncListener()
OV = os.path.join(VAULT, "todoUNDP.md")


def probe(name):
    """(overview mtime, cache mtime, index lookups so far)"""
    return (os.stat(OV).st_mtime,
            os.stat(os.path.join(VAULT, "todoUNDP.cache.json")).st_mtime,
            WINDOW.lookups)


T.sync(WINDOW)
email = os.path.join(VAULT, "Email_2026-01-26 @someone - subject.md")
for label, path in [("email note", email),
                    ("template", os.path.join(VAULT, "programme_template.md"))]:
    before = probe(label)
    listener.on_post_save_async(View(path, WINDOW))
    check("saving a {} touches nothing".format(label), probe(label) == before,
          "{} -> {}".format(before, probe(label)))

unclassified = write("scratch.md", u"# scratch\n\n- [ ] not mine\n")
before = probe("unclassified")
listener.on_post_save_async(View(unclassified, WINDOW))
check("saving an unclassified note touches nothing", probe("unclassified") == before)
check("and it never reaches the overview", "not mine" not in generated())

print("\n15. an in-scope save with no todo change writes nothing")
nochange = os.path.join(VAULT, "recoverableGrants.md")
write("recoverableGrants.md", read("recoverableGrants.md") + u"\nsome prose.\n")
before = probe("in-scope")
listener.on_post_save_async(View(nochange, WINDOW))
after = probe("in-scope")
check("cache refreshed", after[1] != before[1])
check("overview left alone", after[0] == before[0])

print("\n16. a note classified for the first time needs no index")
fresh = write("brandNew.md", u"---\nclass: activity\n---\n- [ ] freshly classified\n")
before_gen = generated()
listener.on_post_save_async(View(fresh, WINDOW))
check("swept on that very save", "freshly classified" in generated(),
      "index still does not list it")
check("overview was rewritten", generated() != before_gen)

print("\n17. an in-scope save that does change a todo does write")
write("recoverableGrants.md", read("recoverableGrants.md") + u"- [ ] a new one\n")
before = probe("changed")
listener.on_post_save_async(View(nochange, WINDOW))
check("overview rewritten", probe("changed")[0] != before[0])
check("todo present", "a new one" in generated())

print("\n" + ("ALL PASS" if not FAILS else "FAILURES: " + ", ".join(FAILS)))
shutil.rmtree(TMP, ignore_errors=True)
sys.exit(1 if FAILS else 0)
