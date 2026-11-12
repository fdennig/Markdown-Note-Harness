"""Headless tests for EmailBatch.py.

    python tests/test_emailbatch.py

Two halves. The pure helpers -- the text edit that files a note, the merge and
the classification that decide what you are shown -- are tested directly. Then
a whole review is driven through a stub window whose quick panel records its
rows and lets the test "press" one, against a throwaway vault and staging
folder: accept, replace, several links, flags, the bulk discards, mute, undo,
the sweep, waiting on the agent, and clearing the batch.

What it cannot cover is what no stub reproduces: Sublime's own overlay timing
and the transient preview. `proposals.json` is written by hand here, so this
suite does not need the vault; `propose_links.py` has its own, in the vault.

Must stay in tests/ and never move to the package root: it assigns a fake into
sys.modules["sublime"] at import, so Sublime auto-loading it would break every
plugin in the process.
"""

import io
import json
import os
import shutil
import sys
import tempfile
import types

TMP = tempfile.mkdtemp(prefix="emailbatch_")
PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VAULT = os.path.join(TMP, "vault")
MAIL = os.path.join(VAULT, "emails")
ROOT = os.path.join(TMP, "staging")
BATCH = os.path.join(ROOT, "20261112_101500")
for d in (MAIL, os.path.join(VAULT, "scripts"), BATCH):
    os.makedirs(d)
io.open(os.path.join(VAULT, "scripts", "propose_links.py"), "w").close()

SETTINGS = {"staging_root": ROOT, "agent": False, "emails_dir": "emails"}
FAILS = []


def check(label, cond, extra=""):
    print(("  ok   " if cond else "  FAIL ") + label +
          (("  " + str(extra)) if extra and not cond else ""))
    if not cond:
        FAILS.append(label)


# ------------------------------------------------------------------
# Stubs
# ------------------------------------------------------------------

class Settings(object):
    def get(self, k, d=None):
        return SETTINGS.get(k, d)


class View(object):
    def __init__(self, path):
        self.path = path
        self.read_only = False
        self.closed = False

    def file_name(self):
        return self.path

    def set_read_only(self, flag):
        self.read_only = flag

    def set_scratch(self, flag):
        pass

    def close(self):
        self.closed = True


class Window(object):
    def __init__(self):
        self.rows = None
        self.callback = None
        self.highlight = None
        self.selected = None
        self.placeholder = ""
        self.previews = []
        self.commands = []

    def folders(self):
        return [VAULT]

    def show_quick_panel(self, rows, on_select, flags=0, selected=-1,
                         on_highlight=None, placeholder=""):
        self.rows, self.callback, self.highlight = rows, on_select, on_highlight
        self.selected, self.placeholder = selected, placeholder

    def open_file(self, path, flags=0):
        self.previews.append(View(path))
        return self.previews[-1]

    def transient_view_in_group(self, group):
        return self.previews[-1] if self.previews else None

    def active_group(self):
        return 0

    def bring_to_front(self):
        pass

    def run_command(self, name, args=None):
        self.commands.append(name)
        if name == "hide_overlay" and self.callback:
            self.callback(-1)

    # -- what the test presses --
    def labels(self):
        return [row[0] for row in self.rows]

    def press(self, text):
        """Select the first row whose label starts with `text`."""
        for i, label in enumerate(self.labels()):
            if label.startswith(text):
                self.callback(i)
                return
        raise AssertionError("no row starting {0!r} in {1}".format(text, self.labels()[:14]))

    def escape(self):
        self.callback(-1)


WINDOW = Window()
MESSAGES = []
TIMERS = []


def run_timers(rounds=4):
    """Fire what is queued, then what that queued, a bounded number of times.

    Bounded because the wait panel's poll reschedules itself for as long as
    the agent is running -- draining the queue to empty would never return.
    """
    for _ in range(rounds):
        due = TIMERS[:]
        del TIMERS[:]
        for fn in due:
            fn()


sublime = types.ModuleType("sublime")
sublime.TRANSIENT = 4
sublime.load_settings = lambda name: Settings()
sublime.status_message = MESSAGES.append
sublime.windows = lambda: [WINDOW]
sublime.active_window = lambda: WINDOW
sublime.set_timeout = lambda fn, ms=0: TIMERS.append(fn)
sublime_plugin = types.ModuleType("sublime_plugin")


class _Command(object):
    def __init__(self, window=None):
        self.window = window


sublime_plugin.ApplicationCommand = _Command
sublime_plugin.WindowCommand = _Command
sys.modules["sublime"] = sublime
sys.modules["sublime_plugin"] = sublime_plugin
sys.path.insert(0, PKG)
import EmailBatch as eb  # noqa: E402


# ------------------------------------------------------------------
# Pure helpers
# ------------------------------------------------------------------

NOTE = (u"﻿---\r\ntoreply: \r\npersonal: \r\npriority: \r\nclass: email\r\n"
        u'subject: "Thing"\r\nfrom_wiki: "[[ana]]"\r\noutlook:\r\n'
        u"  - entry_id: 00\r\n---\r\n\r\nDear all,\r\nbody\r\n")

print("file_text")
out = eb.file_text(NOTE, ["runTheThing"], {})
check("the link replaces the blank line under the frontmatter",
      u"---\r\n[[runTheThing]]\r\nDear all," in out, repr(out[-60:]))
check("nothing else changes", out.replace(u"[[runTheThing]]", u"") == NOTE)
check("BOM and CRLF survive",
      out.startswith(u"﻿---\r\n") and u"\n" not in out.replace(u"\r\n", u""))
out = eb.file_text(NOTE, ["a", "b"], {})
check("two links share one line", u"---\r\n[[a]] [[b]]\r\nDear all," in out)
check("no link leaves the body alone", eb.file_text(NOTE, [], {}) == NOTE)
out = eb.file_text(NOTE, [], {"toreply": True, "priority": True, "personal": True})
check("flags take the exact values the syntax pins",
      u"\r\ntoreply: true\r\n" in out and u"\r\npriority: high\r\n" in out
      and u"\r\npersonal: true\r\n" in out, out[:90])
check("an unset flag keeps the macro's empty key",
      u"\r\ntoreply: \r\n" in eb.file_text(NOTE, [], {"personal": True}))
tight = NOTE.replace(u"---\r\n\r\nDear", u"---\r\nDear")
out = eb.file_text(tight, ["x"], {})
check("no blank line to replace: the link is inserted, the body kept",
      u"---\r\n[[x]]\r\nDear all," in out)
lf = NOTE.replace(u"\r\n", u"\n").lstrip(u"﻿")
out = eb.file_text(lf, ["x"], {"toreply": True})
check("an LF note with no BOM stays that way",
      u"\r" not in out and not out.startswith(u"﻿") and u"---\n[[x]]\nDear" in out)
check("no frontmatter means no edit", eb.file_text(u"just text\n", ["x"], {}) is None)

print("\nunique_target")
taken = set([os.path.join("d", "a.md"), os.path.join("d", "a (2).md")])
check("a free name is used as is",
      eb.unique_target("d", "b.md", taken.__contains__) == os.path.join("d", "b.md"))
check("a taken name gets the next number",
      eb.unique_target("d", "a.md", taken.__contains__) == os.path.join("d", "a (3).md"))

print("\nmerge_proposals / classify")
order = ["thread", "subject", "agent", "sender"]
base = [{"hub": "s1", "tier": "sender", "why": ""}, {"hub": "t1", "tier": "thread", "why": ""}]
agent = [{"hub": "a1", "tier": "agent", "why": ""}, {"hub": "s1", "tier": "agent", "why": ""}]
merged = eb.merge_proposals(base, agent, ["mine"], order)
check("what you just filed the thread under comes first, then by tier",
      [p["hub"] for p in merged] == ["mine", "t1", "a1", "s1"], merged)
check("a hub offered twice keeps its better tier",
      [p["tier"] for p in merged if p["hub"] == "s1"] == ["agent"])
sender = [{"hub": "s1", "tier": "sender", "why": ""}]
check("a confident tier puts the email in front of you",
      eb.classify(merged, False, False, False, True) == "item")
check("sender-only waits for an agent that is still running",
      eb.classify(sender, False, False, False, True) == "wait")
check("sender-only that the agent found nothing for is swept",
      eb.classify(sender, False, False, True, False) == "sweep")
check("sender-only with no agent at all is shown",
      eb.classify(sender, False, False, False, False) == "item")
check("nothing proposed and no agent is swept",
      eb.classify([], False, False, False, False) == "sweep")
check("a personal hint or a rescue always shows it",
      eb.classify([], True, False, True, False) == "item"
      and eb.classify([], False, True, True, False) == "item")

print("\nis_batch_dir")
check("a dated folder directly under the root", eb.is_batch_dir(BATCH, ROOT))
check("not the root, not a nested folder, not an undated one",
      not eb.is_batch_dir(ROOT, ROOT)
      and not eb.is_batch_dir(os.path.join(BATCH, "filed"), ROOT)
      and not eb.is_batch_dir(os.path.join(ROOT, "notes"), ROOT)
      and not eb.is_batch_dir(VAULT, ROOT))

# ------------------------------------------------------------------
# A whole review
# ------------------------------------------------------------------


def stage(name, sender, addr, subject, thread, proposals=(), personal=False, status="new"):
    fn = name + ".md"
    with io.open(os.path.join(BATCH, fn), "w", encoding="utf-8", newline="") as f:
        f.write(NOTE.replace(u"Thing", subject))
    return fn, {"status": status, "subject": subject, "from": sender,
                "from_wiki": "", "from_email": addr, "date": "2026-11-10 09:00",
                "message_id": "<" + name + ">", "thread": thread, "sent": False,
                "personal_hint": personal,
                "proposals": [{"hub": h, "tier": t, "why": "because"} for h, t in proposals]}


ROWS = [
    stage("t1", "Ana", "ana@x.org", "Thread one", "CONV1", [("runTheThing", "thread")]),
    stage("t2", "Ana", "ana@x.org", "Thread two", "CONV1"),
    stage("m1", "Will", "will@x.org", "Two hubs", "CONV2", [("buildTheOther", "subject")]),
    stage("p1", "Zia", "zia@x.org", "Dinner", "CONV3", personal=True),
    stage("s1", "Sam", "sam@x.org", "Sender only", "CONV4", [("kenya", "sender")]),
    stage("n1", "News", "bot@news.org", "Digest one", "CONV5"),
    stage("n2", "News", "bot@news.org", "Digest two", "CONV6"),
    stage("n3", "Other", "other@x.org", "Unknown", "CONV7"),
    stage("n4", "Spam", "spam@x.org", "Offer", "CONV8"),
    stage("dup", "Ana", "ana@x.org", "Already filed", "CONV9", status="duplicate"),
]
DATA = {"version": 1, "vault": VAULT, "tier_order": order,
        "order": [fn for fn, note in ROWS if note["status"] == "new"],
        "notes": dict(ROWS),
        "hubs": [{"name": n, "class": c, "label": n.title()} for n, c in (
            ("runTheThing", "activity"), ("buildTheOther", "activity"),
            ("kenya", "country"), ("anaPerson", "person"))]}
with io.open(os.path.join(BATCH, eb.PROPOSALS), "w", encoding="utf-8") as f:
    json.dump(DATA, f)


def vault_note(name):
    with io.open(os.path.join(MAIL, name), "r", encoding="utf-8", newline="") as f:
        return f.read()


def decisions():
    with io.open(os.path.join(ROOT, eb.DECISIONS), encoding="utf-8") as f:
        return [json.loads(line) for line in f]


def step(action):
    action()
    run_timers()


print("\nstart")
eb.EmailBatchReviewCommand(WINDOW).run()
run_timers()
s = eb._session
check("a prepared batch starts without running the script", s is not None)
check("duplicates are never queued", "dup.md" not in s.left and len(s.left) == 9)
check("the first email is the first in order", WINDOW.placeholder.startswith(u"1/9"))
check("its proposal is the first row, so enter accepts",
      WINDOW.labels()[0] == "[[runTheThing]]" and WINDOW.selected == 0, WINDOW.labels()[:3])
check("the staged note is previewed, read-only",
      WINDOW.previews[-1].path.endswith("t1.md") and WINDOW.previews[-1].read_only)
check("every hub is listed below, the proposed one only once",
      WINDOW.labels().count("[[runTheThing]]") == 1 and "kenya" in WINDOW.labels()
      and "runTheThing" not in WINDOW.labels())

print("\naccept, and the thread carries it")
step(lambda: WINDOW.callback(0))
check("the note is in the vault with its link",
      u"---\r\n[[runTheThing]]\r\nDear all," in vault_note("t1.md"))
check("the staged copy moved to filed/",
      os.path.isfile(os.path.join(BATCH, "filed", "t1.md"))
      and not os.path.exists(os.path.join(BATCH, "t1.md")))
check("the next email in that thread now proposes the same hub",
      WINDOW.placeholder.startswith(u"2/9") and WINDOW.labels()[0] == "[[runTheThing]]",
      WINDOW.labels()[:3])
step(lambda: WINDOW.press("kenya"))
check("typing another hub replaces the proposal", u"[[kenya]]\r\n" in vault_note("t2.md")
      and u"runTheThing" not in vault_note("t2.md"))
log = decisions()
check("the decision log records what was proposed and what was chosen",
      log[-1]["action"] == "accept" and log[-1]["links"] == ["kenya"]
      and log[-1]["proposed"][0] == {"hub": "runTheThing", "tier": "thread"}, log[-1])

print("\nseveral links, and flags")
check("now on the subject-settled email", WINDOW.labels()[0] == "[[buildTheOther]]")
step(lambda: WINDOW.press("More than one"))
step(lambda: WINDOW.press("[[buildTheOther]]"))
check("a picked link becomes a 'File under' row and nothing is filed yet",
      WINDOW.labels()[0] == "File under [[buildTheOther]]"
      and not os.path.exists(os.path.join(MAIL, "m1.md")), WINDOW.labels()[:2])
step(lambda: WINDOW.press("kenya"))
step(lambda: WINDOW.press("Flag: to reply"))
check("toggling a flag keeps the email and the picks",
      WINDOW.labels()[0] == "File under [[buildTheOther]] [[kenya]]", WINDOW.labels()[:1])
step(lambda: WINDOW.callback(0))
text = vault_note("m1.md")
check("both links on one line, and the flag set",
      u"---\r\n[[buildTheOther]] [[kenya]]\r\nDear" in text and u"\r\ntoreply: true\r\n" in text)

print("\npersonal")
check("a personal hint offers 'file as personal' first",
      WINDOW.labels()[0] == "File as personal, with no link", WINDOW.labels()[:2])
step(lambda: WINDOW.callback(0))
text = vault_note("p1.md")
check("filed with personal: true and no link",
      u"\r\npersonal: true\r\n" in text and u"[[" not in text.split(u"---\r\n", 2)[2])

print("\nundo")
check("with no agent, a sender-only email is shown", WINDOW.labels()[0] == "[[kenya]]"
      and u"Sender only" in WINDOW.placeholder)
step(lambda: WINDOW.press("Undo"))
check("undo takes the note back out of the vault",
      not os.path.exists(os.path.join(MAIL, "p1.md"))
      and os.path.isfile(os.path.join(BATCH, "p1.md")))
check("and puts it straight back in front of you", u"Dinner" in WINDOW.placeholder)
check("the log says so", decisions()[-1]["action"] == "undo")
step(lambda: WINDOW.press("Discard"))
check("discard moves the staged copy aside and writes nothing to the vault",
      os.path.isfile(os.path.join(BATCH, "discarded", "p1.md"))
      and not os.path.exists(os.path.join(MAIL, "p1.md")))

print("\npause and resume")
step(WINDOW.escape)
check("esc pauses and says how to resume", "paused with 5 left" in MESSAGES[-1], MESSAGES[-1])
check("the preview is closed", WINDOW.previews[-1].closed)
eb.EmailBatchReviewCommand(WINDOW).run()
run_timers()
check("review resumes on the same email, undo stack intact",
      eb._session is s and u"Sender only" in WINDOW.placeholder
      and any(l.startswith("Undo") for l in WINDOW.labels()))
step(lambda: WINDOW.callback(0))

print("\nthe sweep")
check("what is left has no hub: one list, bulk discard as its first row",
      WINDOW.labels()[0] == "Discard all 4", WINDOW.labels())
check("sorted by sender", [l.split(u" — ")[0] for l in WINDOW.labels()[1:5]]
      == ["News", "News", "Other", "Spam"])
check("the first EMAIL is preselected, not the bulk row", WINDOW.selected == 1)
WINDOW.highlight(3)
check("highlighting a row previews it", WINDOW.previews[-1].path.endswith("n3.md"))
step(lambda: WINDOW.press(u"Other — Unknown"))
check("selecting one pulls it out to be filed", u"Unknown" in WINDOW.placeholder
      and WINDOW.labels()[0] == "File with no link", WINDOW.labels()[:2])
step(lambda: WINDOW.press("runTheThing"))
check("and it is filed like any other", u"[[runTheThing]]" in vault_note("n3.md"))
check("back in the sweep with the rest", WINDOW.labels()[0] == "Discard all 3")
step(lambda: WINDOW.press(u"News — Digest one"))
check("two from one sender offer a bulk row",
      any(l == "Discard all 2 from News" for l in WINDOW.labels()), WINDOW.labels()[:6])
step(lambda: WINDOW.press("Discard and mute"))
with io.open(os.path.join(ROOT, eb.MUTED), encoding="utf-8") as f:
    muted = f.read()
check("mute writes the address and discards everything pending from it",
      muted.startswith(u"bot@news.org") and sorted(os.listdir(
          os.path.join(BATCH, "discarded"))) == ["n1.md", "n2.md", "p1.md"], muted)
check("one email left in the sweep", WINDOW.labels()[0] == "Discard all 1")
step(lambda: WINDOW.callback(0))

print("\nthe end")
check("nothing left: offer to clear, with the tally",
      WINDOW.labels()[0] == "Clear this batch" and WINDOW.rows[0][1] == "5 filed, 4 discarded",
      WINDOW.rows[0])
step(lambda: WINDOW.press("Undo"))
check("the bulk discard can still be undone", u"Offer" in WINDOW.placeholder
      and os.path.isfile(os.path.join(BATCH, "n4.md")))
step(lambda: WINDOW.press("Discard"))
step(lambda: WINDOW.callback(0))
check("clearing removes the batch folder and nothing else",
      not os.path.exists(BATCH) and os.path.isfile(os.path.join(ROOT, eb.DECISIONS))
      and os.path.isfile(os.path.join(ROOT, eb.MUTED)) and eb._session is None)
check("the vault holds exactly what was accepted",
      sorted(os.listdir(MAIL)) == ["m1.md", "n3.md", "s1.md", "t1.md", "t2.md"],
      sorted(os.listdir(MAIL)))
log = decisions()
check("every decision is in the log",
      [r["action"] for r in log].count("accept") == 6
      and [r["action"] for r in log].count("undo") == 2, [r["action"] for r in log])
check("discards say how they were made",
      set(r.get("via") for r in log if r["action"] == "discard") == set(["item", "mute", "sweep"]))

# ------------------------------------------------------------------
# Waiting on the agent
# ------------------------------------------------------------------

print("\nwaiting on the agent")
os.makedirs(BATCH)
ROWS2 = [stage("w1", "Sam", "sam@x.org", "Unread one", "C1", [("kenya", "sender")]),
         stage("w2", "Ann", "ann@x.org", "Unread two", "C2")]
DATA2 = dict(DATA, order=[fn for fn, _ in ROWS2], notes=dict(ROWS2))
s = eb.Session(WINDOW, BATCH, DATA2)
s.agent_running = True
eb._session = s
step(s.advance)
check("with the agent still reading, nothing unread is shown yet",
      WINDOW.labels()[0] == "Waiting for the agent", WINDOW.labels())
with io.open(os.path.join(BATCH, eb.AGENT_PROPOSALS), "w", encoding="utf-8") as f:
    json.dump({"done": False, "notes": {
        "w1.md": [{"hub": "runTheThing", "tier": "agent", "why": "read it"}]}}, f)
run_timers()                    # the poll fires
check("an answer closes the wait panel without pausing the review",
      "hide_overlay" in WINDOW.commands and "paused" not in MESSAGES[-1], MESSAGES[-1])
check("and the email arrives with the agent's hub ahead of the sender's",
      WINDOW.labels()[:2] == ["[[runTheThing]]", "[[kenya]]"], WINDOW.labels()[:3])
step(lambda: WINDOW.callback(0))
check("the other is still unread, so it waits", WINDOW.labels()[0] == "Waiting for the agent")
step(lambda: WINDOW.press("Go on without it"))
check("'go on without it' sends the unread to the sweep",
      WINDOW.labels()[0] == "Discard all 1", WINDOW.labels())
step(WINDOW.escape)

print("\nclear_finished / next_batch")
other = os.path.join(ROOT, "20261113_090000")
os.makedirs(other)
io.open(os.path.join(other, "x.md"), "w").close()
check("an unprepared batch is found, oldest first",
      eb.next_batch(ROOT) == (BATCH, None) or eb.next_batch(ROOT)[0] == BATCH)
shutil.rmtree(BATCH)
check("then the next one, flagged as needing prepare", eb.next_batch(ROOT) == (other, None))
with io.open(os.path.join(other, eb.PROPOSALS), "w", encoding="utf-8") as f:
    json.dump({"order": ["x.md"], "notes": {"x.md": {"status": "duplicate"}}}, f)
check("a batch with nothing pending is not offered", eb.next_batch(ROOT) == (None, None))
check("and is what clear_finished removes", eb.clear_finished(ROOT) == [other]
      and not os.path.exists(other) and os.path.isfile(os.path.join(ROOT, eb.DECISIONS)))

print("\n" + ("ALL PASS" if not FAILS else "FAILURES: " + ", ".join(FAILS)))
shutil.rmtree(TMP, ignore_errors=True)
sys.exit(1 if FAILS else 0)
