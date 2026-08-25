"""TodoSync -- mirrors unchecked todos from classified notes into todoUNDP.md.

The notes are the source of truth. Every pass re-derives the overview from
them; nothing about *which todos exist* is stored anywhere else. The previous
implementation accumulated todos into todo.yml one saved file at a time, which
meant it could neither see the vault nor forget a deleted line.

    scope    window.symbol_locations("<class>_note") for each configured class
    sweep    os.stat each path; read only what changed; one regex per file
    diff     against the cache -> what was ticked in its own note
    render   rewrite only the generated zone of todoUNDP.md

and in reverse, when todoUNDP.md is saved: ticks and "!" toggles inside the
generated zone are written back into the source notes, then the whole thing is
swept and rendered again.

Two gates keep this cheap enough to run on every save. A saved file is ignored
outright unless its own frontmatter class is one we sweep (see watched()), so
saving an email note costs a single regex. And the overview is only written
when the rendered content actually differs, so a sweep that finds nothing new
touches no files at all.

The cache (todoUNDP.cache.json) is derived, never authoritative -- delete it
and the next sweep rebuilds it identically. It exists to gate reads on
(mtime, size), and to locate the source line behind an overview line without
putting a path into the wiki link.

Runs on the 3.3 plugin host: str.format(), no f-strings.
"""

import sublime
import sublime_plugin

import datetime
import fnmatch
import json
import os
import re

SETTINGS_FILE = "TodoSync.sublime-settings"
OVERVIEW_NAME = "todoUNDP.md"
# The work log. Deliberately a second file: it is the one authoritative,
# never-re-derived thing in a system whose whole premise is that the notes are
# the source of truth, and at 267 rows it had grown to half of the overview a
# person actually reads. Being irreplaceable is a reason to keep it, not a
# reason to keep it underfoot.
DONE_NAME = "todoUNDPdone.md"
CACHE_NAME = "todoUNDP.cache.json"

# The generated zone is delimited by markers rather than by heading names, so
# the generator can rewrite one region and copy the others through untouched.
BEGIN_MARK = "<!-- todosync:begin -- generated below, edits here are overwritten -->"
END_MARK = "<!-- todosync:end -->"
BEGIN_RE = re.compile(r"^<!--\s*todosync:begin\b.*-->\s*$")
END_RE = re.compile(r"^<!--\s*todosync:end\b.*-->\s*$")
COMPLETED_RE = re.compile(r"^#{1,6}\s+completed\b", re.I)
# A month section in the work log.
MONTH_RE = re.compile(r"^##[ \t]+(\d{4}-\d{2})[ \t]*$")

# A checkbox line. Group 3 is the box, group 4 the text.
TODO_RE = re.compile(r"^([ \t]*)([-*+])[ \t]+\[([ xX])\][ \t]+(\S.*?)[ \t]*$")
# Urgency is a leading "!", held as a flag rather than as part of the text --
# so toggling it never changes a todo's identity.
URGENT_RE = re.compile(r"^!\s+(.*)$")
CLASS_RE = re.compile(r"^class\s*:\s*([A-Za-z0-9_-]+)\s*$", re.M)
# The vault owner marks their own person note. Kept in the note rather than
# in settings so the fact travels with the vault, as WikiResolve does it.
SELF_RE = re.compile(r"^self\s*:\s*true\s*$", re.M | re.I)
# The assignee slot: a wiki link opening the todo, followed by a colon.
LEAD_RE = re.compile(r"^\[\[([^\[\]]+)\]\]\s*:\s*(.*)$")
WIKI_RE = re.compile(r"\[\[([^\[\]]+)\]\]")
# Frontmatter only, so a 'class:' in a code block cannot pull a note into scope.
FRONTMATTER_RE = re.compile(r"\A---[ \t]*\r?\n(.*?)^(?:---|\.\.\.)[ \t]*\r?\n",
                            re.S | re.M)
# The "last updated" line is excluded when deciding whether the overview
# changed -- otherwise it differs on every render and forces a pointless write.
STAMP_RE = re.compile(r"^_Updated .*$", re.M)

# An overview line: "- [x] ! text [[Source]]". Greedy (.*) so the LAST wiki
# link is taken as the source and links inside the todo text survive.
LINE_RE = re.compile(
    r"^[ \t]*[-*+][ \t]+\[([ xX])\][ \t]+(.*?)(?:[ \t]+\[\[([^\]]*)\]\])?[ \t]*$")
HEADING_RE = re.compile(r"^(#{1,6})[ \t]+(.*?)[ \t]*$")


# ------------------------------------------------------------------
# Settings, logging, small helpers
# ------------------------------------------------------------------

def settings():
    return sublime.load_settings(SETTINGS_FILE)


def log(message):
    """Append to Data/Logs/TodoSync.log. Silent if logging itself fails."""
    if not settings().get("verbose", True):
        return
    try:
        log_dir = os.path.join(os.path.dirname(sublime.packages_path()), "Logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(os.path.join(log_dir, "TodoSync.log"), "a", encoding="utf-8") as f:
            f.write("[{}] {}\n".format(stamp, message))
    except Exception as e:
        print("TodoSync: could not write log: {}".format(e))


def key(path):
    """Cache key. normcase so Windows' casing cannot split one file in two."""
    return os.path.normcase(os.path.abspath(path))


def display_name(path):
    """The name a wiki link uses: basename without extension."""
    return os.path.splitext(os.path.basename(path))[0]


def today():
    return datetime.date.today().isoformat()


def read_lines(path):
    """Lines *with* their terminators, so rewriting one line cannot convert
    the rest of a CRLF file to LF behind the user's back."""
    with open(path, "r", encoding="utf-8", newline="") as f:
        return f.read().splitlines(True)


def write_lines(path, lines):
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.writelines(lines)


def bare(line):
    return line.rstrip("\r\n")


def eol(lines, fallback="\n"):
    """The terminator this file already uses."""
    for line in lines:
        if line.endswith("\r\n"):
            return "\r\n"
        if line.endswith("\n"):
            return "\n"
    return fallback


def frontmatter_class(text):
    """The note's class, read from its frontmatter block. None if unclassified."""
    block = FRONTMATTER_RE.match(text)
    if not block:
        return None
    found = CLASS_RE.search(block.group(1))
    return found.group(1) if found else None


def frontmatter_self(text):
    """True when this is the vault owner's own person note."""
    block = FRONTMATTER_RE.match(text)
    return bool(block and SELF_RE.search(block.group(1)))


def excluded(path):
    name = os.path.basename(path)
    return any(fnmatch.fnmatch(name, p) or fnmatch.fnmatch(path, p)
               for p in settings().get("exclude", []))


# ------------------------------------------------------------------
# Locations
# ------------------------------------------------------------------

def overview_path(window):
    configured = (settings().get("overview") or "").strip()
    if configured:
        return os.path.abspath(os.path.expanduser(os.path.expandvars(configured)))
    for folder in window.folders():
        candidate = os.path.join(folder, OVERVIEW_NAME)
        if os.path.exists(candidate):
            return os.path.abspath(candidate)
    return None


def done_path(overview):
    return os.path.join(os.path.dirname(overview), DONE_NAME)


def month(stamp):
    return stamp[:7]


def cache_path(overview):
    return os.path.join(os.path.dirname(overview), CACHE_NAME)


def load_cache(overview):
    path = cache_path(overview)
    if not os.path.exists(path):
        return {"version": 3, "files": {}}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("version") != 3 or "files" not in data:
            return {"version": 3, "files": {}}
        return data
    except Exception as e:
        log("cache unreadable, starting empty: {}".format(e))
        return {"version": 3, "files": {}}


def save_cache(overview, cache):
    try:
        with open(cache_path(overview), "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=1, ensure_ascii=False)
    except Exception as e:
        log("ERROR saving cache: {}".format(e))


# ------------------------------------------------------------------
# Scope -- which notes get swept
# ------------------------------------------------------------------

# Cache keys of the notes the last sweep covered. Lets on_post_save_async
# decide whether a saved file can possibly matter without touching the index.
_scope = set()


def in_scope(window, overview):
    """Absolute paths to sweep. Resolved entirely from the symbol index and
    the settings; no file is opened and no directory is walked."""
    conf = settings()
    found = {}

    for cls in conf.get("classes", []):
        for loc in window.symbol_locations(cls + "_note"):
            found[key(loc.path)] = os.path.abspath(loc.path)

    base = os.path.dirname(overview)
    for entry in conf.get("always_include", []):
        path = entry if os.path.isabs(entry) else os.path.join(base, entry)
        if os.path.exists(path):
            found[key(path)] = os.path.abspath(path)

    # The overview's capture zone is a real note and is always swept.
    found[key(overview)] = overview

    for k in list(found):
        if k != key(overview) and excluded(found[k]):
            del found[k]

    return found


# ------------------------------------------------------------------
# Sweep
# ------------------------------------------------------------------

def parse(lines, stop_at_marker=False):
    """Split checkbox lines into unchecked and checked entries.

    stop_at_marker is for the overview itself: only its capture zone is a
    source of todos. Everything from the generated marker down is either this
    plugin's own output or the completed history."""
    unchecked, checked = [], []
    for row, raw in enumerate(lines):
        line = bare(raw)
        if stop_at_marker and BEGIN_RE.match(line):
            break
        m = TODO_RE.match(line)
        if not m:
            continue
        text = m.group(4)
        urgent = False
        um = URGENT_RE.match(text)
        if um:
            urgent, text = True, um.group(1).strip()
        if not text:
            continue
        if m.group(3) == " ":
            unchecked.append({"text": text, "row": row, "urgent": urgent})
        else:
            checked.append({"text": text, "row": row})
    return unchecked, checked


def scan_file(path, is_overview):
    lines = read_lines(path)
    unchecked, checked = parse(lines, stop_at_marker=is_overview)
    head = "".join(lines[:40])
    return {
        "path": path,
        "class": frontmatter_class(head),
        "self": frontmatter_self(head),
        "unchecked": unchecked,
        "checked": checked,
    }


def sweep(window, overview, force=False, extra=None):
    """Refresh the cache. Returns (cache, completed, aborted).

    completed is what moved from unchecked to checked in its own note since
    the last sweep -- the half of the round trip that needs no matching at all.

    extra is a path the caller knows qualifies but the index may not list yet.
    """
    global _scope
    cache = load_cache(overview)
    old = cache["files"]
    scope = in_scope(window, overview)

    # The index is not the only witness. A note already recorded with a
    # qualifying class stays in scope even while the index is catching up or
    # cold; if its class is genuinely gone, the re-scan below records that and
    # the next sweep drops it. Together with `extra`, this is what lets a note
    # classified for the first time show up on the save that classified it.
    classes = settings().get("classes", [])
    for k, record in old.items():
        path = record.get("path")
        if k not in scope and record.get("class") in classes and path \
                and not excluded(path) and os.path.exists(path):
            scope[k] = path
    if extra and key(extra) not in scope and not excluded(extra):
        scope[key(extra)] = os.path.abspath(extra)

    _scope = set(scope)

    # Never build a generated zone out of nothing. The retention above already
    # survives a cold index whenever the cache has anything in it, so this can
    # only trip on a genuinely cold start: no cache, and an index still
    # building. Deliberately *not* a proportional test -- narrowing "classes"
    # is supposed to shrink the scope, and a ratio would read that as a fault.
    notes = [k for k in scope if k != key(overview)]
    remembered = [k for k in old if k != key(overview)]
    if settings().get("guard_empty_scope", True) and remembered and not notes:
        log("ABORT: scope empty but cache held {} notes -- index looks cold"
            .format(len(remembered)))
        return cache, [], True

    fresh = {}
    completed = []
    reads = 0

    for k, path in scope.items():
        try:
            st = os.stat(path)
        except OSError:
            continue

        prev = old.get(k)
        if not force and prev and prev.get("mtime") == st.st_mtime \
                and prev.get("size") == st.st_size:
            fresh[k] = prev
            continue

        try:
            record = scan_file(path, k == key(overview))
        except Exception as e:
            log("ERROR reading {}: {}".format(path, e))
            if prev:
                fresh[k] = prev
            continue

        reads += 1
        record["mtime"] = st.st_mtime
        record["size"] = st.st_size
        fresh[k] = record

        if prev:
            was_open = dict((e["text"], e) for e in prev.get("unchecked", []))
            for entry in record["checked"]:
                if entry["text"] in was_open:
                    completed.append({
                        "text": entry["text"],
                        "urgent": was_open[entry["text"]].get("urgent", False),
                        "source": display_name(path),
                    })

    cache["files"] = fresh
    save_cache(overview, cache)
    log("sweep: {} in scope, {} read, {} completed in their notes"
        .format(len(scope), reads, len(completed)))
    return cache, completed, False


# ------------------------------------------------------------------
# Render
# ------------------------------------------------------------------

def split_zones(lines):
    """(capture, generated, tail).

    Falls back for a file written before the markers existed: everything above
    a '## Completed' heading becomes the capture zone and the heading and below
    become the tail, so the first render migrates the file in place."""
    begin = end = None
    for i, raw in enumerate(lines):
        line = bare(raw)
        if begin is None:
            if BEGIN_RE.match(line):
                begin = i
        elif END_RE.match(line):
            end = i
            break
    if begin is not None and end is not None:
        return lines[:begin], lines[begin + 1:end], lines[end + 1:]

    for i, raw in enumerate(lines):
        if COMPLETED_RE.match(bare(raw)):
            return lines[:i], [], lines[i:]
    return lines, [], []


def collect(cache, overview):
    """Every open todo, as a dict. Carries the source note's class, because
    that is half of what decides whose todo it is."""
    ov = key(overview)
    items = []
    for k, record in cache["files"].items():
        path = record.get("path", k)
        name = display_name(path)
        for entry in record.get("unchecked", []):
            items.append({
                "source": name,
                "path": path,
                "row": entry["row"],
                "text": entry["text"],
                "urgent": entry.get("urgent", False),
                "capture": k == ov,
                "class": record.get("class"),
            })
    items.sort(key=lambda i: (i["source"].lower(), i["row"]))
    return items


def people_index(cache):
    """(person basenames, the vault owner's name) from the notes already swept.

    No index call and no extra read: the sweep already recorded every note's
    class, and francis.md already declares `self: true` for WikiResolve."""
    persons, me = set(), None
    for record in cache["files"].values():
        path = record.get("path")
        if not path or record.get("class") != "person":
            continue
        name = display_name(path)
        persons.add(name)
        if record.get("self"):
            me = name
    return persons, me


def assignee(item, persons, me):
    """Whose todo this is: the leading [[person]]: slot, or failing that the
    person note it was written in. None for the vault owner's own todos -- they
    are the Inbox -- and for anything naming nobody."""
    m = LEAD_RE.match(item["text"])
    if m and m.group(1) in persons:
        who = m.group(1)
    elif item["class"] == "person" and item["source"] in persons:
        who = item["source"]
    else:
        return None
    return None if who == me else who


def by_person(items, persons, me):
    """(assigned, mentioned), keyed by person.

    A mention is a person named inside somebody else's todo. It is listed only
    under a person who already has work of their own, so a name you owe nothing
    to never opens a heading of its own."""
    assigned, mentioned = {}, {}
    for item in items:
        who = assignee(item, persons, me)
        if who:
            assigned.setdefault(who, []).append(item)
        for name in WIKI_RE.findall(item["text"]):
            if name in persons and name != me and name != who                     and item not in mentioned.setdefault(name, []):
                mentioned[name].append(item)
    for name in list(mentioned):
        if name not in assigned or not mentioned[name]:
            del mentioned[name]
    return assigned, mentioned


def render_line(text, urgent, source=None):
    line = "- [ ] {}{}".format("! " if urgent else "", text)
    if source:
        line += " [[{}]]".format(source)
    return line


def person_line(item, person):
    """Under a person's own heading the leading [[person]]: slot is repetition
    of the heading, so it comes off. The trailing source link stays: it is not
    decoration, it is what the write-back resolves a tick against.

    A mention row keeps its slot, because there it names somebody else -- who
    is on the hook is the useful half of a row filed under a third party."""
    text = item["text"]
    m = LEAD_RE.match(text)
    if m and m.group(1) == person and m.group(2).strip():
        text = m.group(2).strip()
    return render_line(text, item["urgent"], item["source"])


def build_generated(items, nl, persons=(), me=None):
    """The generated zone: ## Urgent and ## Inbox close the hand-written
    # Aggregate section above the marker, then # By Person regroups the lot.

    There is deliberately no "by file" section. The source hub of every row is
    named by its trailing wiki link, and a hub's own todos are legible in the
    hub itself -- so grouping by file here only reprinted what one hop away
    already says. Grouping by person is the view no single note can give."""
    out = []

    def put(s=""):
        out.append(s + nl)

    # Stamped on every render but ignored when deciding whether anything
    # changed, so it reads as "when the todos last moved", not "last swept".
    put("_Updated {}_".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    put()

    # Urgent cross-cuts everything, the capture zone included.
    urgent = [i for i in items if i["urgent"]]
    if urgent:
        put("## Urgent")
        put()
        for item in urgent:
            put(render_line(item["text"], True, item["source"]))
        put()

    # Inbox skips capture-zone todos -- they are already legible a few lines up
    # under ## Inline capture.
    put("## Inbox")
    put()
    normal = [i for i in items if not i["urgent"] and not i["capture"]]
    if not normal:
        put("_No open todos._")
    else:
        for item in normal:
            put(render_line(item["text"], False, item["source"]))
    put()

    put("# By Person")
    put()
    assigned, mentioned = by_person(items, persons, me)
    if not assigned and not mentioned:
        put("_Nobody named._")
        put()
    for name in sorted(set(assigned) | set(mentioned), key=lambda s: s.lower()):
        put("## {}".format(name))
        for item in assigned.get(name, []):
            put(person_line(item, name))
        if mentioned.get(name):
            put()
            put("_mentioned in_")
            for item in mentioned[name]:
                put(person_line(item, name))
        put()

    return out


DONE_HEADER = "# Completed"
DONE_NOTE = ("_Appended by TodoSync, newest first. Not part of the wiki: nothing here "
             "is indexed and nothing derives from it. It is kept because a `- [x]` "
             "line may be pruned from its note, and then this is the only record._")


def append_done(overview, entries):
    """Add completed todos to the work log, under a `## YYYY-MM` heading.

    Months descend and rows within a month descend, so the file reads newest
    first throughout and the current month is always the first thing under the
    header. Appending never rewrites a row that is already there -- this is the
    one file in the system that is not re-derived from the notes."""
    if not entries:
        return False

    path = done_path(overview)
    try:
        lines = read_lines(path) if os.path.exists(path) else []
    except Exception as e:
        log("ERROR reading work log: {}".format(e))
        return False

    nl = eol(lines)
    if not lines:
        lines = [DONE_HEADER + nl, nl, DONE_NOTE + nl, nl]

    stamp = today()
    new = []
    for entry in entries:
        line = "- [x] {}{}".format("! " if entry.get("urgent") else "", entry["text"])
        if entry.get("source"):
            line += " [[{}]]".format(entry["source"])
        new.append(line + " <!-- resolved: {} -->".format(stamp) + nl)

    heading = "## {}".format(month(stamp))
    at = None
    for i, raw in enumerate(lines):
        found = MONTH_RE.match(bare(raw))
        if not found:
            continue
        if found.group(1) == month(stamp):
            at = i + 1
            if at < len(lines) and not bare(lines[at]).strip():
                at += 1
            break
        if found.group(1) < month(stamp):
            # The first older month: this one belongs directly above it.
            lines[i:i] = [heading + nl, nl] + new + [nl]
            at = -1
            break

    if at is None:
        # No month is older, so either the log is empty of months or every one
        # of them is newer -- in both cases this month goes at the end.
        while lines and not bare(lines[-1]).strip():
            lines.pop()
        lines.extend([nl, heading + nl, nl] + new)
    elif at >= 0:
        lines[at:at] = new

    try:
        write_lines(path, lines)
    except Exception as e:
        log("ERROR writing work log: {}".format(e))
        return False
    log("work log: {} appended under {}".format(len(new), heading))
    return True


def render(window, overview, cache):
    try:
        lines = read_lines(overview) if os.path.exists(overview) else []
    except Exception as e:
        log("ERROR reading overview: {}".format(e))
        return False

    nl = eol(lines)
    capture, _old_generated, tail = split_zones(lines)
    capture = list(capture)
    tail = list(tail)

    # The blank lines around the markers are emitted here, so they have to be
    # stripped off the zones that are carried through -- otherwise every render
    # hands the next one an extra one and the file grows a line per sweep.
    while capture and not bare(capture[-1]).strip():
        capture.pop()
    while tail and not bare(tail[0]).strip():
        tail.pop(0)
    if not capture:
        capture = ["# Aggregate" + nl, nl, "## Inline capture" + nl]

    out = capture + [nl, BEGIN_MARK + nl, nl]
    persons, me = people_index(cache)
    out.extend(build_generated(collect(cache, overview), nl, persons, me))
    out.append(END_MARK + nl)
    out.append(nl)
    out.extend(tail)

    # Renders are idempotent, so identical output means nothing moved. Skipping
    # the write keeps OneDrive quiet, keeps the file out of its own version
    # history, and stops an open overview being reverted under the cursor.
    new = "".join(out)
    if STAMP_RE.sub("", new) == STAMP_RE.sub("", "".join(lines)):
        return False

    try:
        write_lines(overview, out)
    except Exception as e:
        log("ERROR writing overview: {}".format(e))
        return False
    log("overview rewritten")
    return True


def refresh_view(window, overview):
    view = window.find_open_file(overview)
    if view and not view.is_dirty():
        sublime.set_timeout(lambda: view.run_command("revert"), 0)


# ------------------------------------------------------------------
# Write-back: overview -> source notes
# ------------------------------------------------------------------

def locate(cache, source, text):
    """(path, row) for an open todo, or None. Requires the text to match, so a
    shared basename is only ambiguous when the todo text repeats too."""
    hits = []
    for record in cache["files"].values():
        path = record.get("path")
        if not path:
            continue
        if source and display_name(path) != source:
            continue
        for entry in record.get("unchecked", []):
            if entry["text"] == text:
                hits.append((path, entry["row"]))
    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        log("AMBIGUOUS: {!r} matches {} notes; not touching any".format(text, len(hits)))
    return None


def edit_source(path, row, text, check=None, urgent=None):
    """Tick a box and/or set urgency on one line. The cached row is a hint --
    the text is verified there and searched for if the note has since moved."""
    try:
        lines = read_lines(path)
    except Exception as e:
        log("ERROR opening {}: {}".format(path, e))
        return False

    def matches(i):
        if i < 0 or i >= len(lines):
            return False
        m = TODO_RE.match(bare(lines[i]))
        if not m:
            return False
        body = m.group(4)
        um = URGENT_RE.match(body)
        return (um.group(1).strip() if um else body) == text

    if not matches(row):
        row = next((i for i in range(len(lines)) if matches(i)), None)
        if row is None:
            log("NOT FOUND: {!r} in {}".format(text, path))
            return False

    m = TODO_RE.match(bare(lines[row]))
    box = m.group(3) if check is None else ("x" if check else " ")
    if urgent is None:
        urgent = bool(URGENT_RE.match(m.group(4)))

    terminator = lines[row][len(bare(lines[row])):] or "\n"
    lines[row] = "{}{} [{}] {}{}{}".format(
        m.group(1), m.group(2), box, "! " if urgent else "", text, terminator)

    try:
        write_lines(path, lines)
    except Exception as e:
        log("ERROR writing {}: {}".format(path, e))
        return False
    return True


def read_generated(overview):
    """Every checkbox line in the generated zone, as
    (checked, text, urgent, source, person).

    Every generated row now carries its source as a trailing wiki link, so the
    source is never inferred from an enclosing heading -- which is what used to
    truncate a todo whose own text ended in a link. Under '# By Person' the
    '## name' heading is still remembered, because a row there may have had its
    leading [[name]]: slot dropped and needs it back before it can be matched."""
    lines = read_lines(overview)
    _capture, generated, _tail = split_zones(lines)

    found = []
    in_by_person = False
    person = None

    for raw in generated:
        line = bare(raw)
        h = HEADING_RE.match(line)
        if h:
            title = h.group(2)
            if len(h.group(1)) == 1:
                in_by_person = title.lower().startswith("by person")
                person = None
            elif in_by_person:
                person = title
            continue

        m = LINE_RE.match(line)
        if not m:
            continue
        text = m.group(2).strip()
        urgent = False
        um = URGENT_RE.match(text)
        if um:
            urgent, text = True, um.group(1).strip()
        if not text:
            continue
        found.append((m.group(1) != " ", text, urgent, m.group(3),
                      person if in_by_person else None))
    return found


def apply_overview_edits(window):
    overview = overview_path(window)
    if not overview or not os.path.exists(overview):
        return

    cache = load_cache(overview)
    try:
        edits = read_generated(overview)
    except Exception as e:
        log("ERROR reading generated zone: {}".format(e))
        return

    resolved, retitled, missed = [], 0, 0
    unmatched = set()

    # One todo is listed several times over -- under Urgent, under Inbox and
    # under its person -- so the rows are gathered per source line before any
    # of them is acted on. Keyed on the line rather than on the text, because
    # the By Person row may have had its leading [[person]]: slot dropped and
    # is therefore not string-identical to the other two.
    order, groups = [], {}

    for checked, text, urgent, source, person in edits:
        found = locate(cache, source, text)
        if not found and person:
            # Put the dropped slot back before giving up. It has to be the
            # note's own text that goes on from here: edit_source rewrites the
            # line with whatever it is handed, so a stripped text reaching it
            # would erase the slot from the note.
            restored = "[[{}]]: {}".format(person, text)
            again = locate(cache, source, restored)
            if again:
                found, text = again, restored
        if not found:
            if checked and (text, source) not in unmatched:
                unmatched.add((text, source))
                missed += 1
            continue

        path, row = found
        k = (key(path), row)
        if k not in groups:
            order.append(k)
            groups[k] = {"path": path, "row": row, "rows": []}
        groups[k]["rows"].append((checked, text, urgent))

    for k in order:
        group = groups[k]
        path, row = group["path"], group["row"]

        # A tick wins wherever it was made. Without this the first row read
        # decides, so ticking a box in By Person -- the last section, and the
        # one actually used -- would be overruled by the untouched copy of the
        # same todo sitting further up in Inbox.
        tick = next((r for r in group["rows"] if r[0]), None)
        if tick:
            _c, text, urgent = tick
            if edit_source(path, row, text, check=True):
                resolved.append({"text": text, "urgent": urgent,
                                 "source": display_name(path)})
            else:
                missed += 1
            continue

        # Same rule for the ! toggle: the row that disagrees with the note is
        # the one the user just edited.
        text = group["rows"][0][1]
        was = next((e.get("urgent", False)
                    for e in cache["files"][key(path)]["unchecked"]
                    if e["text"] == text), False)
        flip = next((r for r in group["rows"] if r[2] != was), None)
        if flip and edit_source(path, row, flip[1], urgent=flip[2]):
            retitled += 1

    cache, completed, aborted = sweep(window, overview)
    if aborted:
        return
    append_done(overview, resolved + completed)
    if render(window, overview, cache):
        refresh_view(window, overview)

    parts = []
    if resolved:
        parts.append("{} resolved".format(len(resolved)))
    if retitled:
        parts.append("{} re-flagged".format(retitled))
    if missed:
        parts.append("{} not found -- see the log".format(missed))
    window.status_message("TodoSync: " + (", ".join(parts) if parts else "no changes"))


def sync(window, force=False, extra=None):
    overview = overview_path(window)
    if not overview:
        return
    cache, completed, aborted = sweep(window, overview, force=force, extra=extra)
    if aborted:
        window.status_message("TodoSync: symbol index looks cold; overview left alone")
        return
    append_done(overview, completed)
    if render(window, overview, cache):
        refresh_view(window, overview)


# ------------------------------------------------------------------
# Events
# ------------------------------------------------------------------

def watched(view, path):
    """Can saving this file change the overview?

    Answered from the buffer already in hand: no index lookup, no file read,
    no cache load. A note carrying a configured class counts even if Sublime
    has not indexed it yet, so a note classified for the first time appears on
    that very save. Everything else -- an email note, a meeting note, an
    unclassified note, a template -- costs one regex over its frontmatter and
    nothing else."""
    if excluded(path):
        return False
    head = view.substr(sublime.Region(0, min(view.size(), 4000)))
    cls = frontmatter_class(head)
    if cls and cls in settings().get("classes", []):
        return True
    # Still watch a note the last sweep covered, so removing its class (or
    # emptying it) is noticed rather than leaving a ghost in the overview.
    return key(path) in _scope


def qualifies_by_class(view):
    head = view.substr(sublime.Region(0, min(view.size(), 4000)))
    return frontmatter_class(head) in settings().get("classes", [])


class TodoSyncListener(sublime_plugin.EventListener):

    def on_post_save_async(self, view):
        path = view.file_name()
        window = view.window() or sublime.active_window()
        if not path or not window or not path.lower().endswith(".md"):
            return
        overview = overview_path(window)
        if not overview:
            return
        if key(path) == key(overview):
            apply_overview_edits(window)
        elif watched(view, path):
            # Hand the path over when the class says it qualifies: the index
            # will not list a note classified moments ago.
            sync(window, extra=path if qualifies_by_class(view) else None)

    def on_load_async(self, view):
        path = view.file_name()
        window = view.window() or sublime.active_window()
        if not path or not window:
            return
        overview = overview_path(window)
        if overview and key(path) == key(overview):
            sync(window)


def plugin_loaded():
    conf = settings()
    conf.clear_on_change("TodoSync")
    conf.add_on_change("TodoSync",
                       lambda: sublime.set_timeout_async(
                           lambda: sync(sublime.active_window()), 0))
    # The index may still be warming up at startup; the cached classes are
    # what keep a cold one from blanking the overview.
    sublime.set_timeout_async(lambda: sync(sublime.active_window()), 3000)


def plugin_unloaded():
    settings().clear_on_change("TodoSync")


# ------------------------------------------------------------------
# Commands
# ------------------------------------------------------------------

class TodoSyncSweepCommand(sublime_plugin.WindowCommand):
    def run(self):
        sublime.set_timeout_async(lambda: sync(self.window), 0)


class TodoSyncApplyCommand(sublime_plugin.WindowCommand):
    def run(self):
        sublime.set_timeout_async(lambda: apply_overview_edits(self.window), 0)


class TodoSyncRebuildCommand(sublime_plugin.WindowCommand):
    def run(self):
        overview = overview_path(self.window)
        if not overview:
            self.window.status_message("TodoSync: no overview found")
            return
        try:
            os.remove(cache_path(overview))
        except OSError:
            pass
        sublime.set_timeout_async(lambda: sync(self.window, force=True), 0)


class TodoSyncShowScopeCommand(sublime_plugin.WindowCommand):
    """What is currently being swept, and why. The diagnostic to reach for
    while classifying notes."""

    def run(self):
        overview = overview_path(self.window)
        if not overview:
            self.window.status_message("TodoSync: no overview found")
            return
        cache = load_cache(overview)
        self.paths = sorted(in_scope(self.window, overview).values(),
                            key=lambda p: display_name(p).lower())
        if not self.paths:
            self.window.status_message(
                "TodoSync: nothing in scope -- check 'classes' and let the index settle")
            return

        items = []
        for path in self.paths:
            record = cache["files"].get(key(path), {})
            items.append([
                display_name(path),
                "{} open  |  class: {}".format(
                    len(record.get("unchecked", [])), record.get("class") or "--"),
            ])
        self.window.show_quick_panel(items, self.open_at)

    def open_at(self, index):
        if index >= 0:
            self.window.open_file(self.paths[index])
