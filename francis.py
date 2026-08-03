import sublime
import sublime_plugin
import html
import json
import os
import re
import subprocess
import threading
import webbrowser
import tempfile
from urllib.parse import quote

# ---------------------------
# Helper function to open file or URL
# ---------------------------
def open_file(path):
    # Detect URLs
    url_regex = re.compile(
        r'^(?:http|ftp)s?://'  # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'  # domain
        r'localhost|'  # localhost
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}|'  # ipv4
        r'\[?[A-F0-9]*:[A-F0-9:]+\]?)'  # ipv6
        r'(?::\d+)?'  # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)

    if url_regex.match(path):
        webbrowser.open(path)
    else:
        # assume relative path or absolute file path
        if not os.path.isabs(path):
            # fallback to current file directory
            window = sublime.active_window()
            if window:
                view = window.active_view()
                if view and view.file_name():
                    md_dir = os.path.dirname(view.file_name())
                    path = os.path.normpath(os.path.join(md_dir, path))

        if os.path.exists(path):
            try:
                if sublime.platform() == "windows":
                    os.startfile(path)
                elif sublime.platform() == "osx":
                    subprocess.call(['open', path])
                else:
                    subprocess.call(['xdg-open', path])
            except Exception as e:
                print("Failed to open file {0}. Error: {1}".format(path, e))
        else:
            print("File does not exist: {0}".format(path))

# ---------------------------
# Helpers: YAML frontmatter
# ---------------------------

# The whole frontmatter block, delimiters included, carries this scope
# (see MarkdownExtended.sublime-syntax).
FRONTMATTER_SELECTOR = "meta.frontmatter.markdown"

# The '---' delimiters themselves. MarkdownEditing's syntaxes/Fold.tmPreferences
# defines a foldScope between these two, so they also mark out the exact region
# Sublime's own frontmatter fold uses.
FRONTMATTER_BEGIN_SELECTOR = "punctuation.section.frontmatter.begin.markdown"
FRONTMATTER_END_SELECTOR = "punctuation.section.frontmatter.end.markdown"

# Frontmatter lives at the top of the file; no need to read the whole email.
# Generous because parse_frontmatter needs the CLOSING '---' inside the window
# or it matches nothing, and a note carrying an Outlook store_id (several
# hundred hex characters) plus a long attachment list eats into the margin.
FRONTMATTER_READ_CHARS = 16384

def parse_frontmatter(content):
    """Return the raw YAML frontmatter block from Markdown content, or None."""
    m = re.match(r'^---\s*(?:yaml|yml)?\s*\n(.*?)\n---', content, re.DOTALL | re.IGNORECASE)
    return m.group(1) if m else None

def frontmatter_field(frontmatter, key):
    """Return a scalar frontmatter value with quotes stripped, or '' if absent."""
    if not frontmatter:
        return ""
    m = re.search(
        r'^{0}\s*:\s*"?([^"\r\n]+?)"?\s*$'.format(re.escape(key)),
        frontmatter,
        re.MULTILINE
    )
    return m.group(1).strip() if m else ""

def read_frontmatter(path):
    """Read just enough of a Markdown file on disk to parse its frontmatter."""
    try:
        # utf-8-sig, not utf-8: the Outlook macro's SaveAsUTF8 writes a BOM, and
        # a leading ﻿ stops parse_frontmatter's '^---' anchor from matching.
        with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
            return parse_frontmatter(f.read(FRONTMATTER_READ_CHARS))
    except Exception as e:
        print("Could not read frontmatter from {0}: {1}".format(path, e))
        return None

# The Outlook back-references, in the order the macro writes them.
OUTLOOK_REF_KEYS = ("entry_id", "store_id", "message_id", "conversation_id")

def outlook_refs(frontmatter):
    """Read the foldable 'outlook:' sequence into a dict.

    The macro writes the back-references as a YAML sequence rather than as
    top-level keys, so the whole block folds to one line like 'attachments:'
    does. That also puts them out of frontmatter_field's reach: it anchors keys
    at column 0, and these sit indented under a '- '.

    Empty values are dropped so a caller can just test for the key -- the macro
    writes 'entry_id: ' empty rather than omitting the line when the lookup
    failed on its side.
    """
    refs = {}
    if not frontmatter:
        return refs
    in_block = False
    for line in frontmatter.splitlines():
        if re.match(r'^outlook\s*:\s*$', line):
            in_block = True
            continue
        if in_block:
            m = re.match(r'^\s+-\s*(\w+)\s*:\s*"?([^"\r\n]*?)"?\s*$', line)
            if not m:
                break              # first non-item line ends the sequence
            value = m.group(2).strip()
            if value:
                refs[m.group(1)] = value
    return refs

# ---------------------------
# Helpers: the Outlook bridge
# ---------------------------

# outlook_bridge.ps1 sits at the package root but is inert to Sublime -- a .ps1
# is loaded by nothing -- and is reached only through subprocess. The package is
# a loose directory, so the script is on disk and directly executable; were it
# ever zipped into a .sublime-package this would have to become
# sublime.load_resource() plus a write to temp.
BRIDGE_SCRIPT = "outlook_bridge.ps1"
PACKAGE_NAME = "francis_sublime_package"

# powershell.exe, not pwsh: PowerShell 7's COM interop with Outlook is
# unreliable, and 5.1 ships with Windows.
POWERSHELL = "powershell.exe"

CREATE_NO_WINDOW = 0x08000000

# Attaching to a cold Outlook takes seconds, and the first COM call after a
# resume can take considerably longer -- but it still has to be bounded, or a
# wedged call would leak a thread per invocation.
BRIDGE_TIMEOUT_SECONDS = 90

BRIDGE_ACTIONS = ("resolve", "open", "reply", "reply_all", "forward")

# What the user is told for each 'reason' the bridge can report. The bridge's
# own 'message' is COM's wording and is appended underneath.
BRIDGE_REASONS = {
    "bad_request": "The bridge rejected the request.",
    "no_outlook": "Could not start Outlook Classic. The Outlook.Application "
                  "COM server is registered by the classic Win32 client only.",
    "not_found": "Outlook has no item with that entry_id.\n\n"
                 "An EntryID is a lease, not an address: Outlook re-keys it "
                 "when the item moves between stores, so an archived or "
                 "deleted message no longer resolves. The message_id and "
                 "time-window fallbacks that would find it are not built yet.",
    "unsupported_item": "That action does not apply to this kind of item.",
    "com_error": "Outlook refused the call.",
}

def bridge_path():
    """Absolute path to the PowerShell bridge."""
    return os.path.join(sublime.packages_path(), PACKAGE_NAME, BRIDGE_SCRIPT)

def parse_last_json_line(data):
    """The bridge's response is the last JSON object it printed on stdout."""
    text = (data or b"").decode("utf-8", "replace")
    for line in reversed(text.splitlines()):
        line = line.strip()
        if line.startswith("{"):
            try:
                return json.loads(line)
            except ValueError:
                continue
    return None

def run_bridge(request, on_done):
    """Run outlook_bridge.ps1 off the UI thread; call on_done(result, error) on it.

    The request goes through a JSON temp file rather than the command line:
    subjects and display names carry quotes, '$' and backticks, every one of
    which is a quoting hazard through cmd -> PowerShell.

    Exactly one of result/error is set. `error` is for the bridge never
    answering (no PowerShell, no script, a timeout); a bridge that answers
    reports its own failures inside `result` with an 'ok': false.
    """
    def work():
        result, error, req_path = None, None, None
        try:
            fd, req_path = tempfile.mkstemp(suffix=".json", prefix="francis_outlook_")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(request, f)

            startupinfo, creationflags = None, 0
            if sublime.platform() == "windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = CREATE_NO_WINDOW

            proc = subprocess.Popen(
                [POWERSHELL, "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-File", bridge_path(),
                 "-RequestPath", req_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                startupinfo=startupinfo,
                creationflags=creationflags
            )
            try:
                out, err = proc.communicate(timeout=BRIDGE_TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                proc.kill()
                out, err = proc.communicate()
                out = b""
                error = "Outlook did not respond within {0} seconds.".format(
                    BRIDGE_TIMEOUT_SECONDS
                )

            if error is None:
                result = parse_last_json_line(out)
                if result is None:
                    # The bridge routes every exit through Emit, so an
                    # unparseable stdout means it never ran at all.
                    error = (err or out or b"").decode("utf-8", "replace").strip() \
                        or "The bridge produced no response."
        except Exception as e:
            error = str(e)
        finally:
            if req_path:
                try:
                    os.unlink(req_path)
                except OSError:
                    pass
        sublime.set_timeout(lambda: on_done(result, error), 0)

    threading.Thread(target=work, daemon=True).start()

def clean_name(value):
    """Unwrap a '[[@Person Name]]' wiki link written by the Outlook macro."""
    value = value.strip()
    if value.startswith("[[") and value.endswith("]]"):
        value = value[2:-2]
    if value.startswith("@"):
        value = value[1:]
    return value.strip()

def collapse_whitespace(value):
    """Squeeze runs of whitespace; the macro does not compress subject lines."""
    return re.sub(r'\s+', ' ', value).strip()

# ---------------------------
# Helpers: email discovery
# ---------------------------

# Indexed symbol names produced by Symbol List - Email Flags.tmPreferences.
ALL_EMAILS_SYMBOL = "email_note"
TOREPLY_SYMBOL = "toreply_note"
PRIORITY_SYMBOL = "priority_note"
PERSONAL_SYMBOL = "personal_note"

# Filter categories offered in the browser, as (key, symbol, label). Also the
# flags carried per row, so each is both something to subset by and something
# to badge with.
CATEGORIES = [
    ("toreply", TOREPLY_SYMBOL, "to-reply"),
    ("priority", PRIORITY_SYMBOL, "high-priority"),
    ("personal", PERSONAL_SYMBOL, "personal"),
]

def category_symbol(key):
    """The indexed symbol backing a filter key."""
    for k, symbol, _ in CATEGORIES:
        if k == key:
            return symbol
    return None

def categories_for_symbol(symbol):
    """The filter keys a symbol already implies, as a tuple."""
    return tuple(k for k, s, _ in CATEGORIES if s == symbol)

# Filenames written by the Outlook macro pack every field we display:
#   <prefix><yyyy-MM-dd>_<HHNN> @<sender> - <subject, truncated to 60>.md
# Meetings omit the ' @<sender>' part. Anchoring on the date token means the
# configurable prefix ('Email ', 'Sent_', 'M ') does not matter. The sender is
# lazy and the subject greedy, so the FIRST ' - ' separates them.
EMAIL_FILENAME_RE = re.compile(
    r'^.*?(?P<date>\d{4}-\d{2}-\d{2})_(?P<hh>\d{2})(?P<mm>\d{2})'
    r'(?: @(?P<sender>.*?))?'
    r' - (?P<subject>.*)$'
)

def parse_email_filename(path):
    """Pull date, sender and subject out of a macro-written filename.

    Returns None if the name does not follow the macro's pattern, so the caller
    can fall back to opening the file.
    """
    m = EMAIL_FILENAME_RE.match(os.path.splitext(os.path.basename(path))[0])
    if not m:
        return None
    return {
        "subject": collapse_whitespace(m.group("subject")),
        "sender": clean_name(m.group("sender") or ""),
        "date": "{0} {1}:{2}".format(m.group("date"), m.group("hh"), m.group("mm")),
    }

def read_entry_fields(path):
    """Fallback for notes whose filename does not follow the convention."""
    fm = read_frontmatter(path)
    sender = frontmatter_field(fm, "from") or frontmatter_field(fm, "sender")
    return {
        "subject": collapse_whitespace(frontmatter_field(fm, "subject")),
        "sender": clean_name(sender),
        "date": frontmatter_field(fm, "date"),
    }

def symbol_paths(window, symbol):
    """Set of indexed file paths carrying `symbol`. Empty if it is not indexed."""
    return set(
        loc.path for loc in window.symbol_locations(symbol, sublime.SYMBOL_SOURCE_INDEX)
    )

def email_entries(window, symbols):
    """Entries for files carrying every symbol in `symbols`.

    Nothing is opened: paths come from the symbol index and the displayed
    fields are parsed out of the filename, so cost is flat in vault size. The
    only file read is the fallback for a name that does not match the macro's
    convention.

    Ordered 'priority: high' first, then undated notes, then newest date first.
    """
    if not symbols:
        return []

    paths = symbol_paths(window, symbols[0])
    for extra in symbols[1:]:
        paths &= symbol_paths(window, extra)

    # One set query per flag, then O(1) membership per row.
    flag_sets = [(key, symbol_paths(window, symbol)) for key, symbol, _ in CATEGORIES]

    entries = []
    for path in sorted(paths):
        fields = parse_email_filename(path) or read_entry_fields(path)
        fields["path"] = path
        fields["sender"] = fields["sender"] or "unknown sender"
        fields["subject"] = fields["subject"] or os.path.basename(path)
        fields["flags"] = frozenset(k for k, members in flag_sets if path in members)
        entries.append(fields)

    # Sorted in passes, relying on sort() being stable: each pass keeps the
    # ordering established by the one before it.
    entries.sort(key=lambda e: e["date"], reverse=True)      # newest date first
    entries.sort(key=lambda e: (
        0 if "priority" in e["flags"] else 1,                # high priority on top
        0 if not e["date"] else 1,                           # undated above dated
    ))
    return entries

# ---------------------------
# Helpers: row rendering
# ---------------------------

KIND_FILTER = (sublime.KIND_ID_NAVIGATION, "/", "Filter")
KIND_MODE = (sublime.KIND_ID_COLOR_CYANISH, "⇄", "Search field")
KIND_OPEN_ALL = (sublime.KIND_ID_NAVIGATION, "*", "Open all")
KIND_PRIORITY = (sublime.KIND_ID_COLOR_REDISH, "!", "High priority")
KIND_TOREPLY = (sublime.KIND_ID_COLOR_ORANGISH, "↩", "To reply")
KIND_PERSONAL = (sublime.KIND_ID_COLOR_GREENISH, "⌂", "Personal")
KIND_PLAIN = sublime.KIND_AMBIGUOUS

# A row carries one kind, so the first flag it has wins.
BADGE_ORDER = [
    ("priority", KIND_PRIORITY),
    ("toreply", KIND_TOREPLY),
    ("personal", KIND_PERSONAL),
]

def kind_for(entry, applied=()):
    """Badge for a row, skipping any category the list is already filtered to.

    A badge every row carries says nothing and only costs width, so narrowing
    to /priority drops the priority badge and lets the next flag down show
    through instead.
    """
    for key, kind in BADGE_ORDER:
        if key in applied:
            continue
        if key in entry["flags"]:
            return kind
    return KIND_PLAIN

def row_text(entry, by="sender"):
    """Trigger and detail for a row, as (trigger, detail).

    Only the trigger is matched against typed input (see ListInputItem.text),
    so whichever field is searchable goes on top and the other drops to the
    detail line. The detail renders as minihtml and must be escaped.
    """
    if by == "subject":
        return entry["subject"], html.escape(entry["sender"])
    return entry["sender"], html.escape(entry["subject"])

def is_action_value(value):
    """True for the '/...' rows, which act rather than resolving to a file."""
    return isinstance(value, str) and value.startswith(("cat:", "by:"))

# ---------------------------
# Helpers: preview folding
# ---------------------------

FOLD_POLL_MS = 25         # transient views load async; poll until ready
FOLD_MAX_TRIES = 20

# Views that already had a tab when we folded them, as
# view_id -> (view, region, selection), so both can be put back afterwards.
# Only one browser or panel is ever open at a time, so a module-level record is
# enough and saves threading it through the input handler chain.
_PREVIEW_FOLDS = {}

def fold_frontmatter(window, view, tries=0):
    """Collapse the frontmatter so a preview opens on the email body."""
    if not view.is_valid():
        return   # arrowed past it; the transient view is already gone

    if view.is_loading():
        if tries < FOLD_MAX_TRIES:
            sublime.set_timeout(
                lambda: fold_frontmatter(window, view, tries + 1), FOLD_POLL_MS
            )
        return

    begins = view.find_by_selector(FRONTMATTER_BEGIN_SELECTOR)
    ends = view.find_by_selector(FRONTMATTER_END_SELECTOR)
    if not begins or not ends:
        return   # no frontmatter, or not the MarkdownExtended syntax

    # The exact region Sublime's own foldScope produces: from the end of the
    # opening '---' line to just before the closing '---', with the trailing
    # newline excluded.
    fold_start = view.line(begins[0].begin()).end()
    fold_end = ends[0].begin() - 1
    if fold_start >= fold_end:
        return

    probe = fold_start + 1   # a point inside the fold, to test its state
    if any(r.contains(probe) for r in view.folded_regions()):
        view.set_viewport_position((0, 0), False)
        return

    original_sel = [r for r in view.sel()]

    # The caret must sit at the start of the foldable region -- right after the
    # opening '---', still on that line -- for the built-in fold command to act.
    view.sel().clear()
    view.sel().add(sublime.Region(fold_start))
    view.run_command("fold")

    if not any(r.contains(probe) for r in view.folded_regions()):
        view.fold(sublime.Region(fold_start, fold_end))

    transient = window.transient_view_in_group(window.active_group())
    transient_id = transient.id() if transient else None
    if view.id() != transient_id and view.id() not in _PREVIEW_FOLDS:
        _PREVIEW_FOLDS[view.id()] = (
            view, sublime.Region(fold_start, fold_end), original_sel
        )

    view.set_viewport_position((0, 0), False)

def restore_preview_folds(keep_view=None):
    """Unfold pre-existing tabs we touched, except the one just selected."""
    keep_id = keep_view.id() if keep_view else None
    for view_id, (view, region, original_sel) in _PREVIEW_FOLDS.items():
        if view_id == keep_id or not view.is_valid():
            continue
        view.unfold(region)
        view.sel().clear()
        view.sel().add_all(original_sel)
    _PREVIEW_FOLDS.clear()

# ---------------------------
# Command 1: Open Markdown link
# ---------------------------
class OpenLinkedFileCommand(sublime_plugin.TextCommand):
    """
    Open Markdown links [text](url_or_file) under caret.
    Resolves relative paths relative to the Markdown file automatically.
    """
    def run(self, edit):
        for region in self.view.sel():
            pt = region.begin()

            # Search backward for opening '['
            text_before = self.view.substr(sublime.Region(0, pt))
            start = text_before.rfind("[")
            if start == -1:
                print("No Markdown link under caret")
                continue

            # Search forward for closing ')'
            text_after = self.view.substr(sublime.Region(pt, self.view.size()))
            close_idx = text_after.find(")")
            if close_idx == -1:
                print("No Markdown link under caret")
                continue
            end = pt + close_idx

            # Extract the full link region
            link_region = sublime.Region(start, end + 1)
            full_link = self.view.substr(link_region)

            # Extract the path inside parentheses
            m = re.search(r"\((.*?)\)", full_link)
            if not m:
                print("No path found in link")
                continue

            path = m.group(1).strip("(<\">)")

            print("Resolved link: {0}".format(path))
            open_file(path)

# ---------------------------
# Command 2: Open frontmatter attachment
# ---------------------------
class OpenFrontmatterAttachmentCommand(sublime_plugin.TextCommand):
    """
    Open attachment files from frontmatter.
    Assumes caret is inside the quotation marks of an attachment name.
    Prepends './attachments/' and resolves relative to the Markdown file.
    """
    ATTACHMENTS_FOLDER = "./attachments"

    def run(self, edit):
        for region in self.view.sel():
            pt = region.begin()
            scope_name = self.view.scope_name(pt)

            # Check that caret is in the attachment filename scope
            if "entity.name.attachment.yaml" not in scope_name:
                print("Caret not in a frontmatter attachment string")
                continue

            # Expand region to include the quoted string
            str_region = self.view.extract_scope(pt)
            attachment_name = self.view.substr(str_region).strip("\"")

            # Build full path relative to the Markdown file
            md_file = self.view.file_name()
            if md_file:
                md_dir = os.path.dirname(md_file)
                attachment_path = os.path.normpath(
                    os.path.join(md_dir, self.ATTACHMENTS_FOLDER, attachment_name)
                )
                print("Resolved frontmatter attachment: {0}".format(attachment_path))
                open_file(attachment_path)
            else:
                print("Current file not saved; cannot resolve relative path")

# ---------------------------
# Command 3: Convert Office files to Markdown
# ---------------------------
class ConvertToMarkdownCommand(sublime_plugin.WindowCommand):
    """
    Convert Word/PowerPoint/PDF files to Markdown using Pandoc and marker-pdf.
    Opens the result in a new buffer in Sublime Text.
    """
    PANDOC_PATH = r"C:\portablePrograms\pandoc-3.6.4\pandoc.exe"
    MARKER_SINGLE_PATH = r"C:\portablePrograms\python\Scripts\marker_single.exe"

    def run(self, file_path=None):
        if file_path:
            self.convert_file(file_path)
        else:
            # Check clipboard for a file path first
            initial_path = ""
            clipboard_text = sublime.get_clipboard().strip()

            # Check if clipboard contains a valid file path
            if clipboard_text:
                # Strip quotes if present
                clipboard_text = clipboard_text.strip('"').strip("'")
                # Normalize path
                clipboard_path = os.path.normpath(clipboard_text)
                # Use clipboard path if it's a valid file
                if os.path.isfile(clipboard_path):
                    initial_path = clipboard_path

            self.window.show_input_panel(
                "File to convert to Markdown (Word/PowerPoint/PDF):",
                initial_path,
                self.convert_file,
                None,
                None
            )

    def convert_file(self, file_path):
        if not file_path:
            sublime.error_message("No file path provided")
            return

        # Strip quotes and whitespace
        file_path = file_path.strip().strip('"').strip("'")

        # Normalize path
        file_path = os.path.normpath(file_path)

        print("ConvertToMarkdown: Checking file path: {0}".format(file_path))

        if not os.path.exists(file_path):
            sublime.error_message("File does not exist: {0}".format(file_path))
            print("ConvertToMarkdown: File not found")
            return

        print("ConvertToMarkdown: File found, proceeding with conversion")

        # Check file extension
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        if ext not in ['.docx', '.pptx', '.doc', '.ppt', '.pdf']:
            sublime.error_message("Unsupported file type. Please use .docx, .pptx, or .pdf files.")
            return

        try:
            file_dir = os.path.dirname(file_path)
            basename = os.path.basename(file_path)
            name_without_ext = os.path.splitext(basename)[0]

            # On Windows, prevent console window from appearing
            startupinfo = None
            if sublime.platform() == "windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            # Use different converter based on file type
            if ext == '.pdf':
                # Use marker-pdf for PDF files
                output_dir = tempfile.mkdtemp()
                cmd = [
                    self.MARKER_SINGLE_PATH,
                    file_path,
                    '--output_dir', output_dir,
                    '--output_format', 'markdown'
                ]

                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    startupinfo=startupinfo,
                    universal_newlines=True
                )

                stdout, stderr = process.communicate()

                if process.returncode != 0:
                    error_msg = "marker-pdf conversion failed (exit code {0}):\n{1}".format(
                        process.returncode,
                        stderr if stderr else "No error details available"
                    )
                    sublime.error_message(error_msg)
                    print("marker-pdf stderr: {0}".format(stderr))
                    return

                # Read the generated markdown file
                output_md = os.path.join(output_dir, name_without_ext, "{0}.md".format(name_without_ext))
                if not os.path.exists(output_md):
                    # Try alternative path structure
                    output_md = os.path.join(output_dir, "{0}.md".format(name_without_ext))

                if not os.path.exists(output_md):
                    sublime.error_message("marker-pdf completed but output file not found:\n{0}".format(output_md))
                    return

                with open(output_md, 'r', encoding='utf-8') as f:
                    markdown_content = f.read()

            else:
                # Use Pandoc for Word/PowerPoint files
                media_dir = os.path.join(file_dir, 'media')
                cmd = [
                    self.PANDOC_PATH,
                    file_path,
                    '-t', 'markdown',
                    '--wrap=none',
                    '--extract-media={0}'.format(media_dir)
                ]

                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    cwd=file_dir,
                    startupinfo=startupinfo,
                    universal_newlines=True
                )

                stdout, stderr = process.communicate()

                if process.returncode != 0:
                    error_msg = "Pandoc conversion failed (exit code {0}):\n{1}".format(
                        process.returncode,
                        stderr if stderr else "No error details available"
                    )
                    sublime.error_message(error_msg)
                    print("Pandoc stderr: {0}".format(stderr))
                    return

                markdown_content = stdout

            # Create a new buffer and insert the markdown
            new_view = self.window.new_file()
            new_view.set_syntax_file('Packages/Markdown/Markdown.sublime-syntax')
            new_view.run_command('append', {'characters': markdown_content})

            # Set the name to indicate source
            new_view.set_name("{0}.md".format(name_without_ext))

            # Set scratch to false so it can be saved
            new_view.set_scratch(False)

            self.window.status_message("Converted {0} to Markdown".format(basename))
        except Exception as e:
            sublime.error_message("Error during conversion: {0}".format(str(e)))


# ---------------------------
# Command 4: Reply to email
# ---------------------------
class ReplyToEmailCommand(sublime_plugin.TextCommand):
    """
    Parse YAML frontmatter for from_email and subject, then open
    a mailto: link to compose a reply.
    """
    def run(self, edit):
        content = self.view.substr(sublime.Region(0, self.view.size()))

        # Extract YAML frontmatter between --- delimiters
        frontmatter = parse_frontmatter(content)
        if frontmatter is None:
            sublime.error_message("No YAML frontmatter found in this file.")
            return

        email = frontmatter_field(frontmatter, "from_email")
        if not email:
            sublime.error_message("No from_email field found in frontmatter.")
            return

        subject = frontmatter_field(frontmatter, "subject")

        # Prepend Re: if not already present
        if subject and not re.match(r'^Re:', subject, re.IGNORECASE):
            subject = "Re: " + subject

        # Build mailto URL
        mailto = "mailto:{0}".format(quote(email, safe='@'))
        if subject:
            mailto += "?subject={0}".format(quote(subject, safe=''))

        webbrowser.open(mailto)


# ---------------------------
# Command 5: Open files by frontmatter symbol
# ---------------------------
class OpenFrontmatterSymbolCommand(sublime_plugin.WindowCommand):
    """
    Find every indexed file carrying a given frontmatter symbol and either
    browse them in a quick panel (pick=True) or open them all in tabs.

    Symbol-agnostic: pass any indexed symbol name via args, e.g.
        {"symbol": "toreply_note",  "label": "to-reply"}
        {"symbol": "personal_note", "label": "personal"}
    """
    OPEN_ALL_THRESHOLD = 20   # above this, confirm before spraying tabs

    def run(self, symbol=ALL_EMAILS_SYMBOL, label=None, pick=True):
        self.label = label or symbol
        # Every row already carries whatever this symbol implies, so that
        # badge is suppressed and the next flag down shows instead.
        self.applied = categories_for_symbol(symbol)
        self.entries = email_entries(self.window, [symbol])
        _PREVIEW_FOLDS.clear()

        if not self.entries:
            self.window.status_message(
                "No files marked {0} (is the folder in this project and indexed?)".format(self.label)
            )
            return

        if pick:
            self.show_panel()
        else:
            self.open_all()

    # --- Open all ---------------------------------------------------------

    def open_all(self):
        count = len(self.entries)
        if count > self.OPEN_ALL_THRESHOLD:
            confirmed = sublime.ok_cancel_dialog(
                "{0} files are marked {1}. Open all of them in tabs?".format(count, self.label),
                "Open All"
            )
            if not confirmed:
                return

        for entry in self.entries:
            self.window.open_file(entry["path"])

        self.window.status_message("Opened {0} file(s) marked {1}".format(count, self.label))

    # --- Browse -----------------------------------------------------------

    def show_panel(self):
        items = [sublime.QuickPanelItem(
            "Open all ({0})".format(len(self.entries)),
            details="Open every file marked {0} in tabs".format(self.label),
            kind=KIND_OPEN_ALL
        )]

        for entry in self.entries:
            # Sender on the trigger line with the date right-aligned beside it;
            # the subject sits underneath, free to run the full panel width.
            trigger, detail = row_text(entry)
            items.append(sublime.QuickPanelItem(
                trigger,
                details=detail,
                annotation=entry["date"],
                kind=kind_for(entry, self.applied)
            ))

        self.window.show_quick_panel(items, self.on_done, on_highlight=self.on_highlight)

    def on_highlight(self, index):
        if index <= 0:   # row 0 is 'Open all', nothing to preview
            return
        fold_frontmatter(
            self.window,
            self.window.open_file(self.entries[index - 1]["path"], sublime.TRANSIENT)
        )

    def on_done(self, index):
        if index == 0:
            restore_preview_folds()
            self.open_all()
            return

        selected = None
        if index > 0:
            selected = self.window.open_file(self.entries[index - 1]["path"])
        restore_preview_folds(selected)


# ---------------------------
# Command 6: Browse emails, narrowing by category
# ---------------------------
class EmailListInput(sublime_plugin.ListInputHandler):
    """
    One step of the email browser.

    The '/category' rows sit in the same list as the emails, so typing '/'
    simply fuzzy-filters down to them. That is deliberate: the input offers no
    callback for typing, so there is no way to intercept a keystroke and switch
    modes -- but since almost no sender contains a slash, the effect is the
    same. Choosing a category returns a fresh handler with that filter applied,
    which is what makes Sublime draw the chip and re-populate the list.

    list_items() is called once per handler activation (see setup_ in
    sublime_plugin.py), and filtering then runs natively, so typing costs
    nothing in Python no matter how many emails are listed.
    """
    def __init__(self, window, applied=(), by="sender", depth=0):
        self.window = window
        self.applied = tuple(applied)
        self.by = by            # which field is the searchable trigger
        self.depth = depth      # counts filters AND field flips

    def name(self):
        # One command argument per step; the depth keeps the names distinct.
        # It is tracked separately from `applied` because a field flip is also
        # a step but does not add a filter.
        return "step{0}".format(self.depth)

    def placeholder(self):
        return "Search {0}  ( / to narrow )".format(self.by)

    def description(self, value, text):
        # Text Sublime puts in the chip once this step is no longer on top.
        return text.lstrip("/")

    def entries(self):
        symbols = [ALL_EMAILS_SYMBOL] + [category_symbol(k) for k in self.applied]
        return email_entries(self.window, symbols)

    def list_items(self):
        items = [
            sublime.ListInputItem(
                "/" + key,
                "cat:" + key,
                details="Show only {0} emails".format(label),
                kind=KIND_FILTER,
            )
            for key, _, label in CATEGORIES if key not in self.applied
        ]

        # Flipping which field is searchable needs a rebuilt list, since the
        # trigger is fixed when list_items() runs -- so it is offered as a row
        # like the filters rather than as a keystroke. The typed text cannot
        # survive a rebuild anyway (it is not readable), so choosing the field
        # up front is the only workflow that does not lose your query.
        flip = "subject" if self.by == "sender" else "sender"
        items.append(sublime.ListInputItem(
            "/" + flip,
            "by:" + flip,
            details="Search {0}s instead".format(flip),
            kind=KIND_MODE,
        ))

        for entry in self.entries():
            trigger, detail = row_text(entry, self.by)
            items.append(sublime.ListInputItem(
                trigger,
                entry["path"],
                details=detail,
                annotation=entry["date"],
                kind=kind_for(entry, self.applied),
            ))
        return items

    def next_input(self, args):
        value = args.get(self.name(), "")
        if isinstance(value, str):
            if value.startswith("cat:"):
                return EmailListInput(self.window, self.applied + (value[4:],),
                                      self.by, self.depth + 1)
            if value.startswith("by:"):
                return EmailListInput(self.window, self.applied,
                                      value[3:], self.depth + 1)
        return None   # an email was chosen; the command runs

    def preview(self, value):
        # The on_highlight equivalent: open the file transiently and fold its
        # frontmatter so the preview lands on the body.
        if value and not is_action_value(value):
            fold_frontmatter(self.window, self.window.open_file(value, sublime.TRANSIENT))
        return None

    def cancel(self):
        restore_preview_folds()


class BrowseEmailsCommand(sublime_plugin.WindowCommand):
    """
    Browse every indexed email, narrowing by '/toreply', '/priority' or
    '/personal'. Each chosen filter becomes a chip at the left of the input and
    subsets the list; backspace on an empty input removes the last one.
    """
    def input(self, args):
        # Rebuild the handler's state from the steps chosen so far.
        applied, by, depth = [], "sender", 0
        while True:
            value = args.get("step{0}".format(depth))
            if not isinstance(value, str):
                break
            if value.startswith("cat:"):
                applied.append(value[4:])
            elif value.startswith("by:"):
                by = value[3:]
            else:
                return None      # an email is already chosen; nothing to ask
            depth += 1

        if depth == 0:
            _PREVIEW_FOLDS.clear()   # fresh browse, not a further narrowing
        return EmailListInput(self.window, tuple(applied), by, depth)

    def input_description(self):
        return "Emails"

    def run(self, **steps):
        path = None
        for key in sorted((k for k in steps if k.startswith("step")),
                          key=lambda k: int(k[4:])):
            value = steps[key]
            if isinstance(value, str) and not is_action_value(value):
                path = value

        if not path:
            return
        restore_preview_folds(self.window.open_file(path))


# ---------------------------
# Command 7: Act on the Outlook item a note came from
# ---------------------------

# How a completed action is reported in the status bar.
ACTION_DONE = {
    "resolve": "resolved",
    "open": "opened",
    "reply": "reply window open for",
    "reply_all": "reply-all window open for",
    "forward": "forward window open for",
}

class OutlookMailActionCommand(sublime_plugin.TextCommand):
    """
    Open, reply to, reply-all to or forward the live Outlook item this note was
    exported from.

    args: {"action": "resolve" | "open" | "reply" | "reply_all" | "forward"}

    Parameterised by args rather than split into five near-identical classes,
    the same shape as open_frontmatter_symbol.

    Resolution is entry_id + store_id and nothing else -- one O(1)
    GetItemFromID. store_id is not decoration: without it Outlook searches only
    the default store, so anything in an archive or a shared mailbox resolves to
    nothing. The macro writes both.

    A note with no entry_id is not an error in the code -- notes exported before
    the macro started writing the 'outlook:' block have none, and sent notes
    never get one, because .Send re-keys the item. Nor is a note whose message
    has since been archived: Outlook re-keys the EntryID on a move between
    stores. Both report and stop; the message_id and time-window search tiers
    that would cover them are deliberately not built yet.

    'resolve' touches nothing and is the safe probe: it exercises the whole path
    and composes no window. Its full response goes to the console.
    """
    def run(self, edit, action="open"):
        if action not in BRIDGE_ACTIONS:
            sublime.error_message("Unknown Outlook action: {0}".format(action))
            return

        content = self.view.substr(sublime.Region(0, self.view.size()))
        frontmatter = parse_frontmatter(content)
        if frontmatter is None:
            sublime.error_message("No YAML frontmatter found in this file.")
            return

        refs = outlook_refs(frontmatter)
        if not refs.get("entry_id"):
            sublime.error_message(
                "This note carries no Outlook entry_id.\n\n"
                "Notes exported before the macro started writing the 'outlook:' "
                "block have none, and sent notes never get one. Use "
                "'Reply to email' for those -- it composes via mailto, which "
                "does not thread but needs nothing from Outlook."
            )
            return

        sublime.status_message("Outlook: {0}...".format(action.replace("_", " ")))
        run_bridge(
            {
                "action": action,
                "entry_id": refs["entry_id"],
                "store_id": refs.get("store_id", ""),
            },
            lambda result, error: self.on_done(action, result, error)
        )

    def on_done(self, action, result, error):
        if error:
            sublime.error_message("Could not reach Outlook.\n\n{0}".format(error))
            return

        print("outlook_mail_action ({0}): {1}".format(action, result))

        if not result.get("ok"):
            reason = result.get("reason", "com_error")
            sublime.error_message("{0}\n\n{1}".format(
                BRIDGE_REASONS.get(reason, reason),
                result.get("message", "")
            ).strip())
            return

        sublime.status_message("Outlook: {0} {1}".format(
            ACTION_DONE.get(action, action),
            collapse_whitespace(result.get("subject") or "(no subject)")
        ))


