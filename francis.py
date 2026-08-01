import sublime
import sublime_plugin
import os
import re
import subprocess
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
