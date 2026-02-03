import sublime
import sublime_plugin
import os
import re
import subprocess
import webbrowser

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
