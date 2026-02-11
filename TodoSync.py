import sublime
import sublime_plugin
import os
import json
import datetime
import re

MARKDOWN_EXT = ".md"
TODO_YAML = "todo.yml"
TODO_MD = "todoUNDP.md"
TODO_PARENT_DIR = "UNDPnotes"

EXCLUDE_FOLDERS = [
    "emails",
]

# ------------------------------
# Helpers: Logging
# ------------------------------

def get_log_path():
    """Get the path to the log file in the Sublime Text data directory."""
    packages_path = sublime.packages_path()
    parent_dir = os.path.dirname(packages_path)
    log_dir = os.path.join(parent_dir, "Logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    return os.path.join(log_dir, "TodoSync.log")

def log_message(message):
    """Write a timestamped message to the log file."""
    try:
        log_path = get_log_path()
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write("[{}] {}\n".format(timestamp, message))
    except Exception as e:
        print("TodoSync: Error writing to log: {}".format(e))

# ------------------------------
# Helpers: File System & YAML
# ------------------------------

def get_paths(window):
    if not window:
        return None, None
    for folder in window.folders():
        if os.path.basename(folder) == TODO_PARENT_DIR:
            return (
                os.path.join(folder, TODO_YAML),
                os.path.join(folder, TODO_MD)
            )
    return None, None

def get_parent_folder(window):
    """Get the TODO_PARENT_DIR folder path for relative path calculation."""
    if not window:
        return None
    for folder in window.folders():
        if os.path.basename(folder) == TODO_PARENT_DIR:
            return folder
    return None

def migrate_database(db):
    """Migrate old database formats to current schema."""
    changed = False
    
    # Migrate open_todos
    for todo in db.get("open_todos", []):
        # Standardize to 'filepath'
        if 'filename' in todo and 'filepath' not in todo:
            todo['filepath'] = todo.pop('filename')
            changed = True
        elif 'source_file' in todo and 'filepath' not in todo:
            todo['filepath'] = todo.pop('source_file')
            changed = True
        
        # Ensure filepath exists
        if 'filepath' not in todo:
            todo['filepath'] = 'Unknown'
            changed = True
        
        # Ensure created_at exists
        if 'created_at' not in todo:
            todo['created_at'] = get_current_date()
            changed = True
    
    # Migrate resolved_todos
    for todo in db.get("resolved_todos", []):
        # Standardize to 'filepath'
        if 'filename' in todo and 'filepath' not in todo:
            todo['filepath'] = todo.pop('filename')
            changed = True
        elif 'source_file' in todo and 'filepath' not in todo:
            todo['filepath'] = todo.pop('source_file')
            changed = True
        
        # Ensure filepath exists
        if 'filepath' not in todo:
            todo['filepath'] = 'Unknown'
            changed = True
        
        # Ensure created_at exists
        if 'created_at' not in todo:
            todo['created_at'] = todo.get('resolved_at', get_current_date())
            changed = True
    
    return db, changed

def load_db(path):
    if not os.path.exists(path):
        return {"open_todos": [], "resolved_todos": []}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if "open_todos" not in data:
                data["open_todos"] = []
            if "resolved_todos" not in data:
                data["resolved_todos"] = []
            
            # Apply migration
            data, changed = migrate_database(data)
            if changed:
                print("TodoSync: Migrated database to current schema")
                save_db(path, data)
            
            return data
    except Exception as e:
        print("TodoSync: Error loading DB {}".format(e))
        return {"open_todos": [], "resolved_todos": []}

def save_db(path, data):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print("TodoSync: Error saving DB {}".format(e))

def get_current_date():
    return datetime.date.today().isoformat()

def get_display_name(filepath_relative):
    """Extract just the filename without extension for display."""
    if not filepath_relative or filepath_relative == "Unknown":
        return "Unknown"
    # Normalize the path and get just the basename
    try:
        normalized = os.path.normpath(filepath_relative)
        basename = os.path.basename(normalized)
        return os.path.splitext(basename)[0]
    except Exception as e:
        log_message("Error in get_display_name: {}".format(e))
        return "Unknown"

def is_urgent(text):
    """Check if a todo is marked urgent (starts with !)."""
    return text.strip().startswith('!')

# ------------------------------
# Core Logic: Sync View -> YAML
# ------------------------------

def sync_view_to_yaml(window, view):
    file_path = view.file_name()
    if not file_path or not file_path.lower().endswith(MARKDOWN_EXT):
        return 0
    
    for ex in EXCLUDE_FOLDERS:
        if ex in file_path.replace("\\", "/"):
            return 0

    yml_path, md_path = get_paths(window)
    if not yml_path:
        return 0
    
    if file_path == md_path:
        return 0

    parent_folder = get_parent_folder(window)
    if not parent_folder:
        return 0
    
    # Store relative path in YAML for uniqueness
    rel_path = os.path.relpath(file_path, parent_folder)
    rel_path_no_ext = os.path.splitext(rel_path)[0]
    
    db = load_db(yml_path)
    all_todos = db.get("open_todos", [])
    
    existing_texts = set(item['text'] for item in all_todos)

    symbols = view.indexed_symbol_regions()
    new_todos_count = 0
    seen_in_view = set()

    for sym_region in symbols:
        if sym_region.kind[1] == 't':
            text = sym_region.name.strip()
            
            if text in seen_in_view:
                continue
            seen_in_view.add(text)

            if text in existing_texts:
                continue
            
            entry = {
                "text": text,
                "filepath": rel_path_no_ext,  # Relative path without extension
                "created_at": get_current_date()
            }
            all_todos.append(entry)
            existing_texts.add(text)
            new_todos_count += 1

    # Check for urgency updates in existing todos from this file
    urgency_changed = False
    for existing_todo in all_todos:
        if existing_todo['filepath'] == rel_path_no_ext:
            db_text = existing_todo['text']

            # Skip if the exact text is still in the view (no change)
            if db_text in seen_in_view:
                continue

            # Check if urgency changed
            if db_text.startswith('!'):
                text_without = db_text[1:].strip()
                # If the non-urgent version is in the view
                if text_without in seen_in_view:
                    existing_todo['text'] = text_without
                    urgency_changed = True
            else:
                text_with = "! " + db_text
                # If the urgent version is in the view
                if text_with in seen_in_view:
                    existing_todo['text'] = text_with
                    urgency_changed = True

    if new_todos_count > 0 or urgency_changed:
        all_todos.sort(key=lambda x: (x['filepath'], x['created_at']))
        db["open_todos"] = all_todos
        save_db(yml_path, db)

    return new_todos_count

# ------------------------------
# Core Logic: Generate Report
# ------------------------------

def generate_markdown_report(window):
    log_message("=== generate_markdown_report START ===")
    
    yml_path, md_path = get_paths(window)
    if not yml_path or not md_path:
        log_message("ERROR: Could not get paths from window")
        return False
    
    log_message("yml_path: {}".format(yml_path))
    log_message("md_path: {}".format(md_path))

    try:
        db = load_db(yml_path)
        log_message("Database loaded successfully")
    except Exception as e:
        log_message("ERROR loading database: {}".format(e))
        return False
    
    todos = db.get("open_todos", [])
    log_message("Total todos loaded: {}".format(len(todos)))

    # Split into urgent and normal
    urgent_todos = [t for t in todos if is_urgent(t.get('text', ''))]
    normal_todos = [t for t in todos if not is_urgent(t.get('text', ''))]
    log_message("Urgent todos: {}, Normal todos: {}".format(len(urgent_todos), len(normal_todos)))

    # Extract existing 'Completed' section (Preserve History)
    completed_section = []
    if os.path.exists(md_path):
        try:
            with open(md_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                in_completed = False
                for line in lines:
                    if line.strip().lower().startswith("## completed"):
                        in_completed = True
                        completed_section.append(line)
                        continue
                    if in_completed:
                        if line.startswith("## ") and not line.startswith("## Completed"):
                            break
                        completed_section.append(line)
            log_message("Completed section extracted: {} lines".format(len(completed_section)))
        except Exception as e:
            log_message("ERROR reading completed section: {}".format(e))
    
    if not completed_section:
        completed_section = ["## Completed (History)\n", "\n"]
        log_message("No existing completed section, using default")

    output = []
    output.append("# Project TODOs\n\n")
    output.append("_Generated from todo.yml on {}_\n\n".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M')))

    # --- Urgent Section (NEW) ---
    if urgent_todos:
        log_message("Processing {} urgent todos".format(len(urgent_todos)))
        output.append("## Urgent\n\n")
        for t in urgent_todos:
            txt = t.get('text', 'Unknown')
            filepath = t.get('filepath', 'Unknown')
            try:
                display_name = get_display_name(filepath)
                line = "- [ ] {} [[{}]]".format(txt, display_name)
                output.append(line + "\n")
            except Exception as e:
                log_message("ERROR processing urgent todo: {} - {}".format(txt[:50], e))
        output.append("\n")

    # --- Inbox ---
    log_message("Processing inbox section")
    output.append("## Inbox\n\n")
    if not normal_todos:
        output.append("_No open todos._\n\n")
    else:
        for t in normal_todos:
            txt = t.get('text', 'Unknown')
            filepath = t.get('filepath', 'Unknown')
            try:
                display_name = get_display_name(filepath)
                line ="- [ ] {} [[{}]]".format(txt, display_name)
                output.append(line + "\n")
            except Exception as e:
                log_message("ERROR processing inbox todo: {} - {}".format(txt[:50], e))
    output.append("\n")

    # --- By File ---
    log_message("Processing by-file section")
    output.append("## By File\n\n")
    files_map = {}
    for t in todos:  # Include ALL todos (urgent + normal) in By File section
        filepath = t.get('filepath', 'Unknown')
        # Skip inbox-only todos (no real source file)
        if filepath == "todoUNDP":
            continue
        if filepath not in files_map:
            files_map[filepath] = []
        files_map[filepath].append(t)

    log_message("Files found: {}".format(len(files_map)))
    sorted_filepaths = sorted(files_map.keys())

    for filepath in sorted_filepaths:
        try:
            display_name = get_display_name(filepath)
            output.append("### {}\n".format(display_name))
            for t in files_map[filepath]:
                output.append("- [ ] {}\n".format(t.get('text', '')))
            output.append("\n")
        except Exception as e:
            log_message("ERROR processing file section: {} - {}".format(filepath, e))

    # --- Completed (Preserve Existing) ---
    log_message("Adding completed section: {} lines".format(len(completed_section)))
    output.extend(completed_section)

    try:
        with open(md_path, 'w', encoding='utf-8') as f:
            f.writelines(output)
        log_message("Report written successfully: {} lines".format(len(output)))
        log_message("=== generate_markdown_report END (SUCCESS) ===\n")
        return True
    except Exception as e:
        log_message("ERROR writing MD file: {}".format(e))
        log_message("=== generate_markdown_report END (FAILED) ===\n")
        return False

# ------------------------------
# Core Logic: Resolve Todos
# ------------------------------

def update_source_urgency(window, text_content, filepath_relative):
    """
    Updates the source file to mark a todo as urgent by adding ! prefix.
    """
    parent_folder = get_parent_folder(window)
    if not parent_folder:
        return False

    target_path = os.path.join(parent_folder, filepath_relative + ".md")

    if not os.path.exists(target_path):
        print("TodoSync Urgency: Source '{}' not found.".format(target_path))
        return False

    try:
        with open(target_path, 'r', encoding='utf-8') as f:
            content = f.read()

        escaped_text = re.escape(text_content)
        # Match: - [ ] text_content -> - [ ] ! text_content
        pattern = r"^(\s*[-*+]\s*\[\s*\]\s*)({})".format(escaped_text)
        new_content, count = re.subn(pattern, r"\1! \2", content, flags=re.MULTILINE)

        if count > 0:
            with open(target_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print("TodoSync Urgency: Updated source {}".format(target_path))
            return True
        else:
            print("TodoSync Urgency: Text '{}' not found in {}".format(text_content, target_path))
            return False

    except Exception as e:
        print("TodoSync Urgency: Error updating source: {}".format(e))
        return False

def remove_source_urgency(window, text_content_with_exclaim, filepath_relative):
    """
    Updates the source file to remove urgency by removing ! prefix.
    text_content_with_exclaim should include the "! " prefix.
    """
    parent_folder = get_parent_folder(window)
    if not parent_folder:
        return False

    target_path = os.path.join(parent_folder, filepath_relative + ".md")

    if not os.path.exists(target_path):
        print("TodoSync De-urgency: Source '{}' not found.".format(target_path))
        return False

    try:
        with open(target_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # text_content_with_exclaim is like "! Task"
        escaped_text = re.escape(text_content_with_exclaim)
        # Match: - [ ] ! Task -> - [ ] Task
        pattern = r"^(\s*[-*+]\s*\[\s*\]\s*)({})".format(escaped_text)

        # Remove the "! " part
        text_without = text_content_with_exclaim[1:].strip()  # Remove "!"
        new_content, count = re.subn(pattern, r"\1{}".format(text_without), content, flags=re.MULTILINE)

        if count > 0:
            with open(target_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print("TodoSync De-urgency: Updated source {}".format(target_path))
            return True
        else:
            print("TodoSync De-urgency: Text '{}' not found in {}".format(text_content_with_exclaim, target_path))
            return False

    except Exception as e:
        print("TodoSync De-urgency: Error updating source: {}".format(e))
        return False

def resolve_source_file(window, text_content, filepath_relative):
    """
    Attempts to find the source file by relative path and check the box in it.
    """
    parent_folder = get_parent_folder(window)
    if not parent_folder:
        return False
    
    # Reconstruct full path from relative path
    target_path = os.path.join(parent_folder, filepath_relative + ".md")
    
    if not os.path.exists(target_path):
        print("TodoSync Resolve: Source '{}' not found.".format(target_path))
        return False

    try:
        with open(target_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        escaped_text = re.escape(text_content)
        pattern = r"^(\s*[-*+]\s*\[)\s*(\]\s*{})".format(escaped_text)
        new_content, count = re.subn(pattern, r"\1x\2", content, flags=re.MULTILINE)

        if count > 0:
            with open(target_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print("TodoSync Resolve: Updated source {}".format(target_path))
            return True
        else:
            print("TodoSync Resolve: Text '{}' not found in {}".format(text_content, target_path))
            return False

    except Exception as e:
        print("TodoSync Resolve: Error updating source: {}".format(e))
        return False

def process_resolutions(window):
    yml_path, md_path = get_paths(window)
    if not yml_path or not md_path:
        return

    if not os.path.exists(md_path):
        return

    # 1. READ REPORT FILE
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    
    # 2. SEPARATE CONTENT FROM HISTORY
    active_lines = []
    history_lines = []
    in_history = False
    
    for line in lines:
        if line.strip().lower().startswith("## completed"):
            in_history = True
            history_lines.append(line + "\n")  # Preserve newline
            continue
        
        if in_history:
            history_lines.append(line + "\n")  # Preserve newline
        else:
            active_lines.append(line)

    # 3. SCAN ACTIVE LINES FOR CHECKS AND INBOX-ONLY TODOS
    # Pattern for checked: - [x] Text [[DisplayName]]?
    check_pattern = re.compile(r"^\s*[-*+]\s*\[[xX]\]\s*(.*?)((\[\[(.*?)\]\]))?\s*$")
    # Simplified pattern for inbox-only: - [ ] Text (NO wikilink at end)
    unchecked_pattern = re.compile(r"^\s*[-*+]\s*\[\s*\]\s*([^\[]+)$")
    
    resolved_in_report = [] 
    inbox_only_todos = []
    
    for line in active_lines:
        # Check for resolved items
        match_checked = check_pattern.match(line)
        if match_checked:
            raw_text = match_checked.group(1).strip()
            display_name = match_checked.group(4) # "Note" from [[Note]]
            
            resolved_in_report.append({
                "text": raw_text, 
                "display_name": display_name 
            })
            continue
        
        # Check for inbox-only unchecked todos (simplified regex)
        match_unchecked = unchecked_pattern.match(line)
        if match_unchecked:
            text_content = match_unchecked.group(1).strip()
            inbox_only_todos.append(text_content)

    # LOAD DATABASE
    db = load_db(yml_path)
    open_todos = db.get("open_todos", [])
    resolved_db_list = db.get("resolved_todos", [])
    
    newly_resolved_entries = []
    updates_made = False

    # 4. HANDLE INBOX-ONLY TODOS (Add to open_todos if not present)
    existing_texts = set(item['text'] for item in open_todos)

    for inbox_text in inbox_only_todos:
        # Check if this is an urgency update (starts with !)
        if inbox_text.startswith('!'):
            text_without_exclaim = inbox_text[1:].strip()
            # Check if the non-urgent version exists
            matching_entry = None
            for t in open_todos:
                if t['text'] == text_without_exclaim and t.get('filepath') == "todoUNDP":
                    matching_entry = t
                    break

            if matching_entry:
                # Update existing entry to be urgent
                matching_entry['text'] = inbox_text
                existing_texts.discard(text_without_exclaim)
                existing_texts.add(inbox_text)
                updates_made = True
                continue

        # Normal flow: add if not present
        if inbox_text not in existing_texts:
            entry = {
                "text": inbox_text,
                "filepath": "todoUNDP",  # Special marker for inbox-sourced
                "created_at": get_current_date()
            }
            open_todos.append(entry)
            existing_texts.add(inbox_text)
            updates_made = True

    # 4B. HANDLE URGENCY UPDATES FOR SOURCE-FILE TODOS
    # Pattern for unchecked with wikilink: - [ ] Text [[DisplayName]]
    unchecked_with_source_pattern = re.compile(r"^\s*[-*+]\s*\[\s*\]\s*(.*?)\s*\[\[(.*?)\]\]\s*$")

    for line in active_lines:
        match = unchecked_with_source_pattern.match(line)
        if match:
            text_content = match.group(1).strip()
            display_name = match.group(2).strip()

            # Only process if it starts with ! (potential urgency update)
            if text_content.startswith('!'):
                text_without_exclaim = text_content[1:].strip()

                # Find matching entry without !
                for t in open_todos:
                    if t['text'] == text_without_exclaim:
                        # Verify it matches by display name
                        if get_display_name(t.get('filepath')) == display_name:
                            # Update database
                            t['text'] = text_content
                            updates_made = True

                            # Update source file
                            source_filepath = t.get('filepath')
                            if source_filepath and source_filepath != "todoUNDP":
                                update_source_urgency(window, text_without_exclaim, source_filepath)
                            break

    # 4C. HANDLE DE-URGENCY UPDATES (Removing ! from urgent todos)
    # Check inbox-only todos
    for inbox_text in inbox_only_todos:
        # Check if this is NOT urgent (doesn't start with !)
        if not inbox_text.startswith('!'):
            # Check if the urgent version exists in database
            text_with_exclaim = "! " + inbox_text
            matching_entry = None
            for t in open_todos:
                if t['text'] == text_with_exclaim and t.get('filepath') == "todoUNDP":
                    matching_entry = t
                    break

            if matching_entry:
                # Update existing entry to be non-urgent
                matching_entry['text'] = inbox_text
                updates_made = True

    # Check source-file todos for de-urgency
    for line in active_lines:
        match = unchecked_with_source_pattern.match(line)
        if match:
            text_content = match.group(1).strip()
            display_name = match.group(2).strip()

            # Only process if it does NOT start with ! (potential de-urgency update)
            if not text_content.startswith('!'):
                text_with_exclaim = "! " + text_content

                # Find matching entry with !
                for t in open_todos:
                    if t['text'] == text_with_exclaim:
                        # Verify it matches by display name
                        if get_display_name(t.get('filepath')) == display_name:
                            # Update database
                            t['text'] = text_content
                            updates_made = True

                            # Update source file
                            source_filepath = t.get('filepath')
                            if source_filepath and source_filepath != "todoUNDP":
                                remove_source_urgency(window, text_with_exclaim, source_filepath)
                            break

    # 5. HANDLE RESOLUTIONS
    for item in resolved_in_report:
        text_to_resolve = item["text"]
        
        # FIND ORIGINAL IN DB
        original_entry = None
        for t in open_todos:
            if t['text'] == text_to_resolve:
                original_entry = t
                break
        
        if not original_entry:
            continue
            
        source_filepath = original_entry.get('filepath')
        created_at = original_entry.get('created_at', get_current_date())
        
        # A. Add to Resolved DB
        new_resolve_entry = {
            "text": text_to_resolve,
            "filepath": source_filepath,  # Store full relative path
            "created_at": created_at, 
            "resolved_at": get_current_date()
        }
        
        is_dupe_res = False
        for r in resolved_db_list:
            if r['text'] == text_to_resolve and r.get('filepath') == source_filepath:
                is_dupe_res = True
                break
        
        if not is_dupe_res:
             resolved_db_list.append(new_resolve_entry)
             newly_resolved_entries.append(new_resolve_entry)
             updates_made = True

             # B. Update Source File (if external)
             if source_filepath and source_filepath != "todoUNDP":
                 resolve_source_file(window, text_to_resolve, source_filepath)
    
    # 6. COMMIT CHANGES
    if updates_made:
        resolved_texts = set(r['text'] for r in resolved_in_report)
        open_todos = [t for t in open_todos if t['text'] not in resolved_texts]
        
        db["open_todos"] = open_todos
        db["resolved_todos"] = resolved_db_list
        save_db(yml_path, db)
        
        # 7. APPEND TO HISTORY
        if not history_lines:
            history_lines = ["## Completed (History)\n", "\n"]
        
        # Find the header and ensure there's a blank line after it
        header_index = 0
        insert_index = 1  # Default: insert after first line (which should be blank)
        
        for i, hline in enumerate(history_lines):
            if hline.strip().lower().startswith("## completed"):
                header_index = i
                # Check if next line is blank
                if i + 1 < len(history_lines):
                    if history_lines[i + 1].strip() == "":
                        insert_index = i + 2  # Insert after blank line
                    else:
                        insert_index = i + 1  # Insert right after header, will add blank
                        history_lines.insert(i + 1, "\n")  # Add missing blank line
                else:
                    # No line after header, add blank line
                    history_lines.append("\n")
                    insert_index = i + 2
                break
        
        insertion_lines = []
        for r in newly_resolved_entries:
            txt = r.get('text', '')
            filepath = r.get('filepath', '')
            date_res = r.get('resolved_at', '')
            
            if filepath and filepath != "todoUNDP":
                display_name = get_display_name(filepath)
                line = "- [x] {} [[{}]] <!-- resolved: {} -->\n".format(txt, display_name, date_res)
            else:
                line = "- [x] {} <!-- resolved: {} -->\n".format(txt, date_res)
            insertion_lines.append(line)
        
        history_lines[insert_index:insert_index] = insertion_lines

        with open(md_path, 'w', encoding='utf-8') as f:
            f.writelines(["# Temp\n"] + history_lines)
            
        generate_markdown_report(window)
        
        view = window.find_open_file(md_path)
        if view:
            view.run_command("revert")
            
        window.status_message("TodoSync: {} resolved, {} inbox items captured.".format(
            len(newly_resolved_entries), 
            len(inbox_only_todos)
        ))


# ------------------------------
# Event Listener
# ------------------------------

class TodoSyncListener(sublime_plugin.EventListener):
    _processing_todo = False
    _last_load_id = None
    _load_in_progress = {}  # Track views being loaded
    
    def on_post_save(self, view):
        file_path = view.file_name()
        w = view.window()
        if not w:
            return
            
        if file_path and file_path.endswith(TODO_MD):
            if os.path.basename(os.path.dirname(file_path)) == TODO_PARENT_DIR:
                if self._processing_todo:
                    return
                
                self._processing_todo = True
                try:
                    process_resolutions(w)
                    generate_markdown_report(w)
                    view.run_command("revert")
                finally:
                    self._processing_todo = False
                return
        
        sync_view_to_yaml(w, view)

    def on_pre_close(self, view):
        file_path = view.file_name()
        if file_path and not file_path.endswith(TODO_MD):
            w = view.window()
            if w:
                sync_view_to_yaml(w, view)
            
    def on_load(self, view):
        file_path = view.file_name()
        view_id = view.id()
        
        # Prevent duplicate processing
        if view_id == self._last_load_id:
            log_message("on_load: Skipping duplicate view_id {}".format(view_id))
            return
            
        if file_path and file_path.endswith(TODO_MD):
            if os.path.basename(os.path.dirname(file_path)) == TODO_PARENT_DIR:
                log_message("on_load: Processing TODO_MD file, view_id {}".format(view_id))
                self._last_load_id = view_id
                self._load_in_progress[view_id] = True
                
                def regenerate():
                    try:
                        log_message("on_load regenerate: Starting for view_id {}".format(view_id))
                        
                        # Check if view is still valid
                        if not view.is_valid():
                            log_message("on_load regenerate: View {} is invalid, aborting".format(view_id))
                            return
                        
                        # Check if view is loading
                        if view.is_loading():
                            log_message("on_load regenerate: View {} still loading, retrying in 200ms".format(view_id))
                            sublime.set_timeout_async(regenerate, 200)
                            return
                        
                        # Check if view is dirty (being edited)
                        if view.is_dirty():
                            log_message("on_load regenerate: View {} is dirty, skipping regeneration".format(view_id))
                            return
                        
                        log_message("on_load regenerate: Generating report for view_id {}".format(view_id))
                        success = generate_markdown_report(view.window())
                        
                        if success:
                            # Wait a bit before reverting to ensure file system is ready
                            def do_revert():
                                try:
                                    if view.is_valid() and not view.is_dirty():
                                        log_message("on_load regenerate: Reverting view_id {}".format(view_id))
                                        view.run_command("revert")
                                        log_message("on_load regenerate: Revert complete for view_id {}".format(view_id))
                                    else:
                                        log_message("on_load regenerate: Skipping revert, view invalid or dirty")
                                except Exception as e:
                                    log_message("ERROR in do_revert: {}".format(e))
                                finally:
                                    if view_id in self._load_in_progress:
                                        del self._load_in_progress[view_id]
                            
                            sublime.set_timeout(do_revert, 100)
                        else:
                            log_message("on_load regenerate: Report generation failed")
                            if view_id in self._load_in_progress:
                                del self._load_in_progress[view_id]
                            
                    except Exception as e:
                        log_message("ERROR in on_load regenerate: {}".format(e))
                        if view_id in self._load_in_progress:
                            del self._load_in_progress[view_id]
                
                sublime.set_timeout_async(regenerate, 100)

# ------------------------------
# Commands
# ------------------------------

class MarkdownCollectTodosCommand(sublime_plugin.WindowCommand):
    def run(self):
        generate_markdown_report(self.window)

class MarkdownResolveTodosCommand(sublime_plugin.WindowCommand):
    def run(self):
        process_resolutions(self.window)