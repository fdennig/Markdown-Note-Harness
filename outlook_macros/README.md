# Outlook to Obsidian Macro

This project automates saving Outlook emails as **Markdown notes** in an **Obsidian vault**, inspired by [Obsidian-For-Business](https://github.com/tallguyjenks/Obsidian-For-Business).

## 🚀 Features
✅ Extracts emails as **Markdown** with structured **YAML frontmatter** (`toreply`, `personal`, `priority`, `class`, `subject`, `date`, `from`, `to`, `attachments`)  
✅ Saves emails **directly to your Obsidian vault**  
✅ Task integration: Adds `- [ ] email title` at the top of the note  
✅ **Macro Variants**:
   - `ExtractEmail`: Standard email export (`toreply: `, `personal: `, `priority: `)
   - `ExtractEmailToReply`: Sets `toreply: true`
   - `ExtractEmailWithAttachments`: Saves attachments to local folder and links them
   - `ExtractEmailPersonal`: Sets `personal: true`
   - `ExtractEmailHighPriority`: Sets `priority: high`
✅ **Automatically opens** the newly created note in your editor  


## 📂 Installation
To install and use the macro, follow these steps:

### **1️⃣ Enable Macros in Outlook**
1. Open Outlook.
2. Go to **File → Options → Trust Center → Trust Center Settings**.
3. Click **Macro Settings → Enable all macros**.

### **2️⃣ Enable Required References in VBA**
To allow the macro to run correctly, you need to enable some VBA libraries:

1. In the VBA editor, go to **Tools → References**.
2. Find and enable:
   - ✅ **Microsoft Forms 2.0 Object Library**
   - ✅ **Microsoft VBScript Regular Expressions 5.5**

### **3️⃣ Open the Outlook VBA Editor**
1. Press `Alt + F11` to open **Outlook's VBA Editor**.
2. In the VBA editor, go to **Insert → Module**.
3. **Create three modules** and name them exactly:
   - `SaveEmail`
   - `SaveUtilities`
   - `USER_CONFIG`
4. Copy and paste the corresponding `.vb` file contents into each module.
5. Modify the **vault path** in `USER_CONFIG.vb` to match your Obsidian setup:

```vb
vaultPathToSaveFileTo = "C:\Users\YourUsername\Obsidian\Vault\Emails\"
```
Make sure the path ends with a backslash.

## ⚠️ Limitations

- No Attachment Support: This macro does not currently save or link email attachments.
- No Meeting / Calendar Support: It only works with standard mail items (Class = 43). Meetings or calendar invites are not handled.
- Outlook Desktop on Windows Only: Tested only on Windows versions of Outlook with VBA. Other platforms (Mac, Outlook Web) don’t support these VBA macros.
- Conversation-View Stack: You must select an individual email rather than a conversation stack.

## 🔗 Credits

This project builds upon Obsidian-For-Business by @tallguyjenks.



