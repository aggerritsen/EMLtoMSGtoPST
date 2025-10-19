
# EML to MSG and PST Conversion Utility

This project provides Python scripts for converting email archives between formats and preserving folder structures. It is designed for **batch conversion of .eml files into .msg** and subsequently building **.pst archives** per user.

---

## 📦 Project Structure

```
...\EMLtoMSG
│
├── EML\                 # Source .eml folders (input)
│   ├── user1│   └── user2│
├── MSG\                 # Converted .msg folders (output)
│   ├── user1│   └── user2│
├── PST\                 # Final .pst archives
│   ├── user1.pst
│   └── user2.pst
│
├── eml_to_msg_tree.py   # EML → MSG conversion script
└── msg_tree_to_pst.py   # MSG → PST creation script
```
---

## 🧰 Requirements

- Python 3.10 or higher  
- [Aspose.Email for Python via .NET](https://pypi.org/project/Aspose.Email-for-Python-via-NET/)

Install dependencies:

```bash
pip install Aspose.Email-for-Python-via-NET
```

---

## ⚙️ Usage

### 1. Convert all EML to MSG

Converts every `.eml` file under `EML/` into `.msg`, preserving all subfolders.

```bash
python eml_to_msg_tree.py ".\EML" --out ".\MSG"
```

After completion, verify that the `.msg` structure mirrors the original `.eml` folders.

---

### 2. Build PST Archives per User

Creates Outlook-compatible `.pst` files for **user1** and **user2**, preserving their full folder trees.

```bash
python msg_tree_to_pst.py
```

Output:

```
...\EMLtoMSG\PST\user1.pst
...\EMLtoMSG\PST\user2.pst
```

Open these PST files in Outlook via:  
**File → Open & Export → Open Outlook Data File (.pst)**

---

## 🧑‍💻 Developer Notes

### EML → MSG conversion

The `eml_to_msg_tree.py` script uses Aspose.Email’s `MailMessage.load()` and `MailMessage.save()` methods to perform high-fidelity conversion to `.msg` format while retaining message metadata and attachments.

```python
msg = ae.MailMessage.load(str(eml_path), ae.EmlLoadOptions())
msg.save(str(target_path), ae.SaveOptions.default_msg_unicode)
```

### MSG → PST creation

The `msg_tree_to_pst.py` script walks through all `.msg` files and builds a hierarchical PST file using Aspose’s `PersonalStorage` API.

```python
pst = ae.storage.pst.PersonalStorage.create(path, ae.storage.pst.FileFormatVersion.UNICODE)
folder = pst.root_folder.add_sub_folder("Inbox")
mapi = ae.mapi.MapiMessage.from_file(msg_file)
folder.add_message(mapi)
```

It automatically handles Unicode PSTs and reconstructs folder structures recursively.

### Compatibility

Tested successfully on:
- Windows 11 (Python 3.11)
- Aspose.Email for Python via .NET 24.x

---

## 👩‍🏫 User Guide

### How to Use

1. Place all `.eml` files under `EML\<username>\...`
2. Run the EML → MSG script to create `.msg` copies.
3. Run the MSG → PST script to generate `.pst` archives.
4. Import `.pst` files into Outlook.

### Troubleshooting

- **from_file not found:** Update Aspose.Email package.
- **Skipped messages:** Likely invalid `.msg` files or permission issues.

### Backup

Keep copies of the original `.eml` sources before conversion. PSTs can be re-generated anytime from `.msg` sources.

---

## 🧾 License

This utility is open for internal or research use.  
Commercial redistribution requires an Aspose.Email license.

---
