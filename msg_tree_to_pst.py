from pathlib import Path
import aspose.email as ae

ROOT_MSG = Path(r"C:\DEV\AI\EMLtoMSG\MSG")
OUT_DIR  = Path(r"C:\DEV\AI\EMLtoMSG\PST")
ADD_TOP_FOLDER = False  # True => put content under a single top folder inside each PST

def load_mapi_from_path(p: Path) -> ae.mapi.MapiMessage:
    """
    Load a .msg file as MapiMessage, compatible with multiple Aspose.Email builds.
    """
    MM = ae.mapi.MapiMessage
    # Newer API (if available)
    if hasattr(MM, "from_file"):
        return MM.from_file(str(p))
    # Fallback: load MailMessage -> convert to MAPI
    mail = ae.MailMessage.load(str(p))
    conv_opts = ae.mapi.MapiConversionOptions.unicode_format  # safe default
    return MM.from_mail_message(mail, conv_opts)

def get_or_create_subfolder(parent, name: str):
    for f in parent.get_sub_folders():
        if f.display_name == name:
            return f
    return parent.add_sub_folder(name)

def ensure_folder_chain(root_folder, parts):
    cur = root_folder
    for part in parts:
        cur = get_or_create_subfolder(cur, part)
    return cur

def add_all_msgs_to_pst(src_root: Path, pst_path: Path):
    src_root = src_root.resolve()
    pst_path.parent.mkdir(parents=True, exist_ok=True)

    all_msgs = list(src_root.rglob("*.msg"))
    print(f"[INFO] {src_root} -> found {len(all_msgs)} .msg files")
    if not all_msgs:
        print(f"[WARN] No .msg files under {src_root}")
        return 0

    with ae.storage.pst.PersonalStorage.create(
        str(pst_path),
        ae.storage.pst.FileFormatVersion.UNICODE
    ) as pst:
        base = get_or_create_subfolder(pst.root_folder, src_root.name) if ADD_TOP_FOLDER else pst.root_folder

        total = 0
        for msg_file in all_msgs:
            folder_parts = list(msg_file.relative_to(src_root).parts[:-1])  # folders only
            target_folder = ensure_folder_chain(base, folder_parts)
            try:
                mapi = load_mapi_from_path(msg_file)
                target_folder.add_message(mapi)
                total += 1
            except Exception as e:
                print(f"[WARN] Skipped {msg_file}: {e}")
        print(f"[OK] {total} messages added to {pst_path}")
        return total

if __name__ == "__main__":
    # Auto-detect the user folders (adjust if needed)
    candidates = {p.name.lower(): p for p in ROOT_MSG.iterdir() if p.is_dir()}
    user1 = candidates.get("user1")
    user2 = candidates.get("user2")

    if not (user1 and user2):
        print("[HINT] Could not auto-detect both user folders. Found under MSG:")
        for k in sorted(candidates): print(" -", k)
        raise SystemExit("Please set folder names explicitly.")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    add_all_msgs_to_pst(user1, OUT_DIR / f"{user1.name}.pst")
    add_all_msgs_to_pst(user2, OUT_DIR / f"{user2.name}.pst")

