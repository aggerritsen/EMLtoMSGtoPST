from pathlib import Path
import sys
import aspose.email as ae

def convert_eml_to_msg_tree(
    root: Path,
    in_place: bool = True,
    out_root: Path | None = None,
    unicode_msg: bool = True
):
    """
    Converteer alle .eml-bestanden onder 'root' naar .msg met behoud van mapstructuur.

    - in_place=True: schrijf .msg naast de .eml in dezelfde mappen.
    - in_place=False en out_root=<pad>: schrijf naar identieke boom onder out_root.
    - unicode_msg=True gebruikt Unicode .msg (aanbevolen); anders ANSI .msg.

    Bestandsnaam-conflicten worden opgelost door -1, -2, ... toe te voegen.
    """
    if not root.is_dir():
        raise SystemExit(f"Bronmap bestaat niet of is geen map: {root}")

    if not in_place:
        if out_root is None:
            raise SystemExit("Geef --out op als je niet in-place wilt schrijven.")
        out_root.mkdir(parents=True, exist_ok=True)

    # Optioneel: licentie instellen als je die hebt
    # lic = ae.License()
    # lic.set_license("Aspose.Email.lic")

    load_opts = ae.EmlLoadOptions()
    save_opts = ae.SaveOptions.default_msg_unicode if unicode_msg else ae.SaveOptions.default_msg

    def target_for(eml_path: Path) -> Path:
        if in_place:
            base = eml_path.with_suffix(".msg")
            return dedupe(base)
        else:
            rel = eml_path.relative_to(root).with_suffix(".msg")
            target = out_root.joinpath(rel)
            target.parent.mkdir(parents=True, exist_ok=True)
            return dedupe(target)

    def dedupe(p: Path) -> Path:
        if not p.exists():
            return p
        i = 1
        while True:
            cand = p.with_name(f"{p.stem}-{i}{p.suffix}")
            if not cand.exists():
                return cand
            i += 1

    converted = 0
    failed = 0

    for eml in root.rglob("*.eml"):
        try:
            mm = ae.MailMessage.load(str(eml), load_opts)
            tgt = target_for(eml)
            mm.save(str(tgt), save_opts)
            converted += 1
        except Exception as e:
            failed += 1
            print(f"[FOUT] {eml}: {e}")

    print(f"Klaar. Geconverteerd: {converted}, Fouten: {failed}")

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(
        description="Converteer EML → MSG met behoud van mapstructuur (Aspose.Email)."
    )
    ap.add_argument("root", help="Bronmap met .eml bestanden")
    ap.add_argument("--out", help="Uitvoermap (laat weg voor in-place)")
    ap.add_argument("--ansi", action="store_true", help="Schrijf ANSI .msg i.p.v. Unicode")
    args = ap.parse_args()

    root = Path(args.root)
    if args.out:
        convert_eml_to_msg_tree(root, in_place=False, out_root=Path(args.out), unicode_msg=not args.ansi)
    else:
        convert_eml_to_msg_tree(root, in_place=True, out_root=None, unicode_msg=not args.ansi)
cd