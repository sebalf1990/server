# -*- coding: utf-8 -*-
"""Snapshot de trazabilidad de la cuenta de pruebas (a@a.com) y sus personajes.

Vuelca inventario, banco, equipado y hechizos con nombres resueltos a un .md
timestamped en el work folder del plan, para detectar drift de items entre
cambios de recursos/migraciones. Reutilizable: correr cuando se quiera comparar.

Uso: python snapshot_test_account.py [email]   (default a@a.com)
"""
import os
import re
import sqlite3
import sys
from datetime import datetime

DB = r"c:\AO20\dev\server\database.db"
OBJ = r"c:\AO20\dev\Recursos\Dat\obj.dat"
HEC = r"c:\AO20\dev\Recursos\Dat\Hechizos.dat"
OUTDIR = r"c:\AO20\ia\work\2026\junio\13.001.cierre-backlog-diferido-10.001\snapshots"
EMAIL = sys.argv[1] if len(sys.argv) > 1 else "a@a.com"


def load_names(path, prefix, key="Name"):
    txt = open(path, "rb").read().decode("cp1252", errors="replace")
    out = {}
    for m in re.finditer(r"(?im)^\[%s(\d+)\]" % prefix, txt):
        seg = txt[m.end():m.end() + 500]
        g = re.search(r"(?im)^(?:%s|Nombre)=([^\r\n]+)" % key, seg)
        out[int(m.group(1))] = g.group(1).strip() if g else "?"
    return out


def main():
    obj_names = load_names(OBJ, "OBJ")
    hec_names = load_names(HEC, "HECHIZO", "Nombre")
    on = lambda i: obj_names.get(i, "(desconocido)")
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    cur = con.cursor()
    acc = cur.execute("SELECT * FROM account WHERE email = ?", (EMAIL,)).fetchone()
    if not acc:
        print("cuenta no encontrada:", EMAIL)
        return
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    os.makedirs(OUTDIR, exist_ok=True)
    out = [f"# Snapshot cuenta de pruebas `{EMAIL}` — {ts}",
           "",
           "Línea base de trazabilidad (plan 13.001). Re-generar con "
           "`dev/server/tools/snapshot_test_account.py` y diffear contra este archivo "
           "para detectar items que cambien/desaparezcan.",
           "",
           f"Cuenta id={acc['id']} email={EMAIL}", ""]
    chars = cur.execute("SELECT * FROM user WHERE account_id = ? ORDER BY id", (acc["id"],)).fetchall()
    out.append(f"## Personajes ({len(chars)})\n")
    for u in chars:
        out.append(f"### {u['name']}  (id={u['id']}, nivel {u['level']}, status {u['status']})")
        eq = []
        for slot, col in (("body", "body_id"), ("head", "head_id"), ("casco", "helmet_id"),
                          ("arma", "weapon_id"), ("escudo", "shield_id")):
            v = u[col] if col in u.keys() else 0
            if "backpack_id" in u.keys() and col == "shield_id":
                pass
            if v:
                label = on(v) if col in ("helmet_id", "weapon_id", "shield_id") else str(v)
                eq.append(f"{slot}={v}" + (f"({label})" if col in ("helmet_id", "weapon_id", "shield_id") else ""))
        if "backpack_id" in u.keys() and u["backpack_id"]:
            eq.append(f"backpack={u['backpack_id']}({on(u['backpack_id'])})")
        out.append("- equipado: " + (", ".join(eq) if eq else "(nada)"))
        inv = cur.execute("SELECT number, item_id, amount, is_equipped FROM inventory_item WHERE user_id = ? ORDER BY number", (u["id"],)).fetchall()
        out.append(f"- inventario ({len(inv)} items):")
        for it in inv:
            flag = " [E]" if it["is_equipped"] else ""
            out.append(f"    slot {it['number']:>2}: OBJ{it['item_id']} x{it['amount']} = {on(it['item_id'])}{flag}")
        bank = cur.execute("SELECT number, item_id, amount FROM bank_item WHERE user_id = ? ORDER BY number", (u["id"],)).fetchall()
        if bank:
            out.append(f"- banco ({len(bank)} items):")
            for it in bank:
                out.append(f"    slot {it['number']:>2}: OBJ{it['item_id']} x{it['amount']} = {on(it['item_id'])}")
        sp = cur.execute("SELECT number, spell_id FROM spell WHERE user_id = ? ORDER BY number", (u["id"],)).fetchall()
        if sp:
            hechizos = ", ".join(f"{s['spell_id']}({hec_names.get(s['spell_id'],'?')})" for s in sp)
            out.append(f"- hechizos ({len(sp)}): {hechizos}")
        out.append("")
    con.close()
    path = os.path.join(OUTDIR, f"snapshot_{EMAIL.replace('@','_at_')}_{ts}.md")
    open(path, "w", encoding="utf-8", newline="\n").write("\n".join(out) + "\n")
    print("snapshot:", path)
    print("\n".join(out[:4]))
    print("... personajes:", len(chars), "| rem inventario:",
          cur.execute if False else len([1]))


if __name__ == "__main__":
    main()
