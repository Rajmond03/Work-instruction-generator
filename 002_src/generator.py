from pathlib import Path
from datetime import datetime
import os
import json

import pandas as pd
from PIL import Image
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
from openai import OpenAI

import tkinter as tk
from tkinter import messagebox, ttk

try:
    from deep_translator import GoogleTranslator
except ImportError:
    GoogleTranslator = None


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "001_output"
IMAGE_DIR = BASE_DIR / "003_images"
DATA_FILE = BASE_DIR / "004_data" / "WI-adat.xlsx"
TEMPLATE_FILE = BASE_DIR / "005_templates" / "WI-minta.docx"


def fajlnev_tisztitas(szoveg: str) -> str:
    tiltott = '<>:"/\\|?*'
    tiszta = "".join(c for c in szoveg if c not in tiltott)
    tiszta = "".join(c for c in tiszta if c.isalnum() or c in (" ", "_", "-", ".", "(", ")"))
    return tiszta.strip().replace(" ", "_")


def dokumentum_kod_generalas() -> str:
    return "GWP-HU-PR-WI-xxx"


def fajlnev_generalas(cim_hu: str) -> str:
    alap = "GWP-HU-PR-WI-xxx"
    tiszta = fajlnev_tisztitas(cim_hu)
    return f"{alap}_{tiszta}"


def cim_forditas(cim_hu: str) -> str:
    if not GoogleTranslator:
        return cim_hu
    try:
        return GoogleTranslator(source="hu", target="en").translate(cim_hu)
    except Exception:
        return cim_hu


def kep_elokeszites(mappa: Path, fajlnev_alap):
    """
    Visszatér:
    - kép elérési út
    - ideiglenes fájl-e
    """
    lehetseges_kiterjesztesek = [".jpg", ".jpeg", ".png", ".JPG", ".JPEG", ".PNG"]

    if pd.isna(fajlnev_alap):
        return None, False

    fajlnev_alap = str(fajlnev_alap).strip()

    for kiterjesztes in lehetseges_kiterjesztesek:
        utvonal = mappa / f"{fajlnev_alap}{kiterjesztes}"

        if utvonal.exists():
            if kiterjesztes.lower() == ".png":
                atmeneti_jpg = utvonal.with_name(f"{utvonal.stem}_temp.jpg")
                try:
                    with Image.open(utvonal) as img:
                        img.convert("RGB").save(atmeneti_jpg, "JPEG")
                    return atmeneti_jpg, True
                except Exception:
                    return utvonal, False

            return utvonal, False

    return None, False


def ellenorzes():
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not DATA_FILE.exists():
        raise FileNotFoundError(f"Nem található az Excel fájl:\n{DATA_FILE}")

    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"Nem található a Word sablon:\n{TEMPLATE_FILE}")

    if not IMAGE_DIR.exists():
        raise FileNotFoundError(f"Nem található a képmappa:\n{IMAGE_DIR}")


def excel_beolvasas():
    df = pd.read_excel(DATA_FILE)

    kotelezo_oszlopok = ["Sorszám", "Lépés", "Kép név"]
    hianyzo = [col for col in kotelezo_oszlopok if col not in df.columns]
    if hianyzo:
        raise ValueError(f"Hiányzó oszlop(ok) az Excelben: {', '.join(hianyzo)}")

    df = df.dropna(subset=["Sorszám", "Lépés"]).copy()
    df["Lépés"] = df["Lépés"].astype(str).str.strip()
    df = df[df["Lépés"] != ""]

    return df.reset_index(drop=True)


def sorszam_formazas(ertek) -> str:
    raw = str(ertek).strip()

    try:
        if raw.endswith("."):
            return raw
        return f"{int(float(ertek))}."
    except Exception:
        return raw


def lepes_objektum_keszites(doc, sor):
    szoveg = f"{sorszam_formazas(sor['Sorszám'])} {str(sor['Lépés']).strip()}"

    kep_utvonal, ideiglenes = kep_elokeszites(IMAGE_DIR, sor["Kép név"])
    kep = ""
    if kep_utvonal:
        kep = InlineImage(doc, str(kep_utvonal), width=Mm(77.5))

    return {
        "szoveg": szoveg,
        "kep": kep,
        "ideiglenes": kep_utvonal if ideiglenes else None
    }


def lepes_parok_keszites(doc, df):
    lepesek = []
    temp_fajlok = []

    for _, sor in df.iterrows():
        adat = lepes_objektum_keszites(doc, sor)
        lepesek.append(adat)
        if adat["ideiglenes"]:
            temp_fajlok.append(adat["ideiglenes"])

    parok = []
    for i in range(0, len(lepesek), 2):
        bal = lepesek[i]
        jobb = lepesek[i + 1] if i + 1 < len(lepesek) else None

        parok.append({
            "bal_szoveg": bal["szoveg"],
            "bal_kep": bal["kep"],
            "jobb_szoveg": jobb["szoveg"] if jobb else "",
            "jobb_kep": jobb["kep"] if jobb else "",
        })

    return parok, temp_fajlok


def folyamatleiras_szoveg_osszeallitasa(df) -> str:
    sorok = []
    for _, sor in df.iterrows():
        sorszam = sorszam_formazas(sor["Sorszám"])
        lepes = str(sor["Lépés"]).strip()
        sorok.append(f"{sorszam} {lepes}")
    return "\n".join(sorok)

def formatalt_felsorolas(lista) -> RichText:
    rt = RichText()

    for i, sor in enumerate(lista):
        if ":" in sor:
            cimke, leiras = sor.split(":", 1)
            rt.add("• ", font='Arial')
            rt.add(f"{cimke.strip()}:", bold=True, font='Arial')
            rt.add(f"{leiras}", font='Arial')
        else:
            rt.add(f"• {sor}", font='Arial')

        if i < len(lista) - 1:
            rt.add("\n")

    return rt

def ai_kovetelmenyek_generalasa(folyamat_szoveg: str) -> tuple[RichText, RichText, str]:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError(
            "Nincs beállítva az OPENAI_API_KEY környezeti változó. "
            "Állítsd be a Windowsban vagy a cmd-ben futtatás előtt."
        )

    client = OpenAI(api_key=api_key)

    prompt = f"""
Az alábbi munkautasítás folyamatlépései alapján készíts három külön részt magyar nyelven:

1) Minőségi követelmények
2) Balesetvédelmi / munkavédelmi szempontok
3) Szükséges eszközök

Szabályok:
- rövid, szakmai, gyakorlatias megfogalmazás
- ne írj mellébeszélést
- a konkrét folyamatból indulj ki
- a minőségi és balesetvédelmi részek felsorolás legyenek
- minden felsorolási elem így kezdődjön: Rövid címke: magyarázat
- például: Tűzvédelem: ...
- ne hivatkozz konkrét jogszabályra
- ne találj ki irreális technológiai részleteket
- a szükséges eszközök csak konkrét eszköznevek legyenek
- ne írj általános kategóriát (pl. "szerszám")
- az eszközök rövid megnevezések legyenek
- NE használj ```json vagy ``` jelölést
- kizárólag JSON-t adj vissza

Folyamatlépések:
{folyamat_szoveg}

A visszatérési formátum pontosan ez legyen:
{{
  "minosegi_kovetelmenyek": ["...", "..."],
  "balesetvedelmi_szempontok": ["...", "..."],
  "szukseges_eszkozok": ["...", "..."]
}}
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt
    )

    raw_text = response.output_text.strip()

    if "```" in raw_text:
        raw_text = raw_text.split("```")[1]
        raw_text = raw_text.replace("json", "", 1).strip()

    try:
        adat = json.loads(raw_text)
    except json.JSONDecodeError as e:
        raise ValueError(
            "Az AI válasza nem volt értelmezhető JSON.\n"
            f"Nyers válasz:\n{raw_text}"
        ) from e

    minosegi_lista = adat.get("minosegi_kovetelmenyek", [])
    balesetvedelmi_lista = adat.get("balesetvedelmi_szempontok", [])
    szukseges_lista = adat.get("szukseges_eszkozok", [])

    if (
        not isinstance(minosegi_lista, list)
        or not isinstance(balesetvedelmi_lista, list)
        or not isinstance(szukseges_lista, list)
    ):
        raise ValueError("Az AI válasz szerkezete hibás.")

    minosegi_kovetelmenyek = formatalt_felsorolas(minosegi_lista)
    balesetvedelmi_szempontok = formatalt_felsorolas(balesetvedelmi_lista)
    szukseges_eszkozok = ", ".join(szukseges_lista)

    return minosegi_kovetelmenyek, balesetvedelmi_szempontok, szukseges_eszkozok


def dokumentum_generalas(cim_hu: str, keszito: str):
    ellenorzes()

    mai_datum = datetime.now().strftime("%Y.%m.%d.")
    cim_en = cim_forditas(cim_hu)
    dokumentum_kod = dokumentum_kod_generalas()
    fajlnev_alap = fajlnev_generalas(cim_hu)

    doc = DocxTemplate(str(TEMPLATE_FILE))
    df = excel_beolvasas()

    folyamat_szoveg = folyamatleiras_szoveg_osszeallitasa(df)
    minosegi_kovetelmenyek, balesetvedelmi_szempontok, szukseges_eszkozok = ai_kovetelmenyek_generalasa(folyamat_szoveg)

    lepes_parok, temp_fajlok = lepes_parok_keszites(doc, df)

    context = {
        "cim_magyar": cim_hu,
        "cim_angol": cim_en,
        "keszito_neve": keszito,
        "mai_datum": mai_datum,
        "lepes_parok": lepes_parok,
        "dokumentum_kod": dokumentum_kod,
        "minosegi_kovetelmenyek": minosegi_kovetelmenyek,
        "balesetvedelmi_szempontok": balesetvedelmi_szempontok,
        "szukseges_eszkozok": szukseges_eszkozok
    }

    fajlnev = fajlnev_alap
    mentett_fajl = OUTPUT_DIR / f"{fajlnev}.docx"

    try:
        doc.render(context)
        doc.save(str(mentett_fajl))
    finally:
        for temp_fajl in temp_fajlok:
            try:
                if temp_fajl.exists():
                    temp_fajl.unlink()
            except Exception:
                pass

    return mentett_fajl


def inditas_generalas(cim_hu, keszito, ablak):
    if not cim_hu.strip() or not keszito.strip():
        messagebox.showwarning("Hiányzó adat", "Kérlek, töltsd ki mindkét mezőt!")
        return

    try:
        mentett_fajl = dokumentum_generalas(
            cim_hu.strip(),
            keszito.strip()
        )
        messagebox.showinfo("Siker", f"A dokumentum elkészült:\n{mentett_fajl}")
        ablak.destroy()
    except Exception as e:
        messagebox.showerror("Hiba", f"Hiba történt a generálás során:\n{e}")


def interface():
    root = tk.Tk()
    root.title("WI Generátor")
    root.geometry("450x360")
    root.resizable(False, False)

    window_width = 450
    window_height = 360
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

    tk.Label(
        root,
        text="WI Automata Kitöltő",
        font=("Arial", 14, "bold"),
        fg="#2c3e50"
    ).pack(pady=15)

    tk.Label(root, text="Dokumentum címe:", font=("Arial", 10)).pack()
    entry_cim = tk.Entry(root, width=50)
    entry_cim.pack(pady=5)
    entry_cim.focus()

    tk.Label(root, text="Készítő neve:", font=("Arial", 10)).pack()
    entry_nev = tk.Entry(root, width=50)
    entry_nev.pack(pady=5)

    style = ttk.Style()
    style.configure("Green.TButton", font=("Arial", 11, "bold"), foreground="black")

    btn = ttk.Button(
        root,
        text="DOKUMENTUM GENERÁLÁS",
        style="Green.TButton",
        command=lambda: inditas_generalas(
            entry_cim.get(),
            entry_nev.get(),
            root
        )
    )
    btn.pack(pady=25)

    root.mainloop()


if __name__ == "__main__":
    interface()