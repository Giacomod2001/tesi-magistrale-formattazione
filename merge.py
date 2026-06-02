import os
import re

brain_dir = r"C:\Users\giaco\.gemini\antigravity\brain\02c08dc5-12b2-41b1-9be7-a51f2bd1ce74"
files = [
    "introduzione.md",
    "capitolo1_introduzione.md",
    "capitolo2_contesto.md",
    "capitolo3_pilastri_teorici.md",
    "bibliografia.md"
]
output_file = r"C:\Users\giaco\iulm_formatter\tesi_completa.txt"

missing = []
written = []

try:
    with open(output_file, "w", encoding="utf-8") as outfile:
        for fname in files:
            fpath = os.path.join(brain_dir, fname)
            if os.path.exists(fpath):
                with open(fpath, "r", encoding="utf-8") as infile:
                    lines = infile.readlines()
                    for line in lines:
                        stripped = line.strip()
                        # Aggiunge # ai titoli principali se non ce l'hanno già
                        if stripped in ["Introduzione", "Struttura del lavoro"]:
                            if not stripped.startswith("#"):
                                outfile.write("# " + line)
                            else:
                                outfile.write(line)
                        elif stripped.startswith("PARTE ") or stripped.startswith("Capitolo "):
                            if not stripped.startswith("#"):
                                outfile.write("# " + line)
                            else:
                                outfile.write(line)
                        # Aggiunge ## ai sottocapitoli tipo 1.1 o 1.1.1
                        elif re.match(r'^\d+\.\d+(\.\d+)?\s+', stripped):
                            if not stripped.startswith("## "):
                                outfile.write("## " + line)
                            else:
                                outfile.write(line)
                        else:
                            outfile.write(line)
                    outfile.write("\n\n")
                written.append(fname)
            else:
                missing.append(fname)

    print(f"File scritti ({len(written)}): {', '.join(written)}")
    if missing:
        print(f"Attenzione! File non trovati ({len(missing)}): {', '.join(missing)}")
    print(f"Output: {output_file}")

except IOError as e:
    print(f"❌ Errore nella scrittura del file di output: {e}")
