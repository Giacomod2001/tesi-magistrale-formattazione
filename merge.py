import os

brain_dir = r"C:\Users\giaco\.gemini\antigravity\brain\02c08dc5-12b2-41b1-9be7-a51f2bd1ce74"
files = [
    "introduzione.md",
    "capitolo1_introduzione.md",
    "capitolo2_contesto.md",
    "capitolo3_pilastri_teorici.md",
    "bibliografia.md"
]

output_file = r"c:\Users\giaco\iulm_formatter\tesi_completa.txt"

with open(output_file, "w", encoding="utf-8") as outfile:
    for fname in files:
        fpath = os.path.join(brain_dir, fname)
        if os.path.exists(fpath):
            with open(fpath, "r", encoding="utf-8") as infile:
                outfile.write(infile.read())
                outfile.write("\n\n")
