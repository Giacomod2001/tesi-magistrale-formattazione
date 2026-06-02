import streamlit as st
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
import io
import re

def parse_inline_markdown(p, text):
    """
    Parsea il markdown inline per grassetto (**testo**) e corsivo (*testo*)
    e li aggiunge come run al paragrafo docx.
    """
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            run = p.add_run(part[2:-2])
            run.font.bold = True
        elif part.startswith('*') and part.endswith('*'):
            run = p.add_run(part[1:-1])
            run.font.italic = True
        else:
            p.add_run(part)

def format_iulm_doc(testo):
    # Crea un nuovo documento
    doc = docx.Document()
    
    # Imposta i margini (3 cm dx/sx, 2.5 cm top/bottom)
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(3.0)
        
    # Imposta lo stile Normal (Testo principale)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Garamond'
    font.size = Pt(12)
    style.paragraph_format.widow_control = True
    
    # Imposta lo stile Heading 1
    h1_style = doc.styles['Heading 1']
    h1_font = h1_style.font
    h1_font.name = 'Garamond'
    h1_font.size = Pt(20)
    h1_font.bold = True
    h1_font.color.rgb = RGBColor(0, 0, 0)
    h1_style.paragraph_format.keep_with_next = True
    
    # Imposta lo stile Heading 2
    h2_style = doc.styles['Heading 2']
    h2_font = h2_style.font
    h2_font.name = 'Garamond'
    h2_font.size = Pt(16)
    h2_font.bold = True
    h2_font.color.rgb = RGBColor(0, 0, 0)
    h2_style.paragraph_format.keep_with_next = True

    # Inserisci due pagine vuote iniziali per il frontespizio
    doc.add_page_break()
    doc.add_page_break()

    first_chapter = True

    lines = testo.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
            
        if line.startswith('|') and '|' in line[1:]:
            # Gestione Tabella Markdown
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            
            valid_rows = []
            for t_line in table_lines:
                if '---' in t_line:
                    continue
                cells = [cell.strip() for cell in t_line.split('|')[1:-1]]
                valid_rows.append(cells)
                
            if valid_rows:
                num_cols = len(valid_rows[0])
                table = doc.add_table(rows=len(valid_rows), cols=num_cols)
                table.style = 'Table Grid'
                for row_idx, row_data in enumerate(valid_rows):
                    for col_idx, cell_text in enumerate(row_data):
                        # Previeni IndexError se la riga ha meno colonne del previsto
                        if col_idx < num_cols:
                            cell = table.cell(row_idx, col_idx)
                            # Pulisce eventuali markdown inline nella tabella
                            cell_text_clean = cell_text.replace('**', '').replace('*', '')
                            cell.text = cell_text_clean
                            for p_cell in cell.paragraphs:
                                for r in p_cell.runs:
                                    r.font.name = 'Garamond'
                                    r.font.size = Pt(11)
                                    if row_idx == 0:
                                        r.font.bold = True
            continue
            
        if line.startswith('## '):
            # Titolo 2 (sottocapitolo)
            p = doc.add_heading('', level=2)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(line.replace('## ', ''))
            run.font.name = 'Garamond'
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif line.startswith('# '):
            # Titolo 1 (Capitolo)
            if not first_chapter:
                new_section = doc.add_section(WD_SECTION_START.ODD_PAGE)
                new_section.top_margin = Cm(2.5)
                new_section.bottom_margin = Cm(2.5)
                new_section.left_margin = Cm(3.0)
                new_section.right_margin = Cm(3.0)
            first_chapter = False
            
            p = doc.add_heading('', level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line.replace('# ', ''))
            run.font.name = 'Garamond'
            run.font.size = Pt(20)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        else:
            # Testo normale
            p = doc.add_paragraph()
            parse_inline_markdown(p, line)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            
        i += 1
            
    # Salva in un buffer di memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="IULM Thesis Formatter", page_icon="🎓", layout="wide")
st.title("🎓 IULM Thesis Formatter")
st.markdown("Incolla qui il tuo testo in formato grezzo. L'app creerà un file Word formattato con i parametri esatti IULM (Garamond 12, interlinea 1.5, margini 3cm/2.5cm, giustificato, con titoli grandi e in grassetto).")
st.markdown("**Tip:** Usa `# Nome Capitolo` per i titoli dei capitoli e `## Nome Sottocapitolo` per i paragrafi. Usa `**testo**` per il grassetto.")

col1, col2 = st.columns([1, 1])

with col1:
    testo_input = st.text_area("Incolla qui il tuo testo:", height=500)

    if st.button("🪄 Formatta e genera Word"):
        if testo_input.strip():
            with st.spinner("Generazione del documento in corso..."):
                docx_file = format_iulm_doc(testo_input)
                
                st.success("Documento generato con successo!")
                st.download_button(
                    label="📥 Scarica Tesi Formattata (.docx)",
                    data=docx_file,
                    file_name="Tesi_IULM_Formattata.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Inserisci del testo prima di generare il documento!")

with col2:
    st.markdown("### 👀 Preview del Testo Incollato")
    if testo_input.strip():
        st.markdown(testo_input, unsafe_allow_html=True)
    else:
        st.info("La preview apparirà qui non appena incollerai il testo a sinistra.")
