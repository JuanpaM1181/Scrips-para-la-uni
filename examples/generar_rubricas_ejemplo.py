from uni_tools.rubrics import fill_rubric
from uni_tools.config import STUDENTS

template = "/home/juan/Documents/Crimenes de Guerra V9/C2/Programacion para moviles/evidencia/LISTA_COTEJO_ED1_UN2_ProgMoviles_II_2026.docx"
output_dir = "/tmp"

for s in STUDENTS:
    out = f"{output_dir}/Rubrica_{s['matricula']}_{s['name'].replace(' ', '_')}.docx"
    fill_rubric(template, out, s["name"], s["matricula"])
    print(f"Rubrica generada: {out}")
