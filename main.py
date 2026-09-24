import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUTPUT_DIR = "output"

def set_cell_background(cell, fill_hex):
    """Sets background color of a table cell."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets cell padding in dxa (1 pt = 20 dxa)."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_exam_doc(filename="VMC_Sample_Paper_Class10.docx"):
    doc = docx.Document()

    # 1. Page Margins (0.75 in)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 2. Document Title / Header
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_inst = title_p.add_run("VIDYAMANDIR CLASSES\n")
    run_inst.font.name = "Arial"
    run_inst.font.size = Pt(18)
    run_inst.font.bold = True
    run_inst.font.color.rgb = RGBColor(14, 56, 122)

    run_sub = title_p.add_run("IIT JEE | MEDICAL | FOUNDATION\n")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(11)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    run_paper = title_p.add_run("ADMISSION CUM SCHOLARSHIP TEST | 2-YEAR (JEE)\n")
    run_paper.font.name = "Arial"
    run_paper.font.size = Pt(13)
    run_paper.font.bold = True
    run_paper.font.color.rgb = RGBColor(30, 30, 30)

    run_desc = title_p.add_run("For Students Presently in Class 10 (Target: JEE 2029)\nDuration: 2 Hours   |   Maximum Marks: 205")
    run_desc.font.name = "Arial"
    run_desc.font.size = Pt(10)
    run_desc.font.italic = True

    # 3. Paper Scheme & Marking Scheme Callout Box
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, "F2F4F7")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)

    p_box = cell.paragraphs[0]
    r_bh = p_box.add_run("PAPER & MARKING SCHEME\n")
    r_bh.bold = True
    r_bh.font.size = Pt(10)
    r_bh.font.color.rgb = RGBColor(14, 56, 122)

    scheme_text = (
        "• Section I (Q. 1 to 5): Mental Aptitude (Single Choice MCQ). Marking: +4, -1.\n"
        "• Section II (Q. 6 to 15): Science (Single Choice MCQ). Marking: +5, -1.\n"
        "• Section III (Q. 16 to 30): Mathematics (Single Choice MCQ). Marking: +5, -1.\n"
        "• Section IV (Q. 1 to 10): Numerical Value Type (0 to 99). Marking: +6, -1."
    )
    r_bbody = p_box.add_run(scheme_text)
    r_bbody.font.size = Pt(9)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(255, 255, 255)
        p.paragraph_format.keep_with_next = True
        sec_tbl = doc.add_table(rows=1, cols=1)
        sec_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        sec_cell = sec_tbl.cell(0, 0)
        set_cell_background(sec_cell, "0E387A")
        set_cell_margins(sec_cell, top=60, bottom=60, left=120, right=120)
        sp = sec_cell.paragraphs[0]
        s_run = sp.add_run(title)
        s_run.bold = True
        s_run.font.size = Pt(11)
        s_run.font.color.rgb = RGBColor(255, 255, 255)
        doc.paragraphs[-2]._element.getparent().remove(doc.paragraphs[-2]._element)

    def add_mcq(q_num, text, options):
        qp = doc.add_paragraph()
        qp.paragraph_format.space_before = Pt(6)
        qp.paragraph_format.space_after = Pt(2)
        qp.paragraph_format.keep_with_next = True
        rq = qp.add_run(f"Q.{q_num}  {text}")
        rq.bold = True
        rq.font.size = Pt(10)

        op = doc.add_paragraph()
        op.paragraph_format.left_indent = Inches(0.35)
        op.paragraph_format.space_after = Pt(4)
        op.paragraph_format.keep_with_next = False
        opt_text = f"(A) {options[0]:<28} (B) {options[1]:<28}\n(C) {options[2]:<28} (D) {options[3]}"
        ro = op.add_run(opt_text)
        ro.font.size = Pt(9.5)

    def add_num_q(q_num, text):
        qp = doc.add_paragraph()
        qp.paragraph_format.space_before = Pt(6)
        qp.paragraph_format.space_after = Pt(4)
        rq = qp.add_run(f"Q.{q_num}  {text}")
        rq.bold = True
        rq.font.size = Pt(10)

    # ---------------- SECTION I ----------------
    add_section_header("SECTION - I [MENTAL APTITUDE] (5 Questions)")
    add_mcq(1, "Choose the odd letter-group out:", ["DWEV", "HSKP", "JQMN", "BYCX"])
    add_mcq(2, "In a certain code language, if 'TRAIN' is coded as 'WUDLQ', how will 'BUS' be coded in that same language?", ["EXV", "DWU", "EYV", "FXW"])
    add_mcq(3, "Pointing to a photograph of a boy, Suresh said, 'He is the only son of my mother's only son.' How is Suresh related to the boy?", ["Brother", "Uncle", "Father", "Grandfather"])
    add_mcq(4, "One evening before sunset, Rekha and Hema were talking to each other face to face. If Hema's shadow was exactly to the right of Hema, which direction was Rekha facing?", ["North", "South", "East", "West"])
    add_mcq(5, "If '+' means '÷', '-' means '×', '×' means '-', and '÷' means '+', then calculate the value of: 64 + 8 - 4 × 12 ÷ 6", ["26", "32", "20", "18"])

    # ---------------- SECTION II ----------------
    add_section_header("SECTION - II [SCIENCE] (10 Questions)")
    add_mcq(6, "A particle moves along a circular path of radius R. What is the magnitude of displacement after completing half a revolution?", ["0", "πR", "2R", "2πR"])
    add_mcq(7, "A force of 20 N acts on an object of mass 4 kg initially at rest for 3 seconds. The momentum gained by the object is:", ["60 kg·m/s", "15 kg·m/s", "240 kg·m/s", "80 kg·m/s"])
    add_mcq(8, "Two objects of masses m₁ and m₂ have equal kinetic energies. The ratio of their linear momenta p₁ : p₂ is:", ["m₁ : m₂", "m₂ : m₁", "√m₁ : √m₂", "m₁² : m₂²"])
    add_mcq(9, "The value of acceleration due to gravity (g) on the surface of the Earth depends on:", ["Mass of the falling body", "Radius and mass of the Earth", "Density of the falling body", "Surface area of the body"])
    add_mcq(10, "An object is placed at a distance of 20 cm in front of a concave mirror of focal length 15 cm. The nature and size of the image formed is:", ["Virtual, erect, magnified", "Real, inverted, diminished", "Real, inverted, magnified", "Virtual, erect, diminished"])
    add_mcq(11, "When zinc granules are treated with dilute sulfuric acid, the gas evolved burns with a:", ["Rotten egg smell", "Pungent odor", "Pop sound", "Brown fume"])
    add_mcq(12, "Bleaching powder is chemically represented as:", ["CaSO₄·½H₂O", "CaOCl₂", "NaHCO₃", "Na₂CO₃·10H₂O"])
    add_mcq(13, "The total number of valence electrons present in a phosphate ion (PO₄³⁻) is:", ["30", "32", "29", "8"])
    add_mcq(14, "The mass of 0.5 moles of water molecules (H₂O) is:", ["9 g", "18 g", "36 g", "4.5 g"])
    add_mcq(15, "Which of the following species is isoelectronic with a neon atom (Ne)?", ["Cl⁻", "Ca²⁺", "Mg²⁺", "K⁺"])

    # ---------------- SECTION III ----------------
    add_section_header("SECTION - III [MATHEMATICS] (15 Questions)")
    add_mcq(16, "If p and q are distinct prime numbers, what is the total number of positive factors of p²q?", ["4", "5", "6", "8"])
    add_mcq(17, "Which of the following fractions has the greatest value?", ["11/12", "17/18", "23/24", "35/36"])
    add_mcq(18, "If 3^(x+2) - 3^x = 72, then the value of (2x + 1) is:", ["3", "5", "7", "9"])
    add_mcq(19, "A merchant buys an article for Rs. 1200 and marks it 25% above the cost price. If he allows a discount of 10% on the marked price, his overall profit percentage is:", ["15%", "12.5%", "10%", "14%"])
    add_mcq(20, "Two successive discounts of 40% and 30% are equivalent to a single discount of:", ["70%", "58%", "62%", "55%"])
    add_mcq(21, "If (x + 2) is a factor of the polynomial P(x) = 2x³ + kx² - 4x + 8, then the value of k is:", ["0", "2", "-2", "4"])
    add_mcq(22, "The system of linear equations 2x + 3y = 7 and kx + 9y = 21 has infinitely many solutions. The value of k is:", ["3", "6", "9", "12"])
    add_mcq(23, "The sum of all interior angles of a regular octagon is equal to:", ["1080°", "900°", "1260°", "720°"])
    add_mcq(24, "In ΔABC, D and E are points on sides AB and AC respectively such that DE || BC. If AD = 3 cm, DB = 5 cm, and the area of ΔADE is 18 cm², what is the area of quadrilateral BDEC?", ["32 cm²", "110 cm²", "128 cm²", "50 cm²"])
    add_mcq(25, "The sides of a triangle are 9 cm, 12 cm, and 15 cm. The length of the altitude drawn to the hypotenuse is:", ["7.2 cm", "6.4 cm", "5.8 cm", "8.0 cm"])
    add_mcq(26, "If the n-th term of an arithmetic progression is given by aₙ = 5n - 3, what is the sum of the first 20 terms?", ["970", "990", "1020", "940"])
    add_mcq(27, "If the radius of a sphere is doubled, its surface area increases by:", ["100%", "200%", "300%", "400%"])
    add_mcq(28, "If α and β are the roots of the quadratic equation x² - 7x + 12 = 0, what is the value of (1/α + 1/β)?", ["7/12", "12/7", "-7/12", "1/7"])
    add_mcq(29, "If sin θ - cos θ = 0 with 0° < θ < 90°, then the value of (sin⁴θ + cos⁴θ) is:", ["1", "1/2", "1/4", "3/4"])
    add_mcq(30, "A card is drawn at random from a standard deck of 52 cards. The probability that the card drawn is neither a king nor a red card is:", ["6/13", "7/13", "19/26", "1/2"])

    # ---------------- SECTION IV ----------------
    add_section_header("SECTION - IV [NUMERICAL VALUE TYPE] (10 Questions)")
    add_num_q(1, "A monic polynomial P(x) of degree 4 satisfies P(1) = 2, P(2) = 4, P(3) = 6, and P(4) = 8. Determine the value of P(5).")
    add_num_q(2, "The roots of the polynomial equation x³ - 12x² + 47x - 60 = 0 represent the dimensions (length, breadth, height) of a rectangular box. Find the total surface area of this box.")
    add_num_q(3, "If α and β are the roots of the equation x² - 11x + 24 = 0 with α > β, calculate the numerical value of (2α + 5β).")
    add_num_q(4, "For 3 ≤ x ≤ 7, find the integer value of the expression: √(x + 1 - 4√(x - 3)) + √(x + 6 - 6√(x - 3)).")
    add_num_q(5, "In a right-angled triangle, the lengths of all three sides are integers. If the perimeter is 40 and the hypotenuse is 17, find the area of the triangle.")
    add_num_q(6, "ABCD is a square of side length 30 units. Points P, Q, R, S are chosen on sides AB, BC, CD, DA respectively such that AP = BQ = CR = DS = 10. The line segments AQ, BR, CS, DP are drawn, intersecting each other to bound an inner square. Find the area of this inner square.")
    add_num_q(7, "If (sin⁴θ)/2 + (cos⁴θ)/3 = 1/5, evaluate: 13·(tan²θ + cot²θ).")
    add_num_q(8, "Three mutually tangent circles of radius 4 lie flat on a table, with each circle tangent to the other two. A fourth smaller circle is placed in the central gap, touching all three circles externally. If the radius of this smaller circle is written in simplest form as a√b - c (where a, b, c are positive integers and b is square-free), find the value of (a + b + c).")
    add_num_q(9, "An arithmetic progression has first term a₁ = 3 and common difference d = 4. If the sum of the first n terms is equal to 210, find the value of n.")
    add_num_q(10, "From an external point P, a tangent segment PT of length 24 is drawn to a circle. A secant line through P intersects the circle at points A and B (with A lying between P and B). If AB = 20, find the length of the external segment PA.")

    # ---------------- ANSWER KEY TABLE ----------------
    doc.add_page_break()
    ak_h = doc.add_paragraph()
    ak_run = ak_h.add_run("ANSWER KEY & SCORING REFERENCE")
    ak_run.bold = True
    ak_run.font.size = Pt(13)
    ak_run.font.color.rgb = RGBColor(14, 56, 122)

    answers = [
        ("1", "(B)", "6", "(C)", "16", "(C)", "1", "34"),
        ("2", "(A)", "7", "(A)", "17", "(D)", "2", "94"),
        ("3", "(C)", "8", "(C)", "18", "(B)", "3", "31"),
        ("4", "(B)", "9", "(B)", "19", "(B)", "4", "1"),
        ("5", "(A)", "10", "(C)", "20", "(B)", "5", "60"),
        ("-", "-", "11", "(C)", "21", "(B)", "6", "360"),
        ("-", "-", "12", "(B)", "22", "(B)", "7", "37"),
        ("-", "-", "13", "(B)", "23", "(A)", "8", "15"),
        ("-", "-", "14", "(A)", "24", "(B)", "9", "10"),
        ("-", "-", "15", "(C)", "25", "(A)", "10", "16"),
        ("-", "-", "-", "-", "26", "(B)", "-", "-"),
        ("-", "-", "-", "-", "27", "(C)", "-", "-"),
        ("-", "-", "-", "-", "28", "(A)", "-", "-"),
        ("-", "-", "-", "-", "29", "(B)", "-", "-"),
        ("-", "-", "-", "-", "30", "(A)", "-", "-"),
    ]

    tbl_ans = doc.add_table(rows=1, cols=8)
    tbl_ans.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Q. (Sec I)", "Ans", "Q. (Sec II)", "Ans", "Q. (Sec III)", "Ans", "Q. (Sec IV)", "Ans"]
    hdr_cells = tbl_ans.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].text = h_text
        set_cell_background(hdr_cells[i], "0E387A")
        set_cell_margins(hdr_cells[i], top=80, bottom=80, left=60, right=60)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, row_data in enumerate(answers):
        row_cells = tbl_ans.add_row().cells
        bg = "F9FAFB" if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(row_data):
            row_cells[col_idx].text = val
            set_cell_background(row_cells[col_idx], bg)
            set_cell_margins(row_cells[col_idx], top=40, bottom=40, left=60, right=60)
            p = row_cells[col_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.size = Pt(8.5)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    full_path = os.path.join(OUTPUT_DIR, filename)
    doc.save(full_path)
    print(f"File successfully created: {full_path}")

if __name__ == "__main__":
    create_exam_doc()
