from datetime import datetime #Variable items = MSI (Property nickname)
import os
import xlwings as xw

try:
    # Define the source Excel path
    excel_path = r"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates\MSI.xlsx"

    # Extract the base name of the Excel file (without extension)
    excel_file_name = os.path.splitext(os.path.basename(excel_path))[0]

    # Define the destination folder and dynamic date aspect
    current_date = datetime.now().strftime("%Y-%m-%d")
    output_folder = rf"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Details\Daily Detail {current_date}\Misc"

    # Define the output file name dynamically
    output_file = os.path.join(output_folder, f"Daily_Pickup_Summary_{excel_file_name}.html")

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Open the Excel file using XLWings
    app = xw.App(visible=False)
    workbook = app.books.open(excel_path)
    sheet = workbook.sheets[0]  # Access the first sheet; adjust if needed

    # Load the HTML template
    template_path = r"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates\Daily_Pickup_Summary_Template.html"
    with open(template_path, 'r') as file:
        html_template = file.read()

    # Map placeholders to Excel cell values (example for some placeholders)
    placeholder_map = {
        "{{value_r10_c2}}": sheet.range("B9").value,
        "{{value_r10_c3}}": int(sheet.range("C9").value) if sheet.range("C9").value else "0",
        "{{value_r10_c4}}": f"{round(sheet.range('D9').value * 100)}%" if sheet.range("D9").value else "0%",
        "{{value_r10_c5}}": f"${int(sheet.range('E9').value)}" if sheet.range("E9").value else "$0",
        "{{value_r10_c6}}": f"${int(sheet.range('F9').value)}" if sheet.range("F9").value else "$0",
        "{{value_r10_c8}}": sheet.range("H9").value,
        "{{value_r10_c9}}": int(sheet.range("I9").value) if sheet.range("I9").value else "0",
        "{{value_r10_c10}}": f"{round(sheet.range('J9').value * 100)}%" if sheet.range("J9").value else "0%",
        "{{value_r10_c11}}": f"${int(sheet.range('K9').value)}" if sheet.range("K9").value else "$0",
        "{{value_r10_c12}}": f"${int(sheet.range('L9').value)}" if sheet.range("L9").value else "$0",
        "{{value_r10_c14}}": sheet.range("N9").value,
        "{{value_r10_c15}}": int(sheet.range("O9").value) if sheet.range("O9").value else "0",
        "{{value_r10_c16}}": f"{round(sheet.range('P9').value * 100)}%" if sheet.range("P9").value else "0%",
        "{{value_r10_c17}}": f"${int(sheet.range('Q9').value)}" if sheet.range("Q9").value else "$0",
        "{{value_r10_c18}}": f"${int(sheet.range('R9').value)}" if sheet.range("R9").value else "$0",
        "{{value_r10_c19}}": f"${int(sheet.range('S9').value)}" if sheet.range("S9").value else "$0",
        "{{value_r10_c21}}": sheet.range("U9").value,
        "{{value_r10_c22}}": int(sheet.range("V9").value) if sheet.range("V9").value else "0",
        "{{value_r10_c23}}": f"{round(sheet.range('W9').value * 100)}%" if sheet.range("W9").value else "0%",
        "{{value_r10_c24}}": f"${int(sheet.range('X9').value)}" if sheet.range("X9").value else "$0",
        "{{value_r10_c25}}": f"${int(sheet.range('Y9').value)}" if sheet.range("Y9").value else "$0",
        "{{value_r10_c26}}": f"${int(sheet.range('Z9').value)}" if sheet.range("Z9").value else "$0",

        "{{value_r11_c2}}": sheet.range("B10").value,
        "{{value_r11_c3}}": int(sheet.range("C10").value) if sheet.range("C10").value else "0",
        "{{value_r11_c4}}": f"{round(sheet.range('D10').value * 100)}%" if sheet.range("D10").value else "0%",
        "{{value_r11_c5}}": f"${int(sheet.range('E10').value)}" if sheet.range("E10").value else "$0",
        "{{value_r11_c6}}": f"${int(sheet.range('F10').value)}" if sheet.range("F10").value else "$0",
        "{{value_r11_c8}}": sheet.range("H10").value,
        "{{value_r11_c9}}": int(sheet.range("I10").value) if sheet.range("I10").value else "0",
        "{{value_r11_c10}}": f"{round(sheet.range('J10').value * 100)}%" if sheet.range("J10").value else "0%",
        "{{value_r11_c11}}": f"${int(sheet.range('K10').value)}" if sheet.range("K10").value else "$0",
        "{{value_r11_c12}}": f"${int(sheet.range('L10').value)}" if sheet.range("L10").value else "$0",
        "{{value_r11_c14}}": sheet.range("N10").value,
        "{{value_r11_c15}}": int(sheet.range("O10").value) if sheet.range("O10").value else "0",
        "{{value_r11_c16}}": f"{round(sheet.range('P10').value * 100)}%" if sheet.range("P10").value else "0%",
        "{{value_r11_c17}}": f"${int(sheet.range('Q10').value)}" if sheet.range("Q10").value else "$0",
        "{{value_r11_c18}}": f"${int(sheet.range('R10').value)}" if sheet.range("R10").value else "$0",
        "{{value_r11_c19}}": f"${int(sheet.range('S10').value)}" if sheet.range("S10").value else "$0",
        "{{value_r11_c21}}": sheet.range("U10").value,
        "{{value_r11_c22}}": int(sheet.range("V10").value) if sheet.range("V10").value else "0",
        "{{value_r11_c23}}": f"{round(sheet.range('W10').value * 100)}%" if sheet.range("W10").value else "0%",
        "{{value_r11_c24}}": f"${int(sheet.range('X10').value)}" if sheet.range("X10").value else "$0",
        "{{value_r11_c25}}": f"${int(sheet.range('Y10').value)}" if sheet.range("Y10").value else "$0",
        "{{value_r11_c26}}": f"${int(sheet.range('Z10').value)}" if sheet.range("Z10").value else "$0",

        "{{value_r12_c2}}": sheet.range("B11").value,
        "{{value_r12_c3}}": int(sheet.range("C11").value) if sheet.range("C11").value else "0",
        "{{value_r12_c4}}": f"{round(sheet.range('D11').value * 100)}%" if sheet.range("D11").value else "0%",
        "{{value_r12_c5}}": f"${int(sheet.range('E11').value)}" if sheet.range("E11").value else "$0",
        "{{value_r12_c6}}": f"${int(sheet.range('F11').value)}" if sheet.range("F11").value else "$0",
        "{{value_r12_c8}}": sheet.range("H11").value,
        "{{value_r12_c9}}": int(sheet.range("I11").value) if sheet.range("I11").value else "0",
        "{{value_r12_c10}}": f"{round(sheet.range('J11').value * 100)}%" if sheet.range("J11").value else "0%",
        "{{value_r12_c11}}": f"${int(sheet.range('K11').value)}" if sheet.range("K11").value else "$0",
        "{{value_r12_c12}}": f"${int(sheet.range('L11').value)}" if sheet.range("L11").value else "$0",
        "{{value_r12_c14}}": sheet.range("N11").value,
        "{{value_r12_c15}}": int(sheet.range("O11").value) if sheet.range("O11").value else "0",
        "{{value_r12_c16}}": f"{round(sheet.range('P11').value * 100)}%" if sheet.range("P11").value else "0%",
        "{{value_r12_c17}}": f"${int(sheet.range('Q11').value)}" if sheet.range("Q11").value else "$0",
        "{{value_r12_c18}}": f"${int(sheet.range('R11').value)}" if sheet.range("R11").value else "$0",
        "{{value_r12_c19}}": f"${int(sheet.range('S11').value)}" if sheet.range("S11").value else "$0",
        "{{value_r12_c21}}": sheet.range("U11").value,
        "{{value_r12_c22}}": int(sheet.range("V11").value) if sheet.range("V11").value else "0",
        "{{value_r12_c23}}": f"{round(sheet.range('W11').value * 100)}%" if sheet.range("W11").value else "0%",
        "{{value_r12_c24}}": f"${int(sheet.range('X11').value)}" if sheet.range("X11").value else "$0",
        "{{value_r12_c25}}": f"${int(sheet.range('Y11').value)}" if sheet.range("Y11").value else "$0",
        "{{value_r12_c26}}": f"${int(sheet.range('Z11').value)}" if sheet.range("Z11").value else "$0",

        "{{value_r13_c2}}": sheet.range("B12").value,
        "{{value_r13_c3}}": int(sheet.range("C12").value) if sheet.range("C12").value else "0",
        "{{value_r13_c4}}": f"{round(sheet.range('D12').value * 100)}%" if sheet.range("D12").value else "0%",
        "{{value_r13_c5}}": f"${int(sheet.range('E12').value)}" if sheet.range("E12").value else "$0",
        "{{value_r13_c6}}": f"${int(sheet.range('F12').value)}" if sheet.range("F12").value else "$0",
        "{{value_r13_c8}}": sheet.range("H12").value,
        "{{value_r13_c9}}": int(sheet.range("I12").value) if sheet.range("I12").value else "0",
        "{{value_r13_c10}}": f"{round(sheet.range('J12').value * 100)}%" if sheet.range("J12").value else "0%",
        "{{value_r13_c11}}": f"${int(sheet.range('K12').value)}" if sheet.range("K12").value else "$0",
        "{{value_r13_c12}}": f"${int(sheet.range('L12').value)}" if sheet.range("L12").value else "$0",
        "{{value_r13_c14}}": sheet.range("N12").value,
        "{{value_r13_c15}}": int(sheet.range("O12").value) if sheet.range("O12").value else "0",
        "{{value_r13_c16}}": f"{round(sheet.range('P12').value * 100)}%" if sheet.range("P12").value else "0%",
        "{{value_r13_c17}}": f"${int(sheet.range('Q12').value)}" if sheet.range("Q12").value else "$0",
        "{{value_r13_c18}}": f"${int(sheet.range('R12').value)}" if sheet.range("R12").value else "$0",
        "{{value_r13_c19}}": f"${int(sheet.range('S12').value)}" if sheet.range("S12").value else "$0",
        "{{value_r13_c21}}": sheet.range("U12").value,
        "{{value_r13_c22}}": int(sheet.range("V12").value) if sheet.range("V12").value else "0",
        "{{value_r13_c23}}": f"{round(sheet.range('W12').value * 100)}%" if sheet.range("W12").value else "0%",
        "{{value_r13_c24}}": f"${int(sheet.range('X12').value)}" if sheet.range("X12").value else "$0",
        "{{value_r13_c25}}": f"${int(sheet.range('Y12').value)}" if sheet.range("Y12").value else "$0",
        "{{value_r13_c26}}": f"${int(sheet.range('Z12').value)}" if sheet.range("Z12").value else "$0",

        "{{value_r14_c2}}": sheet.range("B13").value,
        "{{value_r14_c3}}": int(sheet.range("C13").value) if sheet.range("C13").value else "0",
        "{{value_r14_c4}}": f"{round(sheet.range('D13').value * 100)}%" if sheet.range("D13").value else "0%",
        "{{value_r14_c5}}": f"${int(sheet.range('E13').value)}" if sheet.range("E13").value else "$0",
        "{{value_r14_c6}}": f"${int(sheet.range('F13').value)}" if sheet.range("F13").value else "$0",
        "{{value_r14_c8}}": sheet.range("H13").value,
        "{{value_r14_c9}}": int(sheet.range("I13").value) if sheet.range("I13").value else "0",
        "{{value_r14_c10}}": f"{round(sheet.range('J13').value * 100)}%" if sheet.range("J13").value else "0%",
        "{{value_r14_c11}}": f"${int(sheet.range('K13').value)}" if sheet.range("K13").value else "$0",
        "{{value_r14_c12}}": f"${int(sheet.range('L13').value)}" if sheet.range("L13").value else "$0",
        "{{value_r14_c14}}": sheet.range("N13").value,
        "{{value_r14_c15}}": int(sheet.range("O13").value) if sheet.range("O13").value else "0",
        "{{value_r14_c16}}": f"{round(sheet.range('P13').value * 100)}%" if sheet.range("P13").value else "0%",
        "{{value_r14_c17}}": f"${int(sheet.range('Q13').value)}" if sheet.range("Q13").value else "$0",
        "{{value_r14_c18}}": f"${int(sheet.range('R13').value)}" if sheet.range("R13").value else "$0",
        "{{value_r14_c19}}": f"${int(sheet.range('S13').value)}" if sheet.range("S13").value else "$0",
        "{{value_r14_c21}}": sheet.range("U13").value,
        "{{value_r14_c22}}": int(sheet.range("V13").value) if sheet.range("V13").value else "0",
        "{{value_r14_c23}}": f"{round(sheet.range('W13').value * 100)}%" if sheet.range("W13").value else "0%",
        "{{value_r14_c24}}": f"${int(sheet.range('X13').value)}" if sheet.range("X13").value else "$0",
        "{{value_r14_c25}}": f"${int(sheet.range('Y13').value)}" if sheet.range("Y13").value else "$0",
        "{{value_r14_c26}}": f"${int(sheet.range('Z13').value)}" if sheet.range("Z13").value else "$0",

        "{{value_r15_c2}}": sheet.range("B14").value,
        "{{value_r15_c3}}": int(sheet.range("C14").value) if sheet.range("C14").value else "0",
        "{{value_r15_c4}}": f"{round(sheet.range('D14').value * 100)}%" if sheet.range("D14").value else "0%",
        "{{value_r15_c5}}": f"${int(sheet.range('E14').value)}" if sheet.range("E14").value else "$0",
        "{{value_r15_c6}}": f"${int(sheet.range('F14').value)}" if sheet.range("F14").value else "$0",
        "{{value_r15_c8}}": sheet.range("H14").value,
        "{{value_r15_c9}}": int(sheet.range("I14").value) if sheet.range("I14").value else "0",
        "{{value_r15_c10}}": f"{round(sheet.range('J14').value * 100)}%" if sheet.range("J14").value else "0%",
        "{{value_r15_c11}}": f"${int(sheet.range('K14').value)}" if sheet.range("K14").value else "$0",
        "{{value_r15_c12}}": f"${int(sheet.range('L14').value)}" if sheet.range("L14").value else "$0",
        "{{value_r15_c14}}": sheet.range("N14").value,
        "{{value_r15_c15}}": int(sheet.range("O14").value) if sheet.range("O14").value else "0",
        "{{value_r15_c16}}": f"{round(sheet.range('P14').value * 100)}%" if sheet.range("P14").value else "0%",
        "{{value_r15_c17}}": f"${int(sheet.range('Q14').value)}" if sheet.range("Q14").value else "$0",
        "{{value_r15_c18}}": f"${int(sheet.range('R14').value)}" if sheet.range("R14").value else "$0",
        "{{value_r15_c19}}": f"${int(sheet.range('S14').value)}" if sheet.range("S14").value else "$0",
        "{{value_r15_c21}}": sheet.range("U14").value,
        "{{value_r15_c22}}": int(sheet.range("V14").value) if sheet.range("V14").value else "0",
        "{{value_r15_c23}}": f"{round(sheet.range('W14').value * 100)}%" if sheet.range("W14").value else "0%",
        "{{value_r15_c24}}": f"${int(sheet.range('X14').value)}" if sheet.range("X14").value else "$0",
        "{{value_r15_c25}}": f"${int(sheet.range('Y14').value)}" if sheet.range("Y14").value else "$0",
        "{{value_r15_c26}}": f"${int(sheet.range('Z14').value)}" if sheet.range("Z14").value else "$0",

        "{{value_r16_c2}}": sheet.range("B15").value,
        "{{value_r16_c3}}": int(sheet.range("C15").value) if sheet.range("C15").value else "0",
        "{{value_r16_c4}}": f"{round(sheet.range('D15').value * 100)}%" if sheet.range("D15").value else "0%",
        "{{value_r16_c5}}": f"${int(sheet.range('E15').value)}" if sheet.range("E15").value else "$0",
        "{{value_r16_c6}}": f"${int(sheet.range('F15').value)}" if sheet.range("F15").value else "$0",
        "{{value_r16_c8}}": sheet.range("H15").value,
        "{{value_r16_c9}}": int(sheet.range("I15").value) if sheet.range("I15").value else "0",
        "{{value_r16_c10}}": f"{round(sheet.range('J15').value * 100)}%" if sheet.range("J15").value else "0%",
        "{{value_r16_c11}}": f"${int(sheet.range('K15').value)}" if sheet.range("K15").value else "$0",
        "{{value_r16_c12}}":  f"${int(sheet.range('L15').value)}" if sheet.range("L15").value else "$0",
        "{{value_r16_c14}}": sheet.range("N15").value,
        "{{value_r16_c15}}": int(sheet.range("O15").value) if sheet.range("O15").value else "0",
        "{{value_r16_c16}}": f"{round(sheet.range('P15').value * 100)}%" if sheet.range("P15").value else "0%",
        "{{value_r16_c17}}": f"${int(sheet.range('Q15').value)}" if sheet.range("Q15").value else "$0",
        "{{value_r16_c18}}": f"${int(sheet.range('R15').value)}" if sheet.range("R15").value else "$0",
        "{{value_r16_c19}}": f"${int(sheet.range('S15').value)}" if sheet.range("S15").value else "$0",
        "{{value_r16_c21}}": sheet.range("U15").value,
        "{{value_r16_c22}}": int(sheet.range("V15").value) if sheet.range("V15").value else "0",
        "{{value_r16_c23}}": f"{round(sheet.range('W15').value * 100)}%" if sheet.range("W15").value else "0%",
        "{{value_r16_c24}}": f"${int(sheet.range('X15').value)}" if sheet.range("X15").value else "$0",
        "{{value_r16_c25}}": f"${int(sheet.range('Y15').value)}" if sheet.range("Y15").value else "$0",
        "{{value_r16_c26}}": f"${int(sheet.range('Z15').value)}" if sheet.range("Z15").value else "$0",

        "{{value_r17_c2}}": sheet.range("B16").value,
        "{{value_r17_c3}}": int(sheet.range("C16").value) if sheet.range("C16").value else "0",
        "{{value_r17_c4}}": f"{round(sheet.range('D16').value * 100)}%" if sheet.range("D16").value else "0%",
        "{{value_r17_c5}}": f"${int(sheet.range('E16').value)}" if sheet.range("E16").value else "$0",
        "{{value_r17_c6}}": f"${int(sheet.range('F16').value)}" if sheet.range("F16").value else "$0",
        "{{value_r17_c8}}": sheet.range("H16").value,
        "{{value_r17_c9}}": int(sheet.range("I16").value) if sheet.range("I16").value else "0",
        "{{value_r17_c10}}": f"{round(sheet.range('J16').value * 100)}%" if sheet.range("J16").value else "0%",
        "{{value_r17_c11}}": f"${int(sheet.range('K16').value)}" if sheet.range("K16").value else "$0",
        "{{value_r17_c12}}": f"${int(sheet.range('L16').value)}" if sheet.range("L16").value else "$0",
        "{{value_r17_c14}}": sheet.range("N16").value,
        "{{value_r17_c15}}": int(sheet.range("O16").value) if sheet.range("O16").value else "0",
        "{{value_r17_c16}}": f"{round(sheet.range('P16').value * 100)}%" if sheet.range("P16").value else "0%",
        "{{value_r17_c17}}": f"${int(sheet.range('Q16').value)}" if sheet.range("Q16").value else "$0",
        "{{value_r17_c18}}": f"${int(sheet.range('R16').value)}" if sheet.range("R16").value else "$0",
        "{{value_r17_c19}}": f"${int(sheet.range('S16').value)}" if sheet.range("S16").value else "$0",
        "{{value_r17_c21}}": sheet.range("U16").value,
        "{{value_r17_c22}}": int(sheet.range("V16").value) if sheet.range("V16").value else "0",
        "{{value_r17_c23}}": f"{round(sheet.range('W16').value * 100)}%" if sheet.range("W16").value else "0%",
        "{{value_r17_c24}}": f"${int(sheet.range('X16').value)}" if sheet.range("X16").value else "$0",
        "{{value_r17_c25}}": f"${int(sheet.range('Y16').value)}" if sheet.range("Y16").value else "$0",
        "{{value_r17_c26}}": f"${int(sheet.range('Z16').value)}" if sheet.range("Z16").value else "$0",

        "{{value_r18_c2}}": sheet.range("B17").value,
        "{{value_r18_c3}}": int(sheet.range("C17").value) if sheet.range("C17").value else "0",
        "{{value_r18_c4}}": f"{round(sheet.range('D17').value * 100)}%" if sheet.range("D17").value else "0%",
        "{{value_r18_c5}}": f"${int(sheet.range('E17').value)}" if sheet.range("E17").value else "$0",
        "{{value_r18_c6}}": f"${int(sheet.range('F17').value)}" if sheet.range("F17").value else "$0",
        "{{value_r18_c8}}": sheet.range("H17").value,
        "{{value_r18_c9}}": int(sheet.range("I17").value) if sheet.range("I17").value else "0",
        "{{value_r18_c10}}": f"{round(sheet.range('J17').value * 100)}%" if sheet.range("J17").value else "0%",
        "{{value_r18_c11}}": f"${int(sheet.range('K17').value)}" if sheet.range("K17").value else "$0",
        "{{value_r18_c12}}": f"${int(sheet.range('L17').value)}" if sheet.range("L17").value else "$0",
        "{{value_r18_c14}}": sheet.range("N17").value,
        "{{value_r18_c15}}": int(sheet.range("O17").value) if sheet.range("O17").value else "0",
        "{{value_r18_c16}}": f"{round(sheet.range('P17').value * 100)}%" if sheet.range("P17").value else "0%",
        "{{value_r18_c17}}": f"${int(sheet.range('Q17').value)}" if sheet.range("Q17").value else "$0",
        "{{value_r18_c18}}": f"${int(sheet.range('R17').value)}" if sheet.range("R17").value else "$0",
        "{{value_r18_c19}}": f"${int(sheet.range('S17').value)}" if sheet.range("S17").value else "$0",
        "{{value_r18_c21}}": sheet.range("U17").value,
        "{{value_r18_c22}}": int(sheet.range("V17").value) if sheet.range("V17").value else "0",
        "{{value_r18_c23}}": f"{round(sheet.range('W17').value * 100)}%" if sheet.range("W17").value else "0%",
        "{{value_r18_c24}}": f"${int(sheet.range('X17').value)}" if sheet.range("X17").value else "$0",
        "{{value_r18_c25}}": f"${int(sheet.range('Y17').value)}" if sheet.range("Y17").value else "$0",
        "{{value_r18_c26}}": f"${int(sheet.range('Z17').value)}" if sheet.range("Z17").value else "$0",

        "{{value_r19_c2}}": sheet.range("B18").value,
        "{{value_r19_c3}}": int(sheet.range("C18").value) if sheet.range("C18").value else "0",
        "{{value_r19_c4}}": f"{round(sheet.range('D18').value * 100)}%" if sheet.range("D18").value else "0%",
        "{{value_r19_c5}}": f"${int(sheet.range('E18').value)}" if sheet.range("E18").value else "$0",
        "{{value_r19_c6}}": f"${int(sheet.range('F18').value)}" if sheet.range("F18").value else "$0",
        "{{value_r19_c8}}": sheet.range("H18").value,
        "{{value_r19_c9}}": int(sheet.range("I18").value) if sheet.range("I18").value else "0",
        "{{value_r19_c10}}": f"{round(sheet.range('J18').value * 100)}%" if sheet.range("J18").value else "0%",
        "{{value_r19_c11}}": f"${int(sheet.range('K18').value)}" if sheet.range("K18").value else "$0",
        "{{value_r19_c12}}": f"${int(sheet.range('L18').value)}" if sheet.range("L18").value else "$0",
        "{{value_r19_c14}}": sheet.range("N18").value,
        "{{value_r19_c15}}": int(sheet.range("O18").value) if sheet.range("O18").value else "0",
        "{{value_r19_c16}}": f"{round(sheet.range('P18').value * 100)}%" if sheet.range("P18").value else "0%",
        "{{value_r19_c17}}": f"${int(sheet.range('Q18').value)}" if sheet.range("Q18").value else "$0",
        "{{value_r19_c18}}": f"${int(sheet.range('R18').value)}" if sheet.range("R18").value else "$0",
        "{{value_r19_c19}}": f"${int(sheet.range('S18').value)}" if sheet.range("S18").value else "$0",
        "{{value_r19_c21}}": sheet.range("U18").value,
        "{{value_r19_c22}}": int(sheet.range("V18").value) if sheet.range("V18").value else "0",
        "{{value_r19_c23}}": f"{round(sheet.range('W18').value * 100)}%" if sheet.range("W18").value else "0%",
        "{{value_r19_c24}}": f"${int(sheet.range('X18').value)}" if sheet.range("X18").value else "$0",
        "{{value_r19_c25}}": f"${int(sheet.range('Y18').value)}" if sheet.range("Y18").value else "$0",
        "{{value_r19_c26}}": f"${int(sheet.range('Z18').value)}" if sheet.range("Z18").value else "$0",

        "{{value_r20_c2}}": sheet.range("B19").value,
        "{{value_r20_c3}}": int(sheet.range("C19").value) if sheet.range("C19").value else "0",
        "{{value_r20_c4}}": f"{round(sheet.range('D19').value * 100)}%" if sheet.range("D19").value else "0%",
        "{{value_r20_c5}}": f"${int(sheet.range('E19').value)}" if sheet.range("E19").value else "$0",
        "{{value_r20_c6}}": f"${int(sheet.range('F19').value)}" if sheet.range("F19").value else "$0",
        "{{value_r20_c8}}": sheet.range("H19").value,
        "{{value_r20_c9}}": int(sheet.range("I19").value) if sheet.range("I19").value else "0",
        "{{value_r20_c10}}": f"{round(sheet.range('J19').value * 100)}%" if sheet.range("J19").value else "0%",
        "{{value_r20_c11}}": f"${int(sheet.range('K19').value)}" if sheet.range("K19").value else "$0",
        "{{value_r20_c12}}": f"${int(sheet.range('L19').value)}" if sheet.range("L19").value else "$0",
        "{{value_r20_c14}}": sheet.range("N19").value,
        "{{value_r20_c15}}": int(sheet.range("O19").value) if sheet.range("O19").value else "0",
        "{{value_r20_c16}}": f"{round(sheet.range('P19').value * 100)}%" if sheet.range("P19").value else "0%",
        "{{value_r20_c17}}": f"${int(sheet.range('Q19').value)}" if sheet.range("Q19").value else "$0",
        "{{value_r20_c18}}": f"${int(sheet.range('R19').value)}" if sheet.range("R19").value else "$0",
        "{{value_r20_c19}}": f"${int(sheet.range('S19').value)}" if sheet.range("S19").value else "$0",
        "{{value_r20_c21}}": sheet.range("U19").value,
        "{{value_r20_c22}}": int(sheet.range("V19").value) if sheet.range("V19").value else "0",
        "{{value_r20_c23}}": f"{round(sheet.range('W19').value * 100)}%" if sheet.range("W19").value else "0%",
        "{{value_r20_c24}}": f"${int(sheet.range('X19').value)}" if sheet.range("X19").value else "$0",
        "{{value_r20_c25}}": f"${int(sheet.range('Y19').value)}" if sheet.range("Y19").value else "$0",
        "{{value_r20_c26}}": f"${int(sheet.range('Z19').value)}" if sheet.range("Z19").value else "$0",

        "{{value_r21_c2}}": sheet.range("B20").value,
        "{{value_r21_c3}}": int(sheet.range("C20").value) if sheet.range("C20").value else "0",
        "{{value_r21_c4}}": f"{round(sheet.range('D20').value * 100)}%" if sheet.range("D20").value else "0%",
        "{{value_r21_c5}}": f"${int(sheet.range('E20').value)}" if sheet.range("E20").value else "$0",
        "{{value_r21_c6}}": f"${int(sheet.range('F20').value)}" if sheet.range("F20").value else "$0",
        "{{value_r21_c8}}": sheet.range("H20").value,
        "{{value_r21_c9}}": int(sheet.range("I20").value) if sheet.range("I20").value else "0",
        "{{value_r21_c10}}": f"{round(sheet.range('J20').value * 100)}%" if sheet.range("J2").value else "0%",
        "{{value_r21_c11}}": f"${int(sheet.range('K20').value)}" if sheet.range("K20").value else "$0",
        "{{value_r21_c12}}": f"${int(sheet.range('L20').value)}" if sheet.range("L20").value else "$0",
        "{{value_r21_c14}}": sheet.range("N20").value,
        "{{value_r21_c15}}": int(sheet.range("O20").value) if sheet.range("O20").value else "0",
        "{{value_r21_c16}}": f"{round(sheet.range('P20').value * 100)}%" if sheet.range("P20").value else "0%",
        "{{value_r21_c17}}": f"${int(sheet.range('Q20').value)}" if sheet.range("Q20").value else "$0",
        "{{value_r21_c18}}": f"${int(sheet.range('R20').value)}" if sheet.range("R20").value else "$0",
        "{{value_r21_c19}}": f"${int(sheet.range('S20').value)}" if sheet.range("S20").value else "$0",
        "{{value_r21_c21}}": sheet.range("U20").value,
        "{{value_r21_c22}}": int(sheet.range("V20").value) if sheet.range("V20").value else "0",
        "{{value_r21_c23}}": f"{round(sheet.range('W20').value * 100)}%" if sheet.range("W20").value else "0%",
        "{{value_r21_c24}}": f"${int(sheet.range('X20').value)}" if sheet.range("X20").value else "$0",
        "{{value_r21_c25}}": f"${int(sheet.range('Y20').value)}" if sheet.range("Y20").value else "$0",
        "{{value_r21_c26}}": f"${int(sheet.range('Z20').value)}" if sheet.range("Z20").value else "$0",
    
        # Add more mappings for other placeholders
    }

    # Replace placeholders in the template
    for placeholder, value in placeholder_map.items():
        html_template = html_template.replace(placeholder, str(value) if value is not None else "")

    # Save the populated HTML
    with open(output_file, 'w') as file:
        file.write(html_template)

    # Close the workbook and quit the app
    workbook.close()
    app.quit()

    print(f"Populated HTML saved to {output_file}")

except Exception as e:
    # Log the error and ensure the script continues
    print(f"An error occurred: {e}")
