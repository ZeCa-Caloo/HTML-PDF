
"""
Uso (Windows + MS Office):
    python excel_to_pdf_win.py "C:\caminho\arquivo.xlsx" "C:\saida\arquivo.pdf"
"""
import sys, os
import win32com.client as win32

def export_excel_to_pdf(xlsx_path, pdf_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(xlsx_path)
    try:
        # 0 = xlTypePDF
        wb.ExportAsFixedFormat(0, pdf_path)
    finally:
        wb.Close(False)
        excel.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python excel_to_pdf_win.py <arquivo.xlsx> <saida.pdf>")
        sys.exit(1)
    xlsx = os.path.abspath(sys.argv[1])
    pdf = os.path.abspath(sys.argv[2])
    export_excel_to_pdf(xlsx, pdf)
    print("PDF gerado em:", pdf)
