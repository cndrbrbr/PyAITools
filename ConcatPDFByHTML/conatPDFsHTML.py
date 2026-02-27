# BeautifulSoup zum Parsen der HTML-Datei, pypdf zum Zusammenfügen der PDFs
# pip install beautifulsoup4 pypdf

import os
from bs4 import BeautifulSoup
from pypdf import PdfMerger


def extract_pdf_order_from_html(html_path):
    """
    Liest die HTML-Datei und extrahiert die PDF-Dateinamen
    in der dort angegebenen Reihenfolge.
    """
    with open(html_path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    pdf_files = []

    # Suche alle <a>-Tags mit PDF-Links
    for link in soup.find_all("a", href=True):
        href = link["href"]
        if href.lower().endswith(".pdf"):
            pdf_name = os.path.basename(href)
            pdf_files.append(pdf_name)

    return pdf_files


def merge_pdfs_in_order(folder_path, html_filename, output_filename):
    """
    Führt die PDFs entsprechend der Reihenfolge in der HTML-Datei zusammen.
    """
    html_path = os.path.join(folder_path, html_filename)

    if not os.path.exists(html_path):
        raise FileNotFoundError(f"HTML-Datei nicht gefunden: {html_path}")

    pdf_order = extract_pdf_order_from_html(html_path)

    if not pdf_order:
        raise ValueError("Keine PDF-Dateien in der HTML-Datei gefunden.")

    merger = PdfMerger()

    for pdf_name in pdf_order:
        pdf_path = os.path.join(folder_path, pdf_name)

        if not os.path.exists(pdf_path):
            print(f"⚠ WARNUNG: PDF nicht gefunden und wird übersprungen: {pdf_name}")
            continue

        print(f"Füge hinzu: {pdf_name}")
        merger.append(pdf_path)

    output_path = os.path.join(folder_path, output_filename)
    merger.write(output_path)
    merger.close()

    print(f"\n✅ Zusammengeführte PDF gespeichert unter:\n{output_path}")


if __name__ == "__main__":
    # >>> HIER ORDNERPFAD ANPASSEN <<<
    folder = r"E:\dev\_data\Akte_vollständig_305_C_337_23"
    html_file = "305_C_337_23.html"
    output_file = "Gesamtdatei.pdf"

    merge_pdfs_in_order(folder, html_file, output_file)