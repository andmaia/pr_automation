import re
import pdfplumber
import io


def extract_text_pdf(pdf_bytes: bytes):
    texto_completo = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                texto_completo += "\n" + page_text

    return texto_completo

def extract_text_between_keywords(text, start_keyword, end_keyword):
    start = re.escape(start_keyword)
    end = re.escape(end_keyword)

    pattern = re.compile(f'{start}.*?{end}', re.DOTALL)
    match = re.search(pattern, text)

    if not match:
        return []

    extract = match.group()
    lines = extract.split('\n')

    transactions = [line.split() for line in lines if line.strip()]
    transactions = transactions[1:-1]
    return transactions




def verify_pattern_returned_from_pdf(texto):
    padrao = r'(\d+[.,]?\d{0,2})([CD])(\w)'

    # Substitui o padrão encontrado por preço + C/D + próxima letra por preço + C/D + quebra de linha + próxima letra
    texto_corrigido = re.sub(padrao, r'\1\2\n\3', texto)

    return texto_corrigido

   

