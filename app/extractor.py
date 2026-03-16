import pdfplumber
import re
import io
from datetime import datetime
import pandas as pd


class Extractor:

    def extract_sales_per_payment_block(self, text):
        pattern = (
            r"Plano de Pagamento %Total Qtde\. Vendas Valor Total Vendas\s*\n"
            r"((?:.*\n)*?)"
            r"(?=100,00%%? \d+ R\$ [\d\.,]+\s*$)"
        )
        match = re.search(pattern, text, re.MULTILINE)
        return match.group(1).strip() if match else ""

    def extract_discount_block(self, text):
        pattern = (
            r"Saídas\s+-R\$ [\d\.,]+\s*\n"
            r"((?:.*\n)*?)"
            r"^Valor Total dos Itens no Caixa:.*?$"
        )
        match = re.search(pattern, text, re.MULTILINE)
        return match.group(1).strip() if match else ""

    def extract_sales_per_salesman_block(self, text):
        pattern = (
            r"Funcionário Qtde\.Produtos Percentual Valor Total Vendas\s*\n"
            r"((?:.*\n)*?)"
            r"^Valor Total das Vendas: R\$ [\d\.,]+"
        )
        match = re.search(pattern, text, re.MULTILINE)
        return match.group(1).strip() if match else ""

    def _catch_date(self, text):
        match = re.search(r"Movimento de Caixa de (\d{2}/\d{2}/\d{4})", text)
        return match.group(1) if match else None

    def _add_date_to_block(self, block, date):
        return "\n".join([line + f" {date}" for line in block.splitlines() if line.strip()])

    def _process_file(self, pdf_bytes: bytes, include_date=False):
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

        date = self._catch_date(full_text) if include_date else None

        spp = self.extract_sales_per_payment_block(full_text)
        spf = self.extract_sales_per_salesman_block(full_text)
        dsc = self.extract_discount_block(full_text)

        if include_date and date:
            spp = self._add_date_to_block(spp, date)
            spf = self._add_date_to_block(spf, date)

        return [spp], [dsc], [spf]

    def extract_all(self, files_bytes: list[bytes], include_date=True):
        all_spp, all_dsc, all_spf = [], [], []
        for pdf_bytes in files_bytes:
            spp, dsc, spf = self._process_file(pdf_bytes, include_date=include_date)
            all_spp.extend(spp)
            all_dsc.extend(dsc)
            all_spf.extend(spf)
        return all_spp, all_dsc, all_spf

    def parse_fechamento_caixa(self, spp_list: list[str]) -> pd.DataFrame:
        rows = []
        pat = re.compile(
            r'(?:\d+\s*-\s*)?'
            r'(.+?)\s+'
            r'[\d,]+%\s+'
            r'\d+\s+'
            r'R\$\s+([\d\.,]+)'
            r'(?:\s+(\d{2}/\d{2}/\d{4}))?'
        )
        for block in spp_list:
            if not block:
                continue
            for line in block.splitlines():
                line = line.strip()
                if not line:
                    continue
                m = pat.match(line)
                if m:
                    rows.append({
                        "FP":             m.group(1).strip(),
                        "Entrada de caixa": _parse_valor(m.group(2)),
                        "Data":           _parse_data(m.group(3)) if m.group(3) else None,
                    })
        return pd.DataFrame(rows) if rows else pd.DataFrame(
            columns=["FP", "Entrada de caixa", "Data"])

    def parse_sangrias(self, dsc_list: list[str]) -> pd.DataFrame:
        rows = []
        pat_sep = re.compile(r'^(.+?)(?:/\s*-\s*|\s+-\s+(?=\S))')
        pat_dt  = re.compile(r'(\d{2}/\d{2}/\d{4})')
        pat_val = re.compile(r'-([\d\.,]+)\s*$')

        for block in dsc_list:
            if not block:
                continue
            for line in block.splitlines():
                line = line.strip()
                if not line:
                    continue
                cat_m = pat_sep.match(line)
                dt_m  = pat_dt.search(line)
                val_m = pat_val.search(line)
                if not (cat_m and dt_m and val_m):
                    continue
                comp = line[cat_m.end():dt_m.start()].strip()
                comp = re.sub(r'\s*\d{2}:\d{2}:\d{2}.*', '', comp).strip()
                rows.append({
                    "Categoria":   cat_m.group(1).strip(),
                    "Complemento": comp,
                    "Data":        _parse_data(dt_m.group(1)),
                    "Valor":       _parse_valor(val_m.group(1)),
                })
        return pd.DataFrame(rows) if rows else pd.DataFrame(
            columns=["Categoria", "Complemento", "Data", "Valor"])

    def parse_vendedores(self, spf_list: list[str]) -> pd.DataFrame:
        rows = []
        pat = re.compile(
            r'^(.+?)\s+'
            r'(\d+)\s+'
            r'([\d,]+)%\s+'
            r'R\$\s+([\d\.,]+)'
            r'(?:\s+(\d{2}/\d{2}/\d{4}))?'
        )
        for block in spf_list:
            if not block:
                continue
            for line in block.splitlines():
                line = line.strip()
                if not line:
                    continue
                m = pat.match(line)
                if m:
                    rows.append({
                        "Nome ":      m.group(1).strip(),
                        "QTD":        int(m.group(2)),
                        "Porcentagem": _parse_valor(m.group(3)) / 100,
                        "Valor":      _parse_valor(m.group(4)),
                        "Data":       _parse_data(m.group(5)) if m.group(5) else None,
                    })
        return pd.DataFrame(rows) if rows else pd.DataFrame(
            columns=["Nome ", "QTD", "Porcentagem", "Valor", "Data"])

    def extract_all_as_dataframes(self, files_bytes: list[bytes], include_date=True) -> dict:
        spp_list, dsc_list, spf_list = self.extract_all(files_bytes, include_date)
        return {
            "caixa":      self.parse_fechamento_caixa(spp_list),
            "sangrias":   self.parse_sangrias(dsc_list),
            "vendedores": self.parse_vendedores(spf_list),
            "spp_raw":    spp_list,
            "dsc_raw":    dsc_list,
            "spf_raw":    spf_list,
        }


def _parse_valor(s: str) -> float:
    if not s:
        return 0.0
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0

def _parse_data(s: str):
    if not s:
        return None
    try:
        return datetime.strptime(s.strip(), "%d/%m/%Y")
    except ValueError:
        return None