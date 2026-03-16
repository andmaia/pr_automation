import csv
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


class Converter:
    def __init__(self):
        pass

  
    def to_excel(self, data_dict: dict[str, list], filename="output.xlsx"):
        try:
            with pd.ExcelWriter(filename, engine="openpyxl") as writer:  # filename pode ser str ou BytesIO
                for sheet_name, data in data_dict.items():

                    rows = []
                    for item in data:
                        if isinstance(item, str):
                            rows.extend(item.splitlines())
                        else:
                            rows.append(item)

                    df = pd.DataFrame(rows, columns=["Conteúdo"])
                    df.to_excel(writer, index=False, sheet_name=sheet_name)

            return filename
        except Exception as e:
            raise Exception(f"Erro ao criar o arquivo Excel: {str(e)}")
