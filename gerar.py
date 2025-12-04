from docx import Document
import pandas as pd
import os
from docx2pdf import convert

os.makedirs("temp", exist_ok=True)
os.makedirs("saida", exist_ok=True)

dados = pd.read_csv("alunos.csv")

for _, row in dados.iterrows():
    doc = Document("modelo.docx")

    # Substitui os placeholders (ex: {{{NOME}}})
    for p in doc.paragraphs:
        for key, value in row.items():
            placeholder = f"{{{{{{{key}}}}}}}"  # {{{NOME}}}
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))

    nome = str(row['NOME']).replace(" ", "_")

    docx_path = f"temp/{nome}.docx"
    pdf_path = f"saida/{nome}.pdf"

    # Salva temporário em docx
    doc.save(docx_path)

    # Converte para PDF
    convert(docx_path, pdf_path)

    print(f"✅ Certificado gerado: {row['NOME']}")

print("\nTodos os certificados em PDF foram gerados com sucesso!")
