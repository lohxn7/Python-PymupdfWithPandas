import fitz  # PyMuPDF
import re
import pandas as pd


def extrair_dados_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    dados_extraidos = []

    for page in doc:
        texto = page.get_text("text")
        linhas = texto.split('\n')
        for linha in linhas:
            if linha.strip():
                dados_extraidos.append(linha.strip())
    return "\n".join(dados_extraidos)  


def processar_dados(texto):
    
    NumeroDaVenda = []
    NomeDoResponsavel = []
    NomeDoAluno = []
    TipoDeMaterial = []


    registros = re.findall(r"1\n(\d+)\n(\d{2}/\d{2}/\d{4} .+?)\n.*?OBS:\n(.+?)\s*Kit\s*(.+?)(?:\s*$|\n)", texto, re.DOTALL)
    
    for registro in registros:
 
        if len(registro) == 4:
            numero_venda, nome_responsavel, nome_aluno, tipo_material = registro
            NumeroDaVenda.append(numero_venda)

          
            nome_responsavel = nome_responsavel.split("2025")[-1].strip() 
            NomeDoResponsavel.append(nome_responsavel)

         
            NomeDoAluno.append(nome_aluno.strip())  
            TipoDeMaterial.append(tipo_material.strip())  

    return NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial


def salvar_em_excel(NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial, output_path):
    dados = {
        "Nome do Responsável": NomeDoResponsavel,
        "Nome do Aluno": NomeDoAluno,
        "Número da Venda": NumeroDaVenda,
        "Tipo de Material": TipoDeMaterial,
    }

    df = pd.DataFrame(dados)


    df.to_excel(output_path, index=False)
    print(f"Arquivo Excel salvo em: {output_path}")


caminho_pdf = r"C:\Users\Kaleb CJW\Downloads\RELATORIO VENDAS CUPONS 2025 (1).pdf"


output_excel = r"C:\Users\Kaleb CJW\Desktop\NotaFiscal.xlsx"

if __name__ == "__main__":
   
    texto_extraido = extrair_dados_pdf(caminho_pdf)
    
 
    NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial = processar_dados(texto_extraido)

    print("dados salvos.")
    print(NomeDoResponsavel)
    salvar_em_excel(NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial, output_excel), 
#salvo=asdasd