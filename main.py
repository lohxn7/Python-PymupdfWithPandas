import fitz  # PyMuPDF
import re
import pandas as pd

# Função para extrair texto do PDF
def extrair_dados_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    dados_extraidos = []

    for page in doc:
        texto = page.get_text("text")
        linhas = texto.split('\n')
        for linha in linhas:
            if linha.strip():
                dados_extraidos.append(linha.strip())
    return "\n".join(dados_extraidos)  # Junta tudo como uma string para análise mais simples

# Função para processar os dados extraídos
def processar_dados(texto):
    # Listas para armazenar os dados
    NumeroDaVenda = []
    NomeDoResponsavel = []
    NomeDoAluno = []
    TipoDeMaterial = []

    # Regex para capturar blocos de dados a partir de "1"
    # Ajustando a regex para melhorar a captura dos nomes e separar corretamente as informações
    registros = re.findall(r"1\n(\d+)\n(\d{2}/\d{2}/\d{4} .+?)\n.*?OBS:\n(.+?)\s*Kit\s*(.+?)(?:\s*$|\n)", texto, re.DOTALL)
    
    for registro in registros:
        # Verificando se o registro possui os 4 valores esperados
        if len(registro) == 4:
            numero_venda, nome_responsavel, nome_aluno, tipo_material = registro
            NumeroDaVenda.append(numero_venda)

            # Corrigindo os nomes dos responsáveis para remover espaços extras e separar da data
            nome_responsavel = nome_responsavel.split("2025")[-1].strip()  # Remove a data e espaços extras
            NomeDoResponsavel.append(nome_responsavel)

            # Ajustando a extração para evitar problemas com espaços
            NomeDoAluno.append(nome_aluno.strip())  # Garantir que não haja espaços extras
            TipoDeMaterial.append(tipo_material.strip())  # Garantir que não haja espaços extras

    return NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial

# Função para salvar os dados no Excel
def salvar_em_excel(NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial, output_path):
    dados = {
        "Nome do Responsável": NomeDoResponsavel,
        "Nome do Aluno": NomeDoAluno,
        "Número da Venda": NumeroDaVenda,
        "Tipo de Material": TipoDeMaterial,
    }

    # Criar um DataFrame com os dados
    df = pd.DataFrame(dados)

    # Salvar em um arquivo Excel
    df.to_excel(output_path, index=False)
    print(f"Arquivo Excel salvo em: {output_path}")

# Caminho para o PDF
caminho_pdf = r"C:\Users\Kaleb CJW\Downloads\RELATORIO VENDAS CUPONS 2025 (1).pdf"

# Caminho de saída do Excel
output_excel = r"C:\Users\Kaleb CJW\Desktop\NotaFiscal.xlsx"

if __name__ == "__main__":
    print("Extraindo dados do PDF...")
    texto_extraido = extrair_dados_pdf(caminho_pdf)
    
    print("Processando dados...")
    NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial = processar_dados(texto_extraido)

    print("Salvando dados no Excel...")
    salvar_em_excel(NumeroDaVenda, NomeDoResponsavel, NomeDoAluno, TipoDeMaterial, output_excel)
