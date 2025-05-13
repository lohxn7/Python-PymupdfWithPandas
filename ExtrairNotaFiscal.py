import fitz  # PyMuPDF

def extrair_dados_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    dados_extraidos = []

    for page in doc:
        texto = page.get_text("text")
        linhas = texto.split('\n')
        for linha in linhas:
            if linha.strip():
                dados_extraidos.append(linha.strip())
    return dados_extraidos

caminho_pdf = r"C:\Users\Kaleb CJW\Downloads\RELATORIO VENDAS CUPONS 2025 (1).pdf"

if __name__ == "__main__":
    print("Extraindo dados do PDF...")
    dados = extrair_dados_pdf(caminho_pdf)
    
    print("Todos os dados extraídos do PDF:")
    for dado in dados:
        print(dado)
