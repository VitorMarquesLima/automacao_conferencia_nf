import pdfplumber
import re
import os
import pandas as pd

# Base do projeto (dois níveis acima de /scripts/)
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# Caminhos corretos
PASTA_PDFS = os.path.join(BASE_DIR, 'entrada', 'notas_pdf')
SAIDA_ARQUIVO = os.path.join(BASE_DIR, 'saida', 'relatorios_gerados', 'dados_extraidos.xlsx')


# Expressões para extrair dados da NF de São Paulo
REGEXES = {
    'numero_nota': r'Numero da Nota\s+(\d+)',
    'data_emissao': r'Data e Hora da Emissão\s+(\d{2}/\d{2}/\d{4})',
    'cnpj_prestador': r'CPF/CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
    'cnpj_tomador': r'C\.P\.F\.\s*/\s*C\.N\.P\.J\.\s*([0-9./-]+)',
    'valor_total': r'VALOR TOTAL DOS SERVIÇOS\s*=\s*R\$[^\d]*([\d,.]+)',
    'base_calculo': r'Base de Calculo \(R\$\)[^\d]*([\d,.]+)',
    'iss': r'Valor do ISS \(R\$\)[^\d]*([\d,.]+)',
    'irrf': r'IRRF \(R\$\)[^\d]*([\d,.]+)',
    'inss': r'INSS \(R\$\)[^\d]*([\d,.]+)',
    'cofins': r'COFINS \(R\$\)[^\d]*([\d,.]+)',
    'pis': r'PIS/PASEP \(R\$\)[^\d]*([\d,.]+)',
}

def extrair_info(texto, regex):
    match = re.search(regex, texto, re.IGNORECASE)
    return match.group(1).replace('.', '').replace(',', '.') if match else None

def processar_pdf(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        texto = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    
    dados = {}
    for campo, regex in REGEXES.items():
        dados[campo] = extrair_info(texto, regex)

    dados['arquivo'] = os.path.basename(caminho_pdf)
    return dados

def main():
    registros = []
    arquivos = [f for f in os.listdir(PASTA_PDFS) if f.endswith('.pdf')]

    if not arquivos:
        print("Nenhum PDF encontrado na pasta de entrada.")
        return

    for arquivo in arquivos:
        caminho = os.path.join(PASTA_PDFS, arquivo)
        print(f"Lendo: {arquivo}")
        dados = processar_pdf(caminho)
        registros.append(dados)

    df = pd.DataFrame(registros)
    os.makedirs(os.path.dirname(SAIDA_ARQUIVO), exist_ok=True)
    df.to_excel(SAIDA_ARQUIVO, index=False)
    print(f"\n✅ Extração concluída. Arquivo salvo em: {SAIDA_ARQUIVO}")

if __name__ == '__main__':
    main()
