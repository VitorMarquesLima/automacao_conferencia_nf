import pdfplumber
import re
import os
import pandas as pd
from datetime import datetime
import logging
from pathlib import Path

# ==============================================
# 1. CONFIGURAÇÃO INICIAL
# ==============================================
def configurar_ambiente():
    """Configura paths e logging"""
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    PROJECT_DIR = os.path.abspath(os.path.join(BASE_DIR, '..'))
    
    # Configuração de logs
    LOG_DIR = os.path.join(PROJECT_DIR, 'logs')
    LOG_FILE = os.path.join(LOG_DIR, 'extrator.log')
    os.makedirs(LOG_DIR, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        filename=LOG_FILE
    )

    # Caminhos de entrada e saída
    config = {
        'PASTA_PDFS': os.path.join(PROJECT_DIR, 'entrada', 'notas_pdf'),
        'PASTA_RELATORIOS': os.path.join(PROJECT_DIR, 'entrada', 'relatorios'),
        'SAIDA_DIR': os.path.join(PROJECT_DIR, 'saida', 'relatorios_gerados'),
        'LOG_FILE': LOG_FILE
    }
    
    # Garantir que todos os diretórios existam
    os.makedirs(config['PASTA_PDFS'], exist_ok=True)
    os.makedirs(config['PASTA_RELATORIOS'], exist_ok=True)
    os.makedirs(config['SAIDA_DIR'], exist_ok=True)
    
    return config

# ==============================================
# 2. DEFINIÇÃO DOS PADRÕES DE EXTRAÇÃO
# ==============================================
def carregar_padroes():
    """Retorna os padrões de regex para extração"""
    return {
        'numero_nota': [
            r'NFS-e\s*n°\s*(\d+\.\d+)',
            r'Nota Fiscal Eletrônica\s*N°\s*(\d+\.\d+)'
        ],
        'data_emissao': [
            r'Data e Hora de Emissão\s*(\d{2}/\d{2}/\d{4}\s*\d{2}:\d{2}:\d{2})',
            r'emitida em\s*(\d{2}/\d{2}/\d{4})'
        ],
        'cnpj_prestador': [
            r'PRESTADOR DE SERVIÇOS.*?CPF/CNPJ\s*([\d./-]+)'
        ],
        'nome_prestador': [
            r'Nome/Razão Social:\s*(.*?)\s*Endereço'
        ],
        'cnpj_tomador': [
            r'TOMADOR DE SERVIÇOS.*?C\.P\.F\./C\.N\.P\.J\.\s*([\d./-]+)'
        ],
        'nome_tomador': [
            r'TOMADOR DE SERVIÇOS.*?Nome/Razão Social:\s*(.*?)\s*C\.P\.F\./C\.N\.P\.J\.'
        ],
        'valor_total': [
            r'VALOR TOTAL DOS SERVIÇOS\s*=\s*R\$\s*([\d.,]+)',
            r'Valor Total da Nota\s*R\$\s*([\d.,]+)'
        ],
        'base_calculo': [
            r'Base de Calculo\s*\(R\$\)\s*([\d.,]+)'
        ],
        'iss': [
            r'Valor do ISS\s*\(R\$\)\s*([\d.,]+)'
        ],
        'aliquota_iss': [
            r'Alíquota\s*\(%\)\s*([\d.,]+)'
        ],
        'municipio_prestacao': [
            r'Município da Prestação de Serviços\s*(\d+\s*-\s*[^/]+/[A-Z]{2})'
        ]
    }

# ==============================================
# 3. FUNÇÕES DE PROCESSAMENTO
# ==============================================
class ProcessadorPDF:
    def __init__(self, padroes):
        self.padroes = padroes
    
    def limpar_valor(self, valor):
        """Converte valores para float, tratando casos especiais"""
        if valor in (None, '', '0,00'):
            return 0.0
        try:
            return float(valor.replace('.', '').replace(',', '.'))
        except ValueError:
            logging.warning(f"Valor não numérico: {valor}")
            return 0.0
    
    def extrair_info(self, texto, padroes):
        """Tenta extrair informação usando múltiplos padrões"""
        for padrao in (padroes if isinstance(padroes, list) else [padroes]):
            try:
                match = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE | re.DOTALL)
                if match:
                    return match.group(1).strip()
            except Exception as e:
                logging.error(f"Erro no padrão {padrao}: {str(e)}")
        return None
    
    def processar_pdf(self, caminho_pdf):
        """Processa um arquivo PDF e retorna dict com dados"""
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto = "\n".join(
                    page.extract_text(x_tolerance=2, y_tolerance=2)
                    for page in pdf.pages
                    if page.extract_text()
                )
            
            if not texto:
                logging.warning(f"Arquivo sem texto extraível: {caminho_pdf}")
                return None
            
            # Debug: visualizar texto extraído (útil para ajustar regex)
            logging.debug(f"Texto extraído de {os.path.basename(caminho_pdf)}:\n{texto[:500]}...")
            
            dados = {'arquivo': os.path.basename(caminho_pdf)}
            
            for campo, padroes in self.padroes.items():
                valor = self.extrair_info(texto, padroes)
                
                # Conversão especial por tipo de campo
                if any(key in campo for key in ['valor', 'iss', 'base', 'aliquota']):
                    dados[campo] = self.limpar_valor(valor)
                else:
                    dados[campo] = valor if valor else ""
            
            return dados
        
        except Exception as e:
            logging.error(f"Erro ao processar {caminho_pdf}: {str(e)}")
            return None

# ==============================================
# 4. GERADOR DE RELATÓRIOS
# ==============================================
class GeradorRelatorio:
    def __init__(self, config):
        self.config = config
        self.saida_arquivo = os.path.join(
            config['SAIDA_DIR'], 
            f'dados_extraidos_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
        )
    
    def gerar_excel(self, registros):
        """Gera relatório final em Excel"""
        if not registros:
            logging.warning("Nenhum dado válido para gerar relatório")
            return False
        
        try:
            colunas = [
                'arquivo', 'numero_nota', 'data_emissao', 
                'cnpj_prestador', 'nome_prestador',
                'cnpj_tomador', 'nome_tomador',
                'valor_total', 'base_calculo', 'iss', 'aliquota_iss',
                'municipio_prestacao'
            ]
            
            # Garante todas colunas mesmo que vazias
            df = pd.DataFrame(registros)
            for col in colunas:
                if col not in df.columns:
                    df[col] = None
            
            # Ordena colunas e preenche valores nulos
            df = df[colunas].fillna({
                'valor_total': 0, 'base_calculo': 0, 'iss': 0, 'aliquota_iss': 0
            })
            
            df.to_excel(self.saida_arquivo, index=False)
            logging.info(f"Relatório gerado: {self.saida_arquivo}")
            return True
        
        except Exception as e:
            logging.error(f"Erro ao gerar Excel: {str(e)}")
            return False

# ==============================================
# 5. CONTROLE PRINCIPAL
# ==============================================
def main():
    # 1. Configuração inicial
    config = configurar_ambiente()
    
    # 2. Carregar padrões de extração
    padroes = carregar_padroes()
    processador = ProcessadorPDF(padroes)
    gerador = GeradorRelatorio(config)
    
    # 3. Processar arquivos
    registros = []
    try:
        arquivos = [f for f in os.listdir(config['PASTA_PDFS']) if f.lower().endswith('.pdf')]
        
        if not arquivos:
            print("⚠️ Nenhum PDF encontrado na pasta de entrada.")
            return

        print(f"🔍 Processando {len(arquivos)} arquivos...")
        
        for arquivo in arquivos:
            caminho = os.path.join(config['PASTA_PDFS'], arquivo)
            print(f"Processando: {arquivo}")
            
            dados = processador.processar_pdf(caminho)
            if dados:
                registros.append(dados)
        
        # 4. Gerar relatório
        if gerador.gerar_excel(registros):
            print(f"✅ Relatório gerado com sucesso: {gerador.saida_arquivo}")
            print("\nResumo dos dados:")
            df = pd.DataFrame(registros)
            print(df.head().to_string(index=False))
        else:
            print("⚠️ Não foi possível gerar o relatório.")
    
    except Exception as e:
        print(f"❌ Erro fatal: {str(e)}")
        logging.error(f"Erro no processamento principal: {str(e)}")

if __name__ == '__main__':
    main()