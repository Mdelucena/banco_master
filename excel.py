import os
import pandas as pd

# Caminho da pasta de downloads
pasta_downloads = os.path.expanduser("~/Downloads")

# Função para localizar o arquivo mais recente


def encontrar_arquivo_mais_recente(pasta):
    try:
        arquivo_mais_recente = max(
            [os.path.join(pasta, f) for f in os.listdir(pasta)
             if os.path.isfile(os.path.join(pasta, f))],
            key=os.path.getctime
        )
        print(f"Arquivo mais recente encontrado: {arquivo_mais_recente}")
        return arquivo_mais_recente
    except ValueError:
        print("Nenhum arquivo encontrado na pasta de downloads.")
        return None

# Verifica os headers do arquivo


def verificar_headers(arquivo, headers_esperados):
    try:
        tabela = pd.read_excel(arquivo)
        headers_do_arquivo = list(tabela.columns)

        if headers_do_arquivo == headers_esperados:
            print("Os headers estão corretos.")
        else:
            print("Os headers não estão corretos.")
            print("Esperados:", headers_esperados)
            print("Encontrados:", headers_do_arquivo)
            exit()
    except Exception as e:
        print(f"Erro ao carregar ou verificar o arquivo: {e}")
        exit()


# Configuração dos headers esperados
headers_esperados = [
    "STATUS FORMALIZAÇÃO", "CONVÊNIO", "DIA DE CORTE", "TIPO", "CPF", "NOME", "MATRÍCULA", "NSU", "PROPOSTA",
    "VALOR SAQUE", "VALOR TROCO", "VALOR PARCELA", "PRAZO", "BANCO", "AGÊNCIA", "CONTA", "DIGITADOR", "DATA SAQUE",
    "CANAL VENDA", "PENDENTE DADOS BANCÁRIOS", "CRÍTICA", "DATA DE APROVAÇÃO", "USUÁRIO APROVAÇÃO", "ID CONVENIO",
    "TIPO AUDITORIA DIGITAL", "STATUS AUDITORIA DIGITAL", "VENDA PACOTE VANTAGENS", "ACEITE PACOTE VANTAGENS",
    "VALOR PREMIO LIQUIDO", "VALOR PREMIO BRUTO", "STATUS PROPOSTA FUNCAO", "ATIVIDADE", "ORIGEM", "STATUS TED",
    "DATA EMISSÃO TED", "ÚLTIMO EVENTO TED", "MOTIVO FUNÇÃO", "TABELA FUNÇÃO", "POSSUI REPRESENTANTE LEGAL",
    "CPF REPRESENTANTE LEGAL", "NOME REPRESENTANTE LEGAL", "DIGITADOR", "COD CORRESPONDENTE", "PRODUTO",
    "TIPO_COBRANCA", "NUM PARCELAS PREMIO", "VALOR PARCELA PREMIO", "NÃO PERTURBE", "FORNECEDOR BIOMETRIA",
    "MASTER", "NÃO"
]

# Executa o Selenium para baixar o arquivo (substitua esta parte pelo seu código Selenium)
print("Executando a automação com Selenium...")

# Após baixar o arquivo com Selenium, localize o mais recente
arquivo_mais_recente = encontrar_arquivo_mais_recente(pasta_downloads)

if arquivo_mais_recente:
    # Verificar os headers do arquivo Excel
    verificar_headers(arquivo_mais_recente, headers_esperados)
else:
    print("Nenhum arquivo encontrado para verificar os headers.")
    exit()

# Finalize a automação e feche o navegador (se aplicável)
# navegador.quit()  # Inclua esta linha se estiver utilizando Selenium
print("Automação encerrada com sucesso.")
