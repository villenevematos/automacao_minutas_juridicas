import pandas as pd
from docx import Document # Para o Word
from pypdf import PdfReader # Para o PDF
from num2words import num2words
# 1. Entrada de Dados (EXCEL)
df = pd.read_excel('dados_teste.xlsx')

# print(nome_cliente)
# 2. Entrada de Dados (PDF)
reader = PdfReader('acordo_teste.pdf')
# Simula a extração de um trecho (aqui, apenas pega o texto da primeira página)
trecho_pdf = reader.pages[0].extract_text()
# 3. Processamento (WORD)
documento = Document('template_minuta.docx')
# Lógica de substituição:
# Dicionário com todos os espaco_reservados e seus respectivos valores


# Aplica todas as substituições em cada parágrafo
for linha in df.iloc:
    primeira_linha = linha # Pega os dados da linha atual
    nome_cliente = primeira_linha['NOME_CLIENTE']
    numero_processo = primeira_linha['NUMERO_PROCESSO']
    valor_acordo = primeira_linha['VALOR_ACORDO']
    quantidade_parcelas = primeira_linha['QUANTIDADE_PARCELAS']
    valor_parcela = valor_acordo / quantidade_parcelas

    substituicoes = {
    '[NOME_CLIENTE]': nome_cliente,
    '[NUMERO_PROCESSO]': str(numero_processo),
    '[QUANTIDADE_PARCELAS]': str(quantidade_parcelas),
    '[VALOR_ACORDO]': str(valor_acordo),
    '[VALOR_PARCELA]': str(valor_parcela),
    '[TRECHO_PDF]': trecho_pdf,
    '[VALOR_ACORDO_POR_EXTENSO]': num2words(valor_acordo, lang='pt_BR', to='currency'),
    '[VALOR_PARCELA_POR_EXTENSO]': num2words(valor_parcela, lang='pt_BR', to='currency'),
    
    }

    for paragrafo in documento.paragraphs:
        for espaco_reservado, valor in substituicoes.items():
            if espaco_reservado in paragrafo.text:
                paragrafo.text = paragrafo.text.replace(espaco_reservado, valor)

# 4. Saída: Salva o novo documento
documento.save(f'minuta_{nome_cliente}_FINAL.docx')
print(f"Sucesso! Minuta final gerada para {nome_cliente}.")
