import docx
from collections import defaultdict

def extrair_informacoes_por_loja(caminho_arquivo):
    # Carrega o documento Word
    doc = docx.Document(caminho_arquivo)
    
    # Dicionário para armazenar informações por loja
    lojas = defaultdict(list)
    loja_atual = None
    
    # Itera sobre os parágrafos do documento
    for paragrafo in doc.paragraphs:
        texto = paragrafo.text.strip()
        if texto:
            # Verifica se o texto é um nome de loja
            if texto.startswith("*") and texto.endswith("*"):
                loja_atual = texto.strip("*")
            elif loja_atual:
                # Adiciona o texto à lista da loja atual
                lojas[loja_atual].append(texto)
    
    return lojas

def main():
    caminho_arquivo = 'dados_extraidos.docx'  # Substitua pelo caminho do seu arquivo Word
    lojas = extrair_informacoes_por_loja(caminho_arquivo)
    
    # Imprime as informações agrupadas por loja
    for loja, detalhes in lojas.items():
        print(f"\nLoja: {loja}")
        print(detalhes)

if __name__ == "__main__":
    main()