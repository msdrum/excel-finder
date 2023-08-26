import pandas as pd

# Escolhendo o documento Excel.
documento = input('Digite o nome do documento em que deseja pesquisar: ')

# Ler a planilha, através do caminho onde ela se encontra.
excel = pd.read_excel(f'{documento}.xlsx')

# Função para aplicar minúsculas a uma célula
def to_lowercase(valor):
    if isinstance(valor, str):
        return valor.lower()
    else:
        pass
    return valor

# Aplicar a função a todas as células da planilha
excel_minusc = excel.applymap(to_lowercase)

# Criar o laço da pesquisa (while)
resposta = ''

while resposta != 'n':

    # termo que deseja buscar na planilha.
    palavra = input('Digite a palavra que você deseja pesquisar: ').lower()

    # Verificar se a palavra existe na planilha
    if excel_minusc.applymap(lambda celula: palavra in str(celula)).any().any():
        # Encontrando a palavra desejada.
        def buscar_item(celula):
            return palavra in str(celula)

        busca = excel_minusc[excel_minusc.applymap(buscar_item).any(axis=1)]
        print(busca)
    else:
        print('Nenhuma ocorrência encontrada.')
        quit()
        
    

    # Perguntar se deseja refazer a pesquisa ou imprimir o resultado.
    resposta = input('Deseja refazer a pesquisa?[S/N] ').lower()

# Criando uma nova planilha com os itens desejados.
def sair(pergunta):
    if pergunta == 's':
        busca.to_excel(f'{palavra}_selected.xlsx')
    quit()

sair(pergunta=input('Deseja gravar os resultados da última pesquisa em outra planilha? [S/N] ').lower())




