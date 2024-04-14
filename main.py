import requests  # Importa o módulo 'requests' para realizar solicitações HTTP
from openpyxl import Workbook  # Importa a classe 'Workbook' do módulo 'openpyxl' para criar planilhas do Excel
from requests.exceptions import HTTPError  # Importa a exceção HTTPError do módulo 'requests.exceptions'

url = 'https://openlibrary.org/search.json'  # Define a URL da API que será consultada

params = {  # Define os parâmetros da consulta, neste caso, a busca por 'the lord of the rings'
    'q': 'the lord of the rings'
}

try:  # Inicia o bloco 'try', onde colocamos o código que pode gerar exceções
    response = requests.get(url, params=params)  # Faz a solicitação HTTP para a URL especificada com os parâmetros
    response.raise_for_status()  # Verifica se houve algum erro na solicitação HTTP, caso sim, gera uma exceção
    
    data = response.json()  # Converte a resposta da solicitação em formato JSON para um dicionário Python

    # Cria uma nova planilha do Excel
    wb = Workbook()
    ws = wb.active  # Acessa a planilha ativa (por padrão, a primeira planilha criada)

    # Define o cabeçalho da planilha com os nomes das colunas
    ws.append(['author_name', 'first_publish_year', 'title', 'publisher', 'publish_date', 'publish_place', 'title_sort', 'first_sentence'])

    # Verifica se há documentos na resposta e se há pelo menos um documento
    if 'docs' in data and data['docs']:
        # Acessa os dados do primeiro documento na lista de documentos
        example = data['docs'][0]

        # Converte listas em strings, se necessário
        author_name = ', '.join(example.get('author_name', []))  # Obtém os nomes dos autores e os une em uma string separada por vírgulas
        publisher = ', '.join(example.get('publisher', []))  # Obtém os nomes dos editores e os une em uma string separada por vírgulas
        publish_date = ', '.join(example.get('publish_date', []))  # Obtém as datas de publicação e as une em uma string separada por vírgulas
        publish_place = ', '.join(example.get('publish_place', [])) if 'publish_place' in example else ''  # Obtém os locais de publicação e os une em uma string separada por vírgulas, se existirem

        # Trata o primeiro_sentence como uma string, mesmo que seja uma lista
        first_sentence = example.get('first_sentence', '')  # Obtém a primeira frase do livro
        if isinstance(first_sentence, list):  # Verifica se a primeira frase é uma lista
            first_sentence = ' '.join(first_sentence)  # Se for uma lista, junta todos os elementos em uma única string

        # Adiciona os dados do livro à planilha
        ws.append([
            author_name,
            example.get('first_publish_year', ''),  # Obtém o ano de primeira publicação do livro
            example.get('title', ''),  # Obtém o título do livro
            publisher,
            publish_date,
            publish_place,
            example.get('title_sort', ''),  # Obtém o título do livro para ordenação
            first_sentence
        ])
    else:
        print("Nenhum documento encontrado na resposta.")  # Mensagem de aviso se nenhum documento for encontrado na resposta

    # Salva a planilha em um arquivo Excel
    wb.save("dados_livros.xlsx")

except HTTPError as e:  # Captura uma exceção HTTPError, se ocorrer
    print(f"Erro HTTP: {e}")  # Imprime a mensagem de erro associada à exceção HTTPError

except Exception as e:  # Captura qualquer outra exceção que não seja HTTPError
    print(f"Ocorreu um erro: {e}")  # Imprime uma mensagem de erro genérica para outras exceções
