import requests
from openpyxl import Workbook

url = 'https://openlibrary.org/search.json'

params = {
    'q': 'the lord of the rings'
}

response = requests.get(url, params=params)
response.raise_for_status()  # Lança uma exceção para códigos de status HTTP de erro
    
data = response.json()

# Criando uma nova planilha
wb = Workbook()
ws = wb.active

# Definindo o cabeçalho da planilha
ws.append(['author_name', 'first_publish_year', 'title', 'publisher', 'publish_date', 'publish_place', 'title_sort', 'first_sentence'])

# Verificando se há documentos na resposta
if 'docs' in data and data['docs']:
    # Acessando os dados do primeiro documento na lista de documentos
    example = data['docs'][0]

    # Convertendo listas para strings
    author_name = ', '.join(example.get('author_name', []))
    publisher = ', '.join(example.get('publisher', []))
    publish_date = ', '.join(example.get('publish_date', []))
    publish_place = ', '.join(example.get('publish_place', [])) if 'publish_place' in example else ''
    
    # Tratando primeiro_sentence como uma string
    first_sentence = example.get('first_sentence', '')
    if isinstance(first_sentence, list):
        first_sentence = ' '.join(first_sentence)

    # Preenchendo a planilha com os dados
    ws.append([
        author_name,
        example.get('first_publish_year', ''),
        example.get('title', ''),
        publisher,
        publish_date,
        publish_place,
        example.get('title_sort', ''),
        first_sentence
    ])
else:
    print("Nenhum documento encontrado na resposta.")

# Salvando a planilha
wb.save("dados_livros.xlsx")
