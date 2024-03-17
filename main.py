import requests
from openpyxl import Workbook

url = 'https://openlibrary.org/search.json'

params = {
    'q': 'the lord of the rings'
}

try:
    response = requests.get(url, params=params)
    response.raise_for_status()  # Lança uma exceção para códigos de status HTTP de erro
    
    data = response.json()
    docs = data['docs']
    
    # Criar uma nova planilha
    wb = Workbook()
    ws = wb.active
    
    # Definir cabeçalhos para as colunas
    ws.append(["Author Name", "Publish Year", "Publisher"])
    
    for doc in docs:
        author_name = doc.get('author_name', '')
        publishers = doc.get('publisher', [])
        publish_year = doc.get('publish_year', '')

        for publisher in publishers:
            if len(publisher) < 5:
                # Converter lista de autores para uma string separada por vírgulas
                authors_str = ', '.join(author_name) if isinstance(author_name, list) else author_name
                
                # Converter lista de anos de publicação para uma string separada por vírgulas
                publish_year_str = ', '.join(map(str, publish_year)) if isinstance(publish_year, list) else publish_year
                
                # Adicionar os dados à planilha
                ws.append([authors_str, publish_year_str, publisher])
        
        break  # Interrompe o loop após a primeira iteração
    
    # Salvar a planilha
    wb.save("dados_publicacao.xlsx")
    
    print("Planilha criada com sucesso.")
    
except requests.exceptions.RequestException as e:
    print(f"Erro ao fazer solicitação: {e}")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
