from urllib.parse import urljoin
import requests
import pandas as pd
import json
from apify import Actor
from bs4 import BeautifulSoup
import time
from io import BytesIO
import xlrd


async def main():
    async with Actor:
        # Read the Actor input
        actor_input = await Actor.get_input() or {}
        
        # Read form data from actor input
        Actor.log.info('Read and prepare Headers ...')
        form_data = {
            'busqueda_proyectos[autor]': actor_input.get('autor', ''),
            'busqueda_proyectos[palabra]': actor_input.get('palabra_clave', ''),
            'busqueda_proyectos[opcion]': actor_input.get('opcion', ''),
            'busqueda_proyectos[palabra2]': actor_input.get('segunda_palabra_clave', ''),
            'busqueda_proyectos[comision]': actor_input.get('comisiones', ''),
            'busqueda_proyectos[tipoDocumento]': actor_input.get('tipo_documento', ''),
            'busqueda_proyectos[expedienteLugar]': actor_input.get('origen_expediente', ''),
            'busqueda_proyectos[expedienteNumeroPre]': actor_input.get('numero', ''),
            'busqueda_proyectos[expedienteNumeroPos]': actor_input.get('año', ''),
            'busqueda_proyectos[expedienteTipo]': actor_input.get('tipo_expediente', '')
        }
        
        Actor.log.info(f'Headers -> {form_data} ...')

        # Create a requests session
        s = requests.Session()

        # Visit the initial URL to get cookies
        initial_url = 'https://www.senado.gob.ar/parlamentario/parlamentaria/'
        response = s.get(initial_url)

        errors = await validate_form_data(form_data, response.text)
    
        if errors:
            await Actor.push_data({
                "Errores": errors
            })
            return

        Actor.log.info('Sesion Started and cookies ready ...')

        # Perform POST request to the second URL
        post_url = 'https://www.senado.gob.ar/parlamentario/parlamentaria/avanzada'
        Actor.log.info('Sending Form ...')
        post_response = s.post(post_url, data=form_data)
        
        soup = BeautifulSoup(post_response.text, 'html.parser')
        element = soup.select_one('div.alert.alert-info strong')

        if element and element.string == ' Sin Resultados':
            await Actor.push_data({
                "Error": "No se encontraron resultados"
            })
            return
        
        
        if post_response.status_code != 200:
            Actor.log.error("Error en la solicitud POST")
            return
        
        Actor.log.info('Downloading data ...')

        # Perform GET request to download the Excel file
        download_url = 'https://www.senado.gob.ar/micrositios/DatosAbiertosExpedientes/BusquedaAvanzada/XLS'
        headers = {
            'Accept-Encoding': 'gzip, deflate, br',
            'Referer': 'https://www.senado.gob.ar/parlamentario/parlamentaria/avanzada'
        }
        get_response = s.get(download_url, headers=headers)
        
        if get_response.status_code != 200:
            Actor.log.error("Error en la solicitud GET")
            return
        if get_response.status_code == 200:
            print("200 en la solicitud GET")
                
        # Esperar para asegurarse de que el archivo se descargue completamente
        time.sleep(20)

        # Save the Excel file
        with open('data.xls', 'wb') as f:
            f.write(get_response.content)

        # Read the Excel file directly from the response content
        try:
            workbook = xlrd.open_workbook_xls('data.xls', ignore_workbook_corruption=True)                       
            data = pd.read_excel(workbook)
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            return

        # Convert the DataFrame to JSON
        json_data = data.to_json(orient='records')

        # Push the JSON data into the default dataset
        await Actor.push_data(json.loads(json_data))

def get_select_options(html, select_name):
    soup = BeautifulSoup(html, 'html.parser')
    select_element = soup.find('select', {'name': select_name})

    if not select_element:
        print(f"No se encontró el elemento <select> con el nombre {select_name}")
        return None

    options = select_element.find_all('option')
    option_values = [option['value'] for option in options if option.has_attr('value')]

    return option_values


async def validate_form_data(form_data, html):
    errors = []
    fields_to_validate = [
        'busqueda_proyectos[autor]', 'busqueda_proyectos[comision]',
        'busqueda_proyectos[tipoDocumento]', 'busqueda_proyectos[expedienteLugar]',
        'busqueda_proyectos[expedienteTipo]'
    ]
    
    for field in fields_to_validate:
        valid_options = get_select_options(html, field)
        if form_data[field] not in valid_options:
            errors.append(f"El valor para '{field}' no es válido")
            
    return errors