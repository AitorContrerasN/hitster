import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
import pandas as pd

def obtener_info_canciones_desde_excel(ruta_archivo):
    # Cargar datos del archivo Excel
    df_urls = pd.read_excel(ruta_archivo)
    
    # Configurar credenciales para la API de Spotify
    client_id = 'tu_client_id'
    client_secret = 'tu_client_secret'
    client_credentials_manager = SpotifyClientCredentials(client_id=client_id, client_secret=client_secret)
    sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)
    
    # Lista para almacenar los datos de las canciones
    canciones_info = []
    
    # Número total de canciones
    total_canciones = len(df_urls)
    
    # Iterar sobre las URLs en el DataFrame, utilizando enumerate para obtener la posición
    for posicion, url in enumerate(df_urls['URL'], start=1):
        track_id = url.split("/")[-1]
        try:
            # Obtener datos de la canción
            track = sp.track(track_id)
            track_name = track['name']
            track_url = track['external_urls']['spotify']
            release_year = track['album']['release_date'][:4]
            proporcion = f'{posicion} / {total_canciones}'
            
            # Añadir la información a la lista
            canciones_info.append({
                'URL Canción': track_url,
                'Nombre Canción': track_name,
                'Año de Lanzamiento': release_year,
                'Posición': posicion,
                'Proporción': proporcion
            })
        except Exception as e:
            print(f"No se pudo obtener información para la URL: {url}. Error: {str(e)}")
    
    # Crear un DataFrame con la información recolectada
    df_resultado = pd.DataFrame(canciones_info)
    
    return df_resultado

# Uso de la función
ruta_excel = '/Users/aitor/Desktop/hitster/modo_excel/modo_excel.xlsx'
df_canciones = obtener_info_canciones_desde_excel(ruta_excel)
print(df_canciones)
