import qrcode
from PIL import Image
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN
from io import BytesIO
import spotipy
from spotipy.oauth2 import SpotifyClientCredentials


################################################
### DATOS A INCLUIR POR PARTE DEL USUARIO ######
################################################

# Ruta del PowerPoint final
pptx_path = "/Users/aitor/Desktop/hitster/archivos_generados/tarjetas_a_imprimir.pptx"

# Datos personales del usuario para el uso de Spotify
client_id = 'b0ce861df37447f5965272f0f620c964'
client_secret = 'be5e4971b3cf4cfe8bb0b585e77684bd'

# URL de la playlist a convertir
playlist_url = 'https://open.spotify.com/playlist/5TOMoP3KEWtRASGeFsksUY?si=fbb64fcd02d5468a'



##################
### SCRIPT ######
##################

auth_manager = SpotifyClientCredentials(client_id=client_id, client_secret=client_secret)
sp = spotipy.Spotify(auth_manager=auth_manager)


def obtener_info_playlist(url):
    playlist_id = url.split("/")[-1].split("?")[0]
    results = sp.playlist_tracks(playlist_id)
    
    total_canciones = len(results['items'])  # Total de canciones en la playlist
    canciones = []
    posicion = 1  # Iniciar contador para la posición de la canción
    
    for item in results['items']:
        track = item['track']
        titulo = track['name']
        artistas = ', '.join([artist['name'] for artist in track['artists']])
        url_cancion = track['external_urls']['spotify']
        album = track['album']
        año_publicacion = album['release_date'][:4]  # Asume formato YYYY-MM-DD
        
        # Calcular la proporción de la posición respecto al total de canciones
        proporcion = f'{posicion} / {total_canciones}'
        
        canciones.append({
            'URL': url_cancion,
            'Título': titulo,
            'Artistas': artistas,
            'Año de publicación': año_publicacion,
            'Posición': posicion,  # Posición actual de la canción
            'Proporción': proporcion  # Formato 'POSICION / TOTAL CANCIONES' como string
        })
        posicion += 1  # Actualizar posición para la próxima canción
    
    return canciones


# Usa la función con una URL de una lista de reproducción de Spotify
info_canciones = obtener_info_playlist(playlist_url)

df = pd.DataFrame(info_canciones)
df.rename(columns={'Artistas': 'Artista', 'Año de publicación': 'Año', 'Posición': 'Pos'}, inplace=True)
print(df)

# Crear una presentación de PowerPoint
prs = Presentation()

# Establecer el tamaño de la diapositiva (19 cm x 19 cm)
prs.slide_width = Cm(19)
prs.slide_height = Cm(19)

# Tamaño del QR y del cuadro de texto
qr_size = Cm(12)  # Tamaño del código QR
text_box_width = Cm(17)  # Ancho del cuadro de texto

# Espacio después de cada párrafo
space_after = Pt(12)

# Calcular posición centrada de la caja de texto
text_box_left = (prs.slide_width - text_box_width) / 2
text_box_top = (prs.slide_height - qr_size) / 2  # Asumiendo que queremos centrar respecto al QR

for index, row in df.iterrows():
    # Datos de cada fila
    url = row['URL']
    title = row['Título']
    artist = row['Artista']
    year = row['Año']

    # Generar el código QR
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    bio = BytesIO()
    img.save(bio)
    bio.seek(0)

    # Añadir diapositiva para el QR, centrado
    slide_qr = prs.slides.add_slide(prs.slide_layouts[5])
    qr_left = (prs.slide_width - qr_size) / 2
    qr_top = (prs.slide_height - qr_size) / 4  # Posicionamos el QR en la parte superior
    slide_qr.shapes.add_picture(bio, qr_left, qr_top, width=qr_size, height=qr_size)

    # Añadir diapositiva para los datos, centrado
    slide_info = prs.slides.add_slide(prs.slide_layouts[5])
    textbox = slide_info.shapes.add_textbox(text_box_left, text_box_top, text_box_width, qr_size)
    tf = textbox.text_frame
    tf.word_wrap = True

    # Configurar el título
    p = tf.add_paragraph()
    p.text = title
    p.space_after = space_after
    p.alignment = PP_ALIGN.CENTER

    # Agregar salto de línea adicional
    p = tf.add_paragraph()
    p.text = ""
    p.space_after = space_after

    # Configurar el artista
    p = tf.add_paragraph()
    p.text = artist
    p.font.bold = True
    p.space_after = space_after
    p.alignment = PP_ALIGN.CENTER

    # Agregar salto de línea adicional
    p = tf.add_paragraph()
    p.text = ""
    p.space_after = space_after

    # Configurar el año con el tamaño deseado
    p = tf.add_paragraph()
    p.text = str(year)
    p.font.size = Pt(80)  # Tamaño de fuente para el año
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER


# Export DataFrame to an Excel file
df.to_excel(pptx_path.replace('.pptx', '.xlsx'), index=False)
# Guardar la presentación
prs.save(pptx_path)
print(f"PowerPoint guardado en {pptx_path} con éxito güey!")
