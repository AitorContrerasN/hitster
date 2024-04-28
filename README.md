# Hitster Generator
Es un script que crea tarjetas para jugar a Hitster. Crea tarjetas en tamaño grande (19cm x 19cm) para jugar a una versión "en grande" (de pie, por ejemplo).

# Funcionamiento
El script toma la URL de una lista de reproducción de Spotify. A partir de esa lista de reproducción, se hace una consulta a la API de Spotify: para cada canción, se obtiene la URL de la misma en Spotify, el título de la misma, el artista, el año de publicación, y también la posición de la canción en la lista. Con todos estos datos se crea un DataFrame de Pandas. 

A partir de los datos del DataFrame, el script genera, para cada canción, un código QR que lleva a la URL de la misma. 

Con ambos datos (los datos de la canción y el código QR de la misma), el script crea una presentación de PPT cuadrada (19cm x 19cm) y pone, en una slide, el código QR, y en la siguiente slide los datos de la canción. Así con todas las canciones. Finalmente se exporta este PPT, que se puede exportar a PDF, imprimir a doble cara, y así obtener las tarjetas para el juego. 

# Modos
El script tiene tres modos: 

- Modo Playlist: el explicado arriba. 
- Modo Excel: en vez de tomar las canciones de una URL de una lista de Spotify, toma los datos de un Excel con URLs de canciones sueltas. 
- Modo Excel sin año: igual que el anterior, pero en lugar de un Excel con URLs, el Excel tiene URLs de canciones y el año (manualmente) de las mismas. 