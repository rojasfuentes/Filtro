import pandas as pd

# Leer el archivo de Excel
path = r'Encuesta de Satisfacción 2023 (Respuestas).xlsx'
df = pd.read_excel(path)

# Filtrar y renombrar las columnas
df = df.iloc[:, [1, 5, 6]]
df.columns = ['Correo', 'Problemas', 'Fortalezas']

# Definir las categorías y palabras clave
palabras_clave = {
    'Almacenamiento y manejo de productos': ['almacenamiento', 'excedidos', 'duerme', 'etiquetado', 'empaque', 'inventario', 'excursiones de temperatura', 'diferencias en inventario'],
    'Procesos de envío y transporte': ['embarque', 'ordenes de salida', 'disponibilidad', 'excedidos', 'embalaje correcto', 'retraso en entregas', 'bultos mal direccionados', 'daño de productos durante el transporte'],
    'Calidad de los productos y servicios': ['disponibilidad', 'área de calidad', 'entrega de productos incorrectos', 'cantidades incorrectas', 'fuera de rango de temperatura', 'corta caducidad', 'productos defectuosos'],
    'Servicio al cliente y comunicación': ['disponibilidad', 'servicio a clientes', 'caso de robo', 'falta de respuesta a correos', 'ignoran solicitudes importantes', 'dificultad para encontrar a la persona responsable', 'retraso en información'],
    'Tiempos de respuesta y seguimiento': ['tiempos excedidos de respuesta', 'lentitud en respuestas', 'planes de acción no efectivos', 'retraso en investigaciones', 'falta de seguimiento'],
    'Incidencias y no conformidades': ['seguimiento a no conformidades', 'respuesta lenta a incidencias', 'acciones correctivas ineficientes', 'investigaciones sobre inconsistencias inventariales'],
    'Colaboración entre equipos y departamentos': ['colaboración entre equipo operativo y de calidad', 'respuestas consistentes durante la investigación', 'falta de comunicación entre departamentos'],
    'Logística inversa': ['ejecución deficiente de logística inversa', 'falta de seguimiento'],
    'Cotizaciones y respuestas comerciales': ['respuestas lentas', 'ausencia de opciones para cubrir necesidades', 'falta de cumplimiento de cotizaciones'],
    'Procesos administrativos y facturación': ['fallas en calidad de archivo para facturación', 'información incompleta para validar consumos', 'demandas de tiempo para revisión']
}

# Crear una lista vacía para almacenar las categorías de cada respuesta
categorias_respuestas = []

# Asignar categorías a cada respuesta
for respuesta in df['Problemas']:
    categorias_respuesta = []
    if isinstance(respuesta, str):
        respuesta = respuesta.lower()  # Convertir a minúsculas
        for categoria, palabras in palabras_clave.items():
            palabras = [palabra.lower() for palabra in palabras]  # Convertir a minúsculas
            if any(palabra in respuesta for palabra in palabras):
                categorias_respuesta.append(categoria)
    categorias_respuestas.append(categorias_respuesta)

# Crear una lista de todas las categorías
categorias = list(palabras_clave.keys())

# Crear un diccionario para contar las ocurrencias de cada categoría
contador_categorias = {categoria: 0 for categoria in categorias}

# Contar las ocurrencias de cada categoría
for categorias in categorias_respuestas:
    for categoria in categorias:
        contador_categorias[categoria] += 1

# Crear el DataFrame con los resultados
df2 = pd.DataFrame({'Categoria': list(contador_categorias.keys()), 'Contador': list(contador_categorias.values())})

# Agregar la columna 'compañia'
df2['compañia'] = df.groupby('Problemas')['Correo'].apply(lambda x: ', '.join(x.dropna())).reset_index(drop=True)


print(df2)
