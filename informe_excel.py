# Importamos las librerías necesarias
import xlsxwriter
import matplotlib.pyplot as plt
import pandas as pd

# Importamos esta librería de cosecha propia con funciones útiles para la realización del informe
# Es parte del análisis previo de los datos
import funciones_2016

def reporte_ejecutivo(documento, dataframe_conjunto, orders_byweek):
    # En la primera hoja del documento se incluirá el reporte ejecutivo con dos gráficas de MatplotLib
    hoja_1 = documento.add_worksheet('Executive_Report')
    # Definimos el título de la hoja
    hoja_1.merge_range('H2:U2', 'Executive Report (se adjuntan las gráficas creadas con la librería MatplotLib de Python)', documento.add_format({'align': 'center', 'bold': True, 'bg_color': '#FFB266', 'border': 1, 'font_name' : 'Times New Roman'})) 
    
    # Sacamos los datos que vamos a representar en la primera gráfica (media semanal de cada tipo de pizza)
    recuento = dataframe_conjunto['pizza_type_id'].value_counts() # Contamos cuántas veces aparece cada pizza
    tipos_pizza = dataframe_conjunto['pizza_type_id'].unique().tolist() # Sacamos los nombres de los tipos de pizzas
    cuentas_pizzas = [] # A esta lista añadiremos la suma total de cada tipo de pizza
    for i in range(len(recuento)):
        cuentas_pizzas.append(round(recuento[tipos_pizza[i]]/53, 2)) # Dividimos por el número de semanas para obtener la media
    
    # Escribimos en negrita las cabeceras de las columnas
    hoja_1.write_string('B4', 'Tipo de pizza', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#FFCC99', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_1.write_string('C4', 'Media semanal', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#FFCC99', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_1.set_column('B:C', 15) # Ampliamos las columnas para que se lean y quepan bien
    for i in range(len(recuento)): # Escribimos cada fila de la tabla, asociando a cada tipo de pizza su recuento
        hoja_1.write_string(4+i, 1, tipos_pizza[i], documento.add_format({'align': 'right', 'bg_color': '#FFE5CC', 'border': 1, 'font_name' : 'Times New Roman'}))
        hoja_1.write_number(4+i, 2, cuentas_pizzas[i], documento.add_format({'align': 'right', 'bg_color': '#FFE5CC', 'border': 1, 'font_name' : 'Times New Roman'}))

    # Pintamos un gráfico de sectores de los tipos de pizzas con MatplotLib
    plt.clf() # Limpiamos los posibles restos de gráficos anteriores
    plt.pie(cuentas_pizzas, labels = tipos_pizza) # Pintamos el gráfico de sectores
    plt.title('Media semanal de cada tipo de pizza') # Título del gráfico
    plt.savefig('pizzas_sectores.png') # Guardamos la imagen en el directorio

    # Insertamos la imagen en la hoja del documento excel
    hoja_1.insert_image('E4', 'pizzas_sectores.png')

    # Sacamos los datos que vamos a representar en la segunda gráfica (ingresos semanales)
    precios = [] # A esta lista se añadirán las ganancias de cada semana
    for semana in orders_byweek: # Redondeamos el precio final a dos decimales para trabajar con números más sencillos
        precios.append(round(semana['price'].sum(), 2))
    # Creamos un dataframe que asocia cada número de semana con su respectivo ingreso
    dataframe_ingresos = pd.DataFrame([[i, precios[i]] for i in range(len(precios))], columns = ['Semana', 'Ingreso semanal'])
    
    # Escribimos en negrita las cabeceras de las columnas
    hoja_1.write_string('Q4', 'Número de semana', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#FFCC99', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_1.write_string('R4', 'Ingreso semanal', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#FFCC99', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_1.set_column('Q:R', 18) # Ampliamos las columnas para que se lean y quepan bien
    for i in range(len(dataframe_ingresos)): # Escribimos cada fila de la tabla, asociando a cada semana su dinero obtenido
        hoja_1.write_number(4+i, 16, dataframe_ingresos.loc[i, 'Semana'], documento.add_format({'align': 'right', 'bg_color': '#FFE5CC', 'border': 1, 'font_name' : 'Times New Roman'}))
        hoja_1.write_number(4+i, 17, dataframe_ingresos.loc[i, 'Ingreso semanal'], documento.add_format({'align': 'right', 'bg_color': '#FFE5CC', 'border': 1, 'font_name' : 'Times New Roman'}))
    
    # Pintamos un gráfico de barras con el ingreso obtenido cada semana usando MatplotLib
    plt.clf() # Limpiamos los posibles restos de gráficos anteriores
    plt.bar(dataframe_ingresos['Semana'], dataframe_ingresos['Ingreso semanal'])
    plt.title('Ingreso por semana - Año 2016') # Título del gráfico
    plt.xlabel('Semana') # Título del eje X
    plt.ylabel('Ganancia ($)') # Título del eje Y
    plt.savefig('ingresos.png') # Guardamos la imagen en el directorio

    # Insertamos la imagen en la hoja del documento excel
    hoja_1.insert_image('T4', 'ingresos.png')

    return


def reporte_ingredientes(documento, media_ingredientes):
    # En la segunda hoja del documento se incluirá una gráfica con la recomendación de compra semanal de cada ingrediente
    hoja_2 = documento.add_worksheet('Ingredient_Report')
    # Definimos el título de la hoja
    hoja_2.merge_range('G2:AB2', 'Ingredient Report (se adjuntan las gráficas creadas con la librería XlsxWriter de Python)', documento.add_format({'align': 'center', 'bold': True, 'bg_color': '#6666FF', 'border': 1, 'font_name' : 'Times New Roman'})) 

    # Sacamos en forma de listas los datos que vamos a representar
    ingredientes = list(media_ingredientes.keys())
    cantidades = list(media_ingredientes.values())

    # Escribimos en negrita las cabeceras de las columnas
    hoja_2.write_string('B4', 'Ingrediente', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#9999FF', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_2.write_string('C4', 'Cantidad en kg', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#9999FF', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_2.set_column('B:B', 25) # Ampliamos las columnas para que se lean y quepan bien
    hoja_2.set_column('C:C', 15) 
    for i in range(len(ingredientes)): # Escribimos cada fila asociando a cada ingrediente su cantidad recomendada
        hoja_2.write_string(4+i, 1, ingredientes[i], documento.add_format({'align': 'right', 'bg_color': '#CCCCFF', 'border': 1, 'font_name' : 'Times New Roman'}))
        hoja_2.write_number(4+i, 2, cantidades[i], documento.add_format({'align': 'right', 'bg_color': '#CCCCFF', 'border': 1, 'font_name' : 'Times New Roman'}))

    # Pintaremos el gráfico de barras directamente en excel
    grafico_ingredientes = documento.add_chart({'type': 'column'})
    grafico_ingredientes.set_size({'x_scale': 4, 'y_scale': 2}) # Ponemos un tamaño grande para que se vean bien todos los ingredientes
    # Nombramos los ejes y el título del gráfico
    grafico_ingredientes.set_x_axis({'name': 'Ingredientes de pizzas', 'name_font': {'name': 'Arial', 'size': 12, 'bold': True}})
    grafico_ingredientes.set_y_axis({'name': 'Cantidades (kg)', 'name_font': {'name': 'Arial','size': 12, 'bold': True}})
    grafico_ingredientes.set_title({'name': 'Recomendación de compra semanal de ingredientes', 'name_font': {'name': 'Arial', 'size': 14, 'bold': True}})
    # Especificamos las celdas de los datos a representar
    grafico_ingredientes.add_series({'categories': '=Ingredient_Report!$B$5:$B$69', 'values': '=Ingredient_Report!$C$5:$C$69'})
    # Quitamos la leyenda puesto que no aporta información de valor
    grafico_ingredientes.set_legend({'none': True})

    # Añadimos a la hoja del documento excel el gráfico creado
    hoja_2.insert_chart('E4', grafico_ingredientes)

    return

def reporte_pedidos(documento, orders_byweek):
    # En la tercera hoja del documento se incluirá unan gráfica con el número de pedidos realizados cada semana del año
    hoja_3 = documento.add_worksheet('Orders_Report')
    # Definimos el título de la hoja
    hoja_3.merge_range('F2:U2', 'Orders Report (se adjuntan las gráficas creadas con la librería XlsxWriter de Python)', documento.add_format({'align': 'center', 'bold': True, 'bg_color': '#66FF66', 'border': 1, 'font_name' : 'Times New Roman'})) 

    # Sacamos en forma de listas los datos que vamos a representar
    numeros_semanas = []
    pedidos_semanas = []
    for i in range(len(orders_byweek)): # Recorremos las 53 semanas del dataset
        numeros_semanas.append(i)
        # Puesto que la lista 'orders_byweek' contiene 53 datasets (correspondientes con cada semana), el número de pedidos será su longitud
        pedidos_semanas.append(len(orders_byweek[i])) 
    
    # Escribimos en negrita las cabeceras de las columnas
    hoja_3.write_string('B4', 'Número de semana', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#99FF99', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_3.write_string('C4', 'Número de pedidos', documento.add_format({'align': 'right', 'bold': True, 'bg_color': '#99FF99', 'border': 1, 'font_name' : 'Times New Roman'}))
    hoja_3.set_column('B:C', 18) # Ampliamos las columnas para que se lean y quepan bien
    for i in range(len(numeros_semanas)): # Escribimos cada fila asociando a cada semana su respectivo número de pedidos
        hoja_3.write_number(4+i, 1, numeros_semanas[i], documento.add_format({'align': 'right', 'bg_color': '#CCFFCC', 'border': 1, 'font_name' : 'Times New Roman'}))
        hoja_3.write_number(4+i, 2, pedidos_semanas[i], documento.add_format({'align': 'right', 'bg_color': '#CCFFCC', 'border': 1, 'font_name' : 'Times New Roman'}))

    # Pintaremos el gráfico de barras directamente en excel
    grafico_pedidos = documento.add_chart({'type': 'column'})
    grafico_pedidos.set_size({'x_scale': 3, 'y_scale': 1.5}) # Establecemos un tamaño adecuado para que se vea fácilmente 
    # Nombramos los ejes y el título del gráfico
    grafico_pedidos.set_x_axis({'name': 'Semanas del año', 'name_font': {'name': 'Arial', 'size': 12, 'bold': True}})
    grafico_pedidos.set_y_axis({'name': 'Número de pedidos', 'name_font': {'name': 'Arial','size': 12, 'bold': True}})
    grafico_pedidos.set_title({'name': 'Recuento de pedidos realizados por semana en 2016', 'name_font': {'name': 'Arial', 'size': 14, 'bold': True}})
    # Especificamos las celdas de los datos a representar
    grafico_pedidos.add_series({'categories': '=Orders_Report!$B$5:$B$57', 'values': '=Orders_Report!$C$5:$C$57'})
    # Quitamos la leyenda puesto que no aporta información de valor
    grafico_pedidos.set_legend({'none': True})

    # Añadimos a la hoja del documento excel el gráfico creado
    hoja_3.insert_chart('E4', grafico_pedidos)

    return

if __name__ == '__main__':
    # Creamos la variable 'documento' como un objeto de la clase 'Workbook'
    documento = xlsxwriter.Workbook('reportes_excel.xlsx')

    # Trabajamos y amoldamos los datos con la librería de funciones 
    dataframes = funciones_2016.extract()
    dataframe_conjunto, orders_byweek = funciones_2016.fix_data(dataframes)
    media_ingredientes = funciones_2016.transform(dataframes, dataframe_conjunto, orders_byweek)

    # Creamos cada sección del informe, con su respectiva gráfica
    reporte_ejecutivo(documento, dataframe_conjunto, orders_byweek)
    reporte_ingredientes(documento, media_ingredientes)
    reporte_pedidos(documento, orders_byweek)
    
    # Cerramos el documento, guardándose en el directorio actual
    documento.close()