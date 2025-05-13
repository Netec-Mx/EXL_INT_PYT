
# Práctica 3. Importación y exportación de datos entre Excel, CSV, JSON y SQLite

## Objetivo de la práctica:

Al finalizar esta práctica, será capaz de:
- Importar datos desde archivos CSV y JSON, manipularlos con `pandas`, y exportarlos tanto a Excel como a una base de datos SQLite. Todo se realizará con un script en Python usando VS Code en Windows.

## Objetivo visual

![Objetivo Visual](../images/cap3_objetivo.png)

## Duración aproximada:
- 45 minutos.

---

## Instrucciones

### Tarea 1. **Configurar el entorno de trabajo**

Paso 1. Crear una carpeta dentro de VS Code llamada `capitulo3_datos`.

![Tarea 1](../images/cap3_1.png)

Paso 2. Crear un archivo Python en esa carpeta, para ello hacer clic derecho sobre la carpeta → **Nuevo archivo** → nómbralo `import_export_datos.py`.

![Tarea 2](../images/cap3_2.png)

Paso 3. Instalar las librerías necesarias. Abrir la terminal en VS Code con `Ctrl + ñ` y escribir este comando:

```bash
pip install pandas xlwings openpyxl
```

![Tarea 3](../images/cap3_3.png)

---

### Tarea 2. **Crear los archivos de entrada (CSV y JSON)**

Paso 4. Crear el archivo `productos.csv`
1. Hacer clic derecho sobre la carpeta → **Nuevo archivo** → nómbralo `productos.csv`.

![Tarea 2](../images/cap3_4.png)

2. Colocar este contenido en el archivo csv y guardarlo:

```csv
codigo,producto,categoria,precio
P001,Lápiz,Papelería,0.5
P002,Cuaderno,Papelería,2.0
P003,Taza,Hogar,3.75
```

![Tarea 2](../images/cap3_5.png)

Paso 5. Crear el archivo `clientes.json`.
1. Hacer clic derecho → **Nuevo archivo** → nómbralo `clientes.json`.

![Tarea 2](../images/cap3_6.png)

2. Colocar este contenido en el archivo json y guardarlo:

```json
[
  {"id": 1, "nombre": "Ana", "ciudad": "Bogotá"},
  {"id": 2, "nombre": "Luis", "ciudad": "Medellín"},
  {"id": 3, "nombre": "Sofía", "ciudad": "Cali"}
]
```

![Tarea 2](../images/cap3_7.png)

---

### Tarea 3. **Leer los archivos desde Python usando pandas**

Paso 6. Abrir el archivo `import_export_datos.py` y escribir:

```python
import pandas as pd

# Leer datos desde CSV
productos = pd.read_csv('capitulo3_datos\\productos.csv')

# Leer datos desde JSON
clientes = pd.read_json('capitulo3_datos\\clientes.json')

# Imprimir resultados en consola
print("Productos:")
print(productos)

print("\nClientes:")
print(clientes)
```

Paso 7. Ejecutar el archivo para visualizar los datos cargados correctamente en la terminal.

![Tarea 3](../images/cap3_8.png)

---

### Tarea 4. **Exportar los datos a Excel usando xlwings**

Paso 8. Añadir el siguiente bloque de código:

```python
import xlwings as xw

# Crear un libro nuevo en Excel
wb = xw.Book()  # Abre Excel con un libro nuevo

# Agregar productos a la primera hoja
ws1 = wb.sheets[0]
ws1.name = 'Productos'
ws1.range('A1').value = productos

# Crear otra hoja para clientes
ws2 = wb.sheets.add('Clientes')
ws2.range('A1').value = clientes

# Guardar el archivo
wb.save('datos_exportados.xlsx')
```

Paso 9. Ejecutar el script nuevamente.
- Excel se abrirá automáticamente.
- Se creará un archivo llamado `datos_exportados.xlsx` en tu carpeta.
- Verificar que haya dos hojas: “Productos” y “Clientes”.

![Tarea 4](../images/cap3_9.png)

---

### Tarea 5. **Guardar los datos en una base de datos SQLite**

Paso 10. Cerrar los archivos de Excel abiertos y agregar este bloque al final del archivo:

```python
import sqlite3

# Crear la base de datos (si no existe)
conn = sqlite3.connect('datos.db')

# Exportar los datos a SQLite
productos.to_sql('productos', conn, if_exists='replace', index=False)
clientes.to_sql('clientes', conn, if_exists='replace', index=False)

# Verificar una consulta simple
consulta = pd.read_sql_query('SELECT * FROM productos', conn)
print("\nConsulta desde la base de datos:")
print(consulta)

conn.close()
```

Paso 11. Ejecutar el script.
- Se creará un archivo `datos.db` en tu carpeta.
- La terminal mostrará los datos consultados desde la base de datos.

![Tarea 5](../images/cap3_10.png)

---

### Tarea 6. **Verificar los archivos generados**

Paso 12. Asegurarte de tener en tu carpeta:

- `productos.csv`  
- `clientes.json`  
- `import_export_datos.py`  
- `datos_exportados.xlsx` (archivo Excel generado)  
- `datos.db` (base de datos SQLite)

Paso 13. Abrir el archivo Excel.
- Verificar que las dos hojas estén correctamente formateadas.

![Tarea 6](../images/cap3_11.png)
![Tarea 6](../images/cap3_12.png)

Paso 14. (Opcional) Abrir la base de datos con un visor SQLite (como [DB Browser for SQLite](https://sqlitebrowser.org/)) si quieres ver las tablas.

![Tarea 6](../images/cap3_13.png)
![Tarea 6](../images/cap3_14.png)

---

### Resultado esperado

![Tarea 6](../images/cap3_11.png)
![Tarea 6](../images/cap3_12.png)
![Tarea 6](../images/cap3_13.png)
![Tarea 6](../images/cap3_14.png)
