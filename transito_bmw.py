# %%
import pandas as pd
import numpy as np
import datetime

# %%
hoy = datetime.datetime.today()

# %%
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Inicializar tkinter y ocultar la ventana principal
root = tk.Tk()
root.withdraw() 

# Abrir el explorador de archivos para seleccionar un archivo
filepath = filedialog.askopenfilename(
    title="Selecciona archivo de descarga de Autoline",
    filetypes=(("Todos los archivos", "*.*"),("Archivos de texto", "*.txt"))
)

# Imprimir la ruta del archivo seleccionado
if filepath:  # Asegurarse de que el usuario no canceló la selección
    print(f"Ruta del archivo seleccionado: {filepath}")

    # Ejemplo de cómo usar la ruta para leer un archivo CSV
    try:
        ruta = filepath
        print("DataFrame cargado exitosamente:")
    
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
else:
    print("No se seleccionó ningún archivo.")

# %%
tr = pd.read_excel(ruta)

# %%
#crear lt fact y total
condicion = [tr['Tipo.pedido']=='S',tr['Tipo.pedido']=='U',tr['Tipo.pedido']=='V']
resultado = [30,10,5]
resultado_2 = [90,25,15]

tr['LT Fact'] = np.select(condicion,resultado,0)
tr['LT Total'] = np.select(condicion,resultado_2,0)



# %%
tr['Fecha.creación'] = pd.to_datetime(tr['Fecha.creación'])

# %%
tr['Fecha O'] = np.where(
    tr['Status.línea.pedido'] == 'O',
    np.where(
        tr['Fecha.creación'] + pd.to_timedelta(tr['LT Fact'], unit='D') >= pd.to_datetime(hoy),
        tr['Fecha.creación'] + pd.to_timedelta(tr['LT Total'], unit='D'),
        pd.to_datetime(hoy) + pd.to_timedelta(tr['LT Total'] * 1.5, unit='D')
    ),
    pd.NaT  # Reemplaza 0 con un valor de fecha/tiempo nulo
)

# %%
tr['Fecha O'] = pd.to_datetime(tr['Fecha O'])

# %%
# Inicializar tkinter y ocultar la ventana principal
root = tk.Tk()
root.withdraw() 

# Abrir el explorador de archivos para seleccionar un archivo
filepath = filedialog.askopenfilename(
    title="Selecciona archivo de seguimiento de facturas",
    filetypes=(("Todos los archivos", "*.*"),("Archivos de texto", "*.txt"))
)

# Imprimir la ruta del archivo seleccionado
if filepath:  # Asegurarse de que el usuario no canceló la selección
    print(f"Ruta del archivo seleccionado: {filepath}")

    # Ejemplo de cómo usar la ruta para leer un archivo CSV
    try:
        ruta = filepath
        print("DataFrame cargado exitosamente:")
    
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
else:
    print("No se seleccionó ningún archivo.")

# %%
etas = pd.read_excel(ruta, sheet_name="EMBARQUES", header=1)
print(etas.columns)
etas = etas[etas['Status']=='PENDIENTE']

# %%
# 1. Convert all date-like columns to datetime format.
#    The `errors='coerce'` will turn any invalid values (like '0') into NaT.
etas['Fecha Arribo Bodega'] = pd.to_datetime(etas['Fecha Arribo Bodega'], errors='coerce')
etas['Retiro puerto'] = pd.to_datetime(etas['Retiro puerto'], errors='coerce')
etas['Manifiesto'] = pd.to_datetime(etas['Manifiesto'], errors='coerce')
etas['ETA inicial'] = pd.to_datetime(etas['ETA inicial'], errors='coerce')

# 2. Apply the selection logic using .notna()
#    This is much cleaner and more reliable than comparing to 0.
etas['Fecha consolidada'] = np.where(
    etas['Fecha Arribo Bodega'].notna(),
    etas['Fecha Arribo Bodega'],
    np.where(
        etas['Retiro puerto'].notna(),
        etas['Retiro puerto'],
        np.where(
            etas['Manifiesto'].notna(),
            etas['Manifiesto'],
            etas['ETA inicial']
        )
    )
)

# %%
tr['Nº.de.referencia'] = tr['Nº.de.referencia'].astype('str')

# %%


tr['Factura'] = np.where(
    tr['Status.línea.pedido'] == 'A',
    '872' + tr['Nº.de.referencia'],
    "0"
)

# %%
etas.drop_duplicates(subset='Factura BMW', inplace=True)

# %%
tr['Factura'] = tr['Factura'].astype('str')
etas['Factura BMW'] = etas['Factura BMW'].astype('str')

# %%
tr_etas = tr.merge(etas[['Factura BMW','Fecha consolidada']], left_on='Factura', right_on='Factura BMW', how='left')

# %%
tr_etas.drop(columns=['Factura BMW'], inplace=True)

# %%
tr_etas['Fecha A'] = np.where(
    tr_etas['Status.línea.pedido']=='A',
    tr_etas['Fecha consolidada'],
    pd.NaT
    
)

# %%
tr_etas['Fecha A'] = pd.to_datetime(tr_etas['Fecha A'])

# %%
tr_etas['Status'] = np.where(
    tr_etas['Status.línea.pedido'] == 'O',
    np.where(
        tr['Fecha.creación'] + pd.to_timedelta(tr['LT Fact'], unit='D') >= pd.to_datetime(hoy),"Transito","Transito Vencido"
    ),
    np.where(tr_etas['Status.línea.pedido']=="A",
             np.where(
                  tr_etas['Fecha A']  >= pd.to_datetime(hoy),"Facturado", "Facturado Vencido"
             ),"0") 
)

# %%
tr_etas['Status'].value_counts()

# %%
# Define the conditions and corresponding choices
conditions = [
    (tr_etas['Status.línea.pedido'] == 'A') & (tr_etas['Tipo.pedido'] == 'S') & (tr_etas['Status'] == 'Facturado Vencido'),
    (tr_etas['Status.línea.pedido'] == 'A') & (tr_etas['Tipo.pedido'] == 'U') & (tr_etas['Status'] == 'Facturado Vencido'),
    (tr_etas['Status.línea.pedido'] == 'A') & (tr_etas['Tipo.pedido'] == 'V') & (tr_etas['Status'] == 'Facturado Vencido'),
    (tr_etas['Status.línea.pedido'] == 'A'),
    (tr_etas['Status.línea.pedido'] == 'O')
]

# Define the choices that correspond to each condition
choices = [
    pd.to_datetime('today').normalize() + pd.Timedelta(days=45),
    pd.to_datetime('today').normalize() + pd.Timedelta(days=15),
    pd.to_datetime('today').normalize() + pd.Timedelta(days=5),
    tr_etas['Fecha A'],
    tr_etas['Fecha O']
]

# Apply the conditions using np.select()
tr_etas['Fecha Final'] = np.select(conditions, choices, default=0)

# %%
tr_etas['Fecha Final'] = pd.to_datetime(tr_etas['Fecha Final'])

# %%
tr_etas['Fecha Final'].value_counts()

# %%
condicion = [tr_etas['Status.línea.pedido']=='A', tr_etas['Status.línea.pedido']=='O']
opcion = [tr_etas['Cantidad.aconsejada'], tr_etas['Cantidad.requerida']]

tr_etas['Cantidad Final'] = np.select(condicion, opcion, 0)
tr_etas = tr_etas[['Cód.N..de.parte(empa', 'Fecha Final', 'Cantidad Final']]
tr_etas.rename(columns={'Cód.N..de.parte(empa':'Material'}, inplace=True)

# %%
# Inicializar tkinter y ocultar la ventana principal
root = tk.Tk()
root.withdraw() 

# Abrir el explorador de archivos para guardar el DataFrame
filepath = filedialog.asksaveasfilename(
    defaultextension=".xlsx",
    filetypes=( ("Archivos de Excel", "*.xlsx"),("Archivos XLSX", "*.xlsx")),
    title="Guardar DataFrame como..."
)

# Exportar el DataFrame si el usuario seleccionó una ubicación
if filepath:
    try:
        # Exportar a CSV
        if filepath.endswith('.csv'):
            tr_etas.to_csv(filepath, index=False)
            print(f"DataFrame exportado exitosamente a: {filepath}")
        
        # Exportar a Excel (requiere la librería openpyxl)
        elif filepath.endswith('.xlsx'):
            tr_etas.to_excel(filepath, index=False, engine='openpyxl')
            print(f"DataFrame exportado exitosamente a: {filepath}")
        
        else:
            print("Extensión de archivo no soportada.")

    except Exception as e:
        print(f"Error al exportar el archivo: {e}")
else:
    print("No se seleccionó una ruta de guardado.")

# %%



