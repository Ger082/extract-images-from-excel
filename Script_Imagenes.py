import os
import time
import xlwings as xw
from PIL import ImageGrab

# ------------- CONFIGURACIÓN -------------
EXCEL_FILE    = r"C:\Users\GerMachine\Desktop\pruebaInv\Programa.xlsm"
SHEET_NAME    = "Inventario"
OUTPUT_FOLDER = r"C:\Users\GerMachine\Desktop\pruebaInv\Imagenes"


def export_images():
    # Crear carpeta de salida si no existe
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Iniciar Excel en segundo plano
    app = xw.App(visible=False)
    wb  = app.books.open(EXCEL_FILE)
    sht = wb.sheets[SHEET_NAME]
    sht_api = sht.api

    # Determinar última fila de datos (columna A)
    last_row = sht.range("A1").expand().last_cell.row

    print(f"Procesando filas 2 a {last_row} en «{SHEET_NAME}»...")

    for row in range(2, last_row + 1):
        # Si ya hay ruta en F, saltar fila
        if sht.cells(row, 6).value not in (None, ""):
            continue

        # Buscar shapes en la celda G de esta fila
        shapes = [shp for shp in sht_api.Shapes
                  if shp.TopLeftCell.Row == row and shp.TopLeftCell.Column == 7 and shp.Type == 13]

        # Si no hay imagen pegada en G, omitir fila
        if not shapes:
            continue

        # Tomar la primera imagen encontrada
        shp = shapes[0]

        # Copiar al portapapeles y extraer con PIL
        shp.CopyPicture(Appearance=1, Format=2)
        time.sleep(0.2)
        img = ImageGrab.grabclipboard()
        if img is None:
            print(f"Fila {row}: imagen en portapapeles no válida, omitiendo.")
            continue

        # Construir nombre de archivo: Pieza_Ubic.jpg
        pieza = sht.cells(row, 1).value or ""
        ubic  = sht.cells(row, 2).value or ""
        nombre = f"{pieza}_{ubic}".strip().replace(" ", "_")
        for ch in r'\/:*?"<>|':
            nombre = nombre.replace(ch, "_")
        jpg_path = os.path.join(OUTPUT_FOLDER, nombre + ".jpg")

       
        img.convert("RGB").save(jpg_path, "JPEG")
        # Escribir ruta en columna F
        sht.cells(row, 6).value = jpg_path
        print(f"Fila {row}: guardada → {jpg_path}")

    
    wb.save()
    wb.close()
    app.quit()

if __name__ == "__main__":
    export_images()
