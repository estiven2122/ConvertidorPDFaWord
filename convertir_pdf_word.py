import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pdf2docx import Converter

# Configura las carpetas
DOWNLOADS_FOLDER = r"C:\Users\estiv\Downloads\Convertidor\Entrada"  # Carpeta de entrada para PDFs
OUTPUT_FOLDER = r"C:\Users\estiv\Downloads\Convertidor\Salida"     # Carpeta de salida para archivos Word

# Asegúrate de que la carpeta de salida exista
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        # Ignorar directorios y archivos que no sean PDF   
        if event.is_directory or not event.src_path.lower().endswith('.pdf'):
            return

        print(f"Nuevo PDF detectado: {event.src_path}")
        
        # Esperar 2 segundos para asegurar que el archivo esté completamente escrito
        time.sleep(2)

        # Verificar si el archivo existe
        pdf_file = event.src_path
        if not os.path.exists(pdf_file):
            print(f"Error: El archivo {pdf_file} no existe o no está accesible.")
            return

        try:
            # Convertir PDF a Word
            docx_file = os.path.join(OUTPUT_FOLDER, os.path.splitext(os.path.basename(pdf_file))[0] + '.docx')
            cv = Converter(pdf_file)
            cv.convert(docx_file)
            cv.close()
            print(f"Convertido a: {docx_file}")

            # Eliminar el PDF original si la conversión fue exitosa
            try:
                os.remove(pdf_file)
                print(f"PDF eliminado: {pdf_file}")
            except Exception as e:
                print(f"Error al eliminar {pdf_file}: {e}")
        except Exception as e:
            print(f"Error al convertir {pdf_file}: {e}")

def monitor_folder():
    event_handler = PDFHandler()
    observer = Observer()
    observer.schedule(event_handler, DOWNLOADS_FOLDER, recursive=False)
    observer.start()
    print(f"Monitoreando la carpeta: {DOWNLOADS_FOLDER}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    monitor_folder() 