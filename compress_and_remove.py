import os
import shutil
import subprocess
from send2trash import send2trash  # pip install Send2Trash

def compress_and_remove_rar(folder_path: str, output_rar: str = None, safe_delete=True):
    """
    Comprime una carpeta en un archivo RAR y luego elimina la carpeta original.
    Requiere tener instalado WinRAR o rar.exe en el PATH.
    """

    # Validar que la ruta existe y es una carpeta
    if not os.path.isdir(folder_path):
        raise ValueError(f"La ruta {folder_path} no es una carpeta válida.")

    # Nombre del archivo RAR
    if output_rar is None:
        output_rar = folder_path.rstrip(os.sep) + ".rar"

    # Evitar sobrescribir un archivo existente
    if os.path.exists(output_rar):
        raise FileExistsError(f"Ya existe un archivo con el nombre {output_rar}")

    try:
        # Ruta al ejecutable de WinRAR (ajusta si lo tienes en otro lugar)
        winrar_path = r"C:\Program Files\WinRAR\rar.exe"
        if not os.path.exists(winrar_path):
            winrar_path = r"C:\Program Files (x86)\WinRAR\rar.exe"
        if not os.path.exists(winrar_path):
            raise FileNotFoundError("No se encontró WinRAR (rar.exe). Instálalo o ajusta la ruta.")

        # Comando para comprimir carpeta en RAR
        cmd = [winrar_path, "a", "-r", output_rar, folder_path]

        # Ejecutar comando
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise Exception(f"Error al crear RAR: {result.stderr}")

        # Si todo salió bien, eliminar la carpeta original
        if safe_delete:
            send2trash(folder_path)  # Enviar a papelera
        else:
            shutil.rmtree(folder_path)  # Eliminar definitivamente

        print(f"✅ Carpeta comprimida en: {output_rar} y eliminada: {folder_path}")

    except Exception as e:
        # Si ocurre un error y existe un archivo .rar incompleto, borrarlo
        if os.path.exists(output_rar):
            os.remove(output_rar)
        print(f"Error: {e}")


# Ejemplo de uso
# folder = r"C:\Users\Orlando\Desktop\On Queue\Py-SpotifyDownloadUtils\CSS"
# compress_and_remove_rar(folder)
