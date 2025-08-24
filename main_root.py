import os
from docx_to_pdf import DocxToPdfConverter
from compress_and_remove import compress_and_remove_rar


def run_pipeline(docs_folder, root_folder):
    """
    Ejecuta el flujo completo:
    1. Convierte DOCX a PDFs en distintos formatos.
    2. Si todo va bien, comprime y elimina la carpeta raíz.
    """

    # 1. Buscar DOCX en la carpeta
    docx_inputs = DocxToPdfConverter.get_docx_inputs(docs_folder)

    if not docx_inputs:
        print("⚠️ No se encontraron archivos DOCX que empiecen con '0'.")
        return

    print("Archivos encontrados:")
    for i, docx in enumerate(docx_inputs, 1):
        print(f"{i}. {docx}")

    # 2. Modos de exportación PDF
    modes = [
        # {"font_style": "arial", "font_size": 28, "line_height": 12 * 1.5, "output": "Arial_28_1-5.pdf"},
        {"font_style": "verdana", "font_size": 32, "line_height": 12 * 1.5, "output": "Verdana_32_1-5.pdf"},
        {"font_style": "verdana", "font_size": 32, "line_height": 12 * 1.15, "output": "Verdana_32_1-15.pdf"},
        {"font_style": "verdana", "font_size": 14, "line_height": 9 * 1.15, "output": "Verdana_14_1-15.pdf"},
        {"font_style": "times", "font_size": 26, "line_height": 9 * 1.15, "output": "Times_26_1-15.pdf"},
    ]

    all_success = True  # bandera para saber si todo salió bien

    # 3. Ejecutar conversiones
    for mode in modes:
        print(f"➡️ Procesando: {mode['output']}")
        converter = DocxToPdfConverter(
            font_style=mode['font_style'],
            font_size=mode['font_size'],
            line_height=mode['line_height']
        )
        success = converter.convert_to_pdf(docx_inputs, mode['output'])
        if not success:
            all_success = False
            print(f"❌ Error en la conversión: {mode['output']}")

    # 4. Si todo fue exitoso → comprimir y eliminar la carpeta
    if all_success:
        print("✅ Todas las conversiones se realizaron correctamente.")
        compress_and_remove_rar(root_folder)  # ← usa tu función original
    else:
        print("⚠️ Algunas conversiones fallaron. La carpeta NO será eliminada.")


if __name__ == "__main__":
    docs_folder = r"C:\Users\Orlando\Documents\Writing\@ACTIVIDAD LITERARIA\0000000000"
    root_folder = r"C:\Users\Orlando\Documents\Writing"
    run_pipeline(docs_folder, root_folder)
