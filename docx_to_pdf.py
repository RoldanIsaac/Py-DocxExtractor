import os
# import argparse
from docx import Document
from fpdf import FPDF


class DocxToPdfConverter:
    def __init__(self, font_style="Arial", font_size=12, line_height=6):
        self.font_style = font_style
        self.font_size = font_size
        self.line_height = line_height

    
    def get_docx_inputs(root_folder):
        inputs = []

        # Recorre todas las subcarpetas
        for subdir, _, files in os.walk(root_folder):
            for file in files:
                # Verifica si es un .docx que empieza con "0"
                if file.lower().endswith(".docx") and file.startswith("0"):
                    full_path = os.path.join(subdir, file)
                    inputs.append(full_path)

        return inputs
    
    def extract_text_from_docx(self, docx_path):
        """Extraer texto de un archivo DOCX"""
        try:
            doc = Document(docx_path)
            full_text = []
            for paragraph in doc.paragraphs:
                # Reemplaza caracteres no decodificables
                safe_text = paragraph.text.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
                full_text.append(safe_text)
            return '\n'.join(full_text)
        except Exception as e:
            print(f"Error al leer {docx_path}: {str(e)}")
            return ""
    
    def convert_to_pdf(self, docx_files, output_pdf):
        """Convertir los documentos DOCX a un único PDF"""
        if not docx_files or not output_pdf:
            print("Error: Debe especificar archivos DOCX y un PDF de salida")
            return False
        
        try:
            # Crear PDF
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            # Registrar y configurar fuente Verdana
            pdf.add_font("Verdana", "", "fonts/verdana.ttf", uni=True)
            pdf.add_font("Verdana", "B", "fonts/verdanab.ttf", uni=True)
            pdf.add_font("Times", "", "fonts/times.ttf", uni=True)
            pdf.add_font("Times", "B", "fonts/timesbd.ttf", uni=True)
            # Configurar fuente
            pdf.set_font(self.font_style, "", self.font_size)
            
            print(f"Procesando {len(docx_files)} documentos...")
            
            # Extraer y agregar texto de cada documento
            for i, docx_file in enumerate(docx_files, 1):
                print(f"Procesando documento {i}: {os.path.basename(docx_file)}")
                text = self.extract_text_from_docx(docx_file)
                
                # Agregar número de documento
                pdf.set_font(self.font_style, "B", self.font_size + 2)
                pdf.cell(0, 10, f"Documento #{i}: {os.path.basename(docx_file)}", ln=True)
                pdf.set_font(self.font_style, "", self.font_size)
                
                # Agregar contenido de manera segura | filtrar caracteres inválidos justo 
                # antes de escribirlos:
                safe_text = ''.join(c if c.isprintable() else '?' for c in text)
                pdf.multi_cell(0, self.line_height, safe_text)
                pdf.ln(5)
            
            # Guardar PDF
            pdf.output(output_pdf)
            print(f"PDF guardado exitosamente en: {output_pdf}")
            return True
        
        except Exception as e:
            print(f"Error durante la conversión: {str(e)}")
            return False