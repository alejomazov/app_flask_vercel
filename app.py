from flask import Flask, request, send_file, render_template
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Alignment
import os

def deploy_render():
    
    app = Flask(__name__)

    @app.route('/')
    def index():
        return render_template('upload.html')

    @app.route('/upload', methods=['POST'])
    def upload_file():
        file = request.files['file']
        if file:
            # Guardar el archivo subido temporalmente
            file_path = os.path.join('uploads', file.filename)
            file.save(file_path)

            # Procesar el archivo y generar el archivo resultante
            processed_file_path = procesar_archivo(file_path)

            # Enviar el archivo resultante de vuelta al cliente para su descarga
            return send_file(processed_file_path, as_attachment=True, download_name='archivo_procesado.xlsx')

    def procesar_archivo(archivo_csv):
        # Leer el CSV y hacer modificaciones
        file_to_modificate = pd.read_csv(archivo_csv, decimal=".")
        modificate_file = file_to_modificate.reindex(
            [' BUZAMIENTO', ' DIRECCIóN DE INCLINACIóN', 'X', ' Y', ' Z', ' RUMBO', ' LONGITUD', ' ÁREA'], axis=1)
        modificate_file = modificate_file.rename(columns={
            ' BUZAMIENTO': 'BUZAMIENTO',
            ' DIRECCIóN DE INCLINACIóN': "DIRECCIÓN DE INCLINACIÓN",
            ' Z': 'Z',
            ' RUMBO': 'RUMBO',
            ' LONGITUD': 'LONGITUD (m)',
            ' ÁREA': 'ÁREA (m)'})
        modificate_file["PERSISTENCIA (m)"] = modificate_file["LONGITUD (m)"].apply(
            lambda x: "<1" if x < 1 else "1 a 3" if x < 3 else "3 a 10" if x < 10 else "10 a 20" if x < 20 else ">20")

        # Guardar el archivo en formato Excel
        output_file = os.path.join('processed', 'archivo_procesado.xlsx')
        modificate_file.to_excel(output_file, index=False, float_format="%.3f")

        # Ajustar estilos en el Excel
        ajustar_estilos_excel(output_file)

        # Eliminar el archivo CSV original para mantener limpio el directorio
        os.remove(archivo_csv)
        return output_file

    def ajustar_estilos_excel(output_file):
        workbook = load_workbook(output_file)
        worksheet = workbook.active
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=False)
            cell.border = Border()
            cell.alignment = Alignment(horizontal='left')

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[column_letter].width = max_length + 2

        workbook.save(output_file)
    return app

if __name__ == "__main__":
    app = deploy_render()
    app.run()
