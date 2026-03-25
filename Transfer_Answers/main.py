import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate

OUTPUT_PATH = './Outputs'
EXCEL_PATH = './Inputs/Forms_Data.xlsx'
WORD_TPL_PATH = './Inputs/Templates/Checklist etico Profesores Guías v3 1.docx'

# Rutina para crear / eliminar carpetas (Versión segura para Windows/OneDrive)
def EliminarCrearCarpetas(path):
    # Si no existe, la crea
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
    else:
        # Si ya existe, simplemente borra los archivos que tenga adentro
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path) # Borra el archivo
            except Exception as e:
                print(f"Advertencia: No se pudo borrar un archivo viejo ({e})")

def LeerDatosProyectos(path, worksheet):
    excel_df = pd.read_excel(path, sheet_name=worksheet)
    excel_df.columns = excel_df.columns.str.strip()
    excel_df = excel_df.fillna("") 
    return excel_df

def GenerarChecklists(df_proyectos):
    for r_idx, r_val in df_proyectos.iterrows():
        docx_tpl = DocxTemplate(WORD_TPL_PATH)
        context = r_val.to_dict()
        #nuevo
        context["rut"]=str(r_val.get("rut","")).strip()
        context["nombre"]=str(r_val.get("nombre",'')).strip()

        requiere_evaluacion = False

        for i in range(1, 14):
            col_si = f'p{i}_si'
            col_no = f'p{i}_no'
            
            val_si = str(context.get(col_si, '')).strip().upper()
            val_no = str(context.get(col_no, '')).strip().upper()
            
            if val_si == 'SI':
                context[col_si] = 'X'
                requiere_evaluacion = True
            else:
                context[col_si] = '' 
                
            if val_no == 'NO':
                context[col_no] = 'X'
            else:
                context[col_no] = ''
        #nuevo
        if requiere_evaluacion:
            context['solicito'] = 'solicito'
        else:
            context['solicito'] = 'no solicito'

        docx_tpl.render(context)

        profesor = str(r_val.get('profesor_guia', f'Profesor_{r_idx}')).replace(" ", "_")
        titulo_corto = "_".join(str(r_val.get('titulo', 'Sin_Titulo')).split()[:3])
        
        nombre_doc = f"Checklist_{profesor}_{titulo_corto}.docx"
        
        save_path = os.path.join(OUTPUT_PATH, nombre_doc)
        docx_tpl.save(save_path)
        print(f"Generado exitosamente: {nombre_doc}")

def main():
    print("Iniciando generación de Checklists...")
    EliminarCrearCarpetas(OUTPUT_PATH)

    try:
        df_proyectos = LeerDatosProyectos(EXCEL_PATH, 'DATOS')
        GenerarChecklists(df_proyectos)
        print("--- Todos los documentos han sido generados con éxito ---")
    except Exception as e:
        print(f"Error en el proceso: {e}")

if __name__ == '__main__':
    main()