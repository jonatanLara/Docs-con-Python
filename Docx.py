import os
import shutil

#import a instalar
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


#variables definidas

#rutas de salida
OUTPUT_PATH ='.\Outputs'
EXCEL_PATH ='.\Inputs\People_Data.xlsx'
ES_WORD_TPL_PATH='.\Inputs\Templates\WordTemplate_ES.docx'
EN_WORD_TPL_PATH='.\Inputs\Templates\WordTemplate_EN.docx'

#EN CASO DE IMPLENTAR IMGS
IMAGE_PATH='.\Inputs\Images'

# Creamos una funcion para crear o elimiar carpetas

def eliminarCrearCarpetas(path):
    if(os.path.exists(path)):
        shutil.rmtree(path)
    os.mkdir(OUTPUT_PATH)

def leerDatos(path, worksheets):
    excel_df = pd.read_excel(path, worksheets)
    return excel_df

def crearWordPersona(df_pers):
    for r_idx, r_val in df_pers.iterrows():
        print(r_idx)
        print(r_val)
        #cargar la plantilla
        l_tpl = ''
        if (r_val['Idioma']=='ES'):
            l_tpl = ES_WORD_TPL_PATH
        elif(r_val['Idioma']=='EN'):
            l_tpl = EN_WORD_TPL_PATH

        #proceso de la plantilla
        docx_tpl = DocxTemplate(l_tpl)

        #añadir una imagen al grafico
        img_path = IMAGE_PATH +'\\'+r_val['Imagen']
        img = InlineImage(docx_tpl, img_path, height=Mm(15))

        #Creamos el contexto
        context = {
            'nombre':r_val['Nombre'],
            'surname1': r_val['Apellido1'],
            'surname2': r_val['Apellido2'],
            'age': r_val['Edad'],
            'picture': img
        }

        #Creamos un render

        docx_tpl.render(context)

        #guardamos el documento
        if (pd.notna(r_val['Apellido2'])):
            nombre_doc = 'Documento_' + r_val['Apellido1'].upper() + '-' + r_val['Apellido2'].upper() + '-' + r_val['Nombre']+'.docx'
        else:
            nombre_doc = 'Documento_' + r_val['Apellido1'].upper() + '-' + r_val['Nombre'] + '.docx'

        docx_tpl.save(OUTPUT_PATH + '\\' + nombre_doc)

def main():
    #
    eliminarCrearCarpetas(OUTPUT_PATH)
    #
    df_personas = leerDatos(EXCEL_PATH,'DATOS')
    #
    crearWordPersona(df_personas)


if __name__ == '__main__':
    main()


