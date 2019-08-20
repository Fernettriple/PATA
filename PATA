#!/usr/bin/env python3

import docx
import os
import pyperclip
import re
import shelve
import math
import csv
import xlsxwriter
import openpyxl 
import shutil
import pandas as pd
from openpyxl import Workbook

    
#Gloriosas funciones
def purge(x):
    '''Es una funcion pedorra que te agarra una palabra y te saca los espacios de ambos lados, y las ","
    '''
    sacar=[' ',',']
    x=x.upper()
    for i in range(len(x)):
        for i in range(len(sacar)):
            x=x.strip(sacar[i])
    return x
def getText(filename):
    '''No se. Lo copie asi como estaba de StackOverflow
    '''
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return ''.join(fullText)
def mostrarlista():
    '''Esta Funcion sirve para Mostrarte todos los elementos cargados en la lista COOP, de una manera mas linda
    '''
    barra()
    for i in range(len(COOP)):
        if COOP[i]=='ERROR EN LA CARGA':
            print('CHECKEAR EL SIGUIENTE VALOR')
            print('V'*80)
        if i < 10:
            print(str(i)+'. '+TITLE[i].ljust(30,'.')+COOP[i].rjust(40, '.'))
        else:
            print(str(i)+'. '+TITLE[i].ljust(29,'.')+COOP[i].rjust(40, '.'))
        if COOP[i]=='ERROR EN LA CARGA':
            print('^'*80)
    if Se_encontraron_drogas_no_soportadas==True:
        print('18. '+Drogas_No_Soportada.ljust(29,'.')+drg[Drogas_No_Soportada].rjust(40, '.'))
    barra()
def barra():
    '''Es una barra, que barrea 80 veces
    '''
    print('~'*80)
def getData(Source,Frase_Inicial,Frase_Final):
    '''Esta funcion sirve para extraer un fragmento de texto variable, que se encuentra entre dos frases constantes en el acta. 
    Si no encuentra la Frase_Inicial o la Frase_Final en el texto, el programa transforma la frase que no se encuentra en REGEX y la busca. 
    Si aun asi, no la encuentra, devuelve ese texto
    '''
    if Frase_Inicial not in Source or Frase_Final not in Source:
        try:
            Source=Transformar_en_ReGex(Frase_Inicial,Source)
            Source=Transformar_en_ReGex(Frase_Final,Source)
        except:
            pass
    try:
        Source=Source.split(Frase_Inicial)
        Source=Source[1]
        Source=Source.split(Frase_Final)
        Source=purge(Source[0])
    except:
        print('Error desconocido al tratar de encontra las frases "{}" y/o "{}". Revisa esa parte del acta'.format(Frase_Inicial,Frase_Final))
    return Source
def Transformar_en_ReGex(Frase_a_regexear,Source):
    '''Esta funcion sirve para agarrar "Frase_a_regexear", la divide en letras, cada letra termina "(letra)?(.)" para que sirva
    para buscarlo usando Regex, en forma de lista. Despues, busca en el Source usando la lista de patrones y busca el match, devolviendotelo
    '''
    dic={}
    Palabras_corregidas=[]
    Palabras=Frase_a_regexear.split()
    for palabra in Palabras:
        palabra=list(palabra)
        if len(palabra)<4:
                continue
        for i in range(len(palabra)):
            nueva_letra='({})?(.)?'.format(palabra[i])
            if i ==0:
                palabra_regexeada =nueva_letra+''.join(palabra[i+1:])
            elif i ==(len(palabra)-1):
                palabra_regexeada= ''.join(palabra[:i])+nueva_letra
            else:
                palabra_regexeada= ''.join(palabra[:i])+nueva_letra+''.join(palabra[i+1:])
            dic[''.join(palabra)+'_'+str(i)]=palabra_regexeada
    for k,v in dic.items():
        Patron=re.compile(v)
        Palabra_mal_escrita=Patron.search(Source)
        if Palabra_mal_escrita:
            Palabras_corregidas.append(Palabra_mal_escrita[0])
            Palabra_bien_escrita=k.split('_')
            Palabra_bien_escrita=Palabra_bien_escrita[0]
            break
    for palabras in Palabras_corregidas:
        if palabras in Source:
            Source=Source.replace(palabras,Palabra_bien_escrita)
    return Source

def replace_string(Doc,Frase_a_cambiar,Frase_cambiada):
    '''Esta funcion agarra un Doc con una tabla (en este caso la hoja de ruta), y escanea por todas las celdas buscando la "Frase_a_cambiar" y la reemplaza por "Frase_cambiada"
    '''
    for table in Doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:                                
                    if Frase_a_cambiar in p.text:
                        inline = p.runs
                        # Loop added to work with runs (strings with same style)
                        for i in range(len(inline)):
                            if Frase_a_cambiar in inline[i].text:
                                text = inline[i].text.replace(Frase_a_cambiar, Frase_cambiada)
                                inline[i].text = text
    return 1     #but why tho
def reemplazar_palabra(Source,*palabra_inicial):
    for palabra in palabra_inicial:
        if palabra in Source:
            Source=Source.replace(palabra,'\n'+palabra)
    return Source

#-------------------------------------------------------------------------------------------------------------------------------
Nums= {'DI BLASI':'29076875',
        'SALBATIERRA': '35096157',
        'FERNANDEZ': '35417114',        
        'REHAK': '31453',
        'ZORNETTA' : '35246371',
        'RODRIGUEZ': '4032',
        'ALARCON' : '34270825',
        'ALARCÓN' : '34270825',
        'LOPEZ':'26.380.362',
        'BUZIN':'6895'
        }
DROGAS=['MDMA', 'MARIHUANA','CLORHIDRATO DE COCAINA','BASE DE COCAINA']
Meses={'ENERO':'01',
        'FEBRERO':'02',
        'MARZO':'03',
        'ABRIL':'04',
        'MAYO' :'05',
        'JUNIO':'06',
        'JULIO':'07',
        'AGOSTO':'08',
        'SEPTIEMBRE':'09',
        'OCTUBRE':'10',
        'NOVIEMBRE':'11',
        'DICIEMBRE':'12'
        } 
#-------------------------------------------------------------------------------------------------------------------------------

#Datos necesarios
Cantidad_de_cooperaciones_leidas=[]
#-------------------------------------------------------------------------------------------------------------------------------
#Creacion de los DF con sus respetivos titulos (Los titulos son usados despues, por eso no los puse adentro)
Titulo_PROD=    ['FECHA DE INGRESO',
                'Nro',
                'Dependencia',
                'Srio. Nro.',
                'Caratula',
                'Cria/Comuna',
                'Magistrado Interventor',
                'Lugar',
                'Peritos',
                'Labor Realizada'
                ]

Titulo_ODI92=[  'N° de procedimiento',
                'Fecha',
                'Hora',
                'JUZGADO/FISCALIA',
                'SECRETARIA',
                'DEPENDENCIA',
                'SUMARIO/CAUSA',
                'Estupefacientes principales',
                'Otros estupefacientes incautados',
                'Cantidad',
                'Unidad de medida'
                ]

Titulo_LIBRO=[  'COOPERACION',
                'GAP',
                'FECHA DE INGRESO',
                'PERITO',
                'JUZGADO/FISCALIA',
                'JUEZ/FISCAL',
                'SECRETARIA',
                'SECRETARIO',
                'NUMERO DE SUMARIO',
                'CARATURA',
                'DAMNIFICADO',
                'IMPUTADO',
                'MATERIAL RECIBIDO',
                'AREA DE TRABAJO',
                'PERITO',
                'FECHA DE SALIDA',
                'DEPENDENCIA QUE RECIBE EL MATERIAL'
                ]

Titulo_Pericias=['FECHA DE INGRESO',
                'FECHA DE TAREA DE CAMPO',
                'FECHA DE ELEVACION',
                'PERITO ASIGNADO',
                'ELEMENTOS A PERITAR',
                'TIPO DE PERICIA',
                'CARATULA',
                'CAUSA',
                'N° DE SUMARIO',
                'DEPENDENCIA INSTRUCTORA',
                'MAGISTRADO INTERVENTOR'
                ]
#-------------------------------------------------------------------------------------------------------------------------------
#Mensaje de bienvenida
print('''
Bienvenido al PATA (Programa de Automatizacion de Tareas Administrativas) version 1.7
Este programa sirve para leer Actas en formado .DOCX y crear automaticamente las hojas de ruta, junto con un Excel llamado 'INFO' que en las 3 paginas te deja listo la informacion para pasar a la ODI92, Pericias pendientes, Productividad y Libro virtual
Cada vez que haya un error(puede ser por error del programa, o porque los oficiales escriben CUALQUIER COSA),
el programa mostrara la seccion donde DEBERIA estar la informacion y te va a pedir que la introduzcas manualmente.
Si cuando leas esa seccion, no encontras la informacion, significa que el oficial no la escribio. Es MUY comun que no pongan la secretaria y cosas asi, asi que tene el acta a mano!
Por favor, ante cualquier duda, consulta

ACLARACION: 
Solo funciona para archivos que se encuentren en EL MISMO DIRECTORIO que este programa
ACLARACION MEGA IMPORTANTE:
EL PROGRAMA TIRA ERRORES SI AL MOMENTO DE INTRODUCIR LOS IMPUTADOS NO SE USA MAYUSCULAS
ASI QUE BUENO... USEN MAYUSCULAS

Desea correr el programa en "Modo Seguro"? (Modo seguro te muestra todos los datos leidos ANTES de pasarlos a los excels
''')
if input('Si/No').upper()=='SI':
    Modo_Seguro=True
else:
    Modo_Seguro=False
CANTIDAD_DE_COOPERACIONES_HECHAS=1
CONTADOR_DE_FILAS_PARA_ODI92=1
wb = Workbook()
ws = wb.active
ws.title = "LIBRO"
sheet=wb['LIBRO']
for ColumnNum in range(len(Titulo_LIBRO)):
    sheet.cell(row=1, column=ColumnNum+1).value =Titulo_LIBRO[ColumnNum]
wb.save('info.xlsx')
ws=wb.create_sheet("PRODUCTIVIDAD")
sheet=wb["PRODUCTIVIDAD"]
for ColumnNum in range(len(Titulo_PROD)):
    sheet.cell(row=1, column=ColumnNum+1).value =Titulo_PROD[ColumnNum]
ws=wb.create_sheet("ODI92")
sheet=wb["ODI92"]
for ColumnNum in range(len(Titulo_PROD)):
    sheet.cell(row=1, column=ColumnNum+1).value =Titulo_ODI92[ColumnNum]
ws=wb.create_sheet("PERICIAS ADEUDADAS")
sheet=wb["PERICIAS ADEUDADAS"]
for ColumnNum in range(len(Titulo_Pericias)):
    sheet.cell(row=1, column=ColumnNum+1).value =Titulo_Pericias[ColumnNum]
wb.save('info.xlsx')

#-------------------------------------------------------------------------------------------------------------------------------
#ACA SE RENOMBRAN LOS DOCX PARA ORDENARSE (IKR)
for filename in os.listdir('.'):
    if (filename!='HDR.docx'
        and filename.startswith("HOJA")==False
        and filename.startswith("-TEMPORAL.docx")==False
        and filename.endswith(".docx")==True):
        try:
            doc=getText(filename)
            doc=doc.upper()
            docfull= docx.Document(filename)
            section = docfull.sections[0]
            header = section.header
            spam=header.paragraphs
            for paragraph in header.paragraphs:
                if 'COOPERACION' in paragraph.text:
                    #Saca el numero de coop para ponerlo como Nombre de archivo
                    Reglon_Coop=paragraph.text
                    break
            Numeros=re.compile(r'((\d)+)')
            Numero_COOP=Numeros.search(Reglon_Coop)
            Numero_COOP=Numero_COOP.group()
            shutil.copy(filename,'SA'+Numero_COOP+' -TEMPORAL.docx')
        except:
            print('Hubo un error inesperado (ni idea que puede ser), pruebe de Frase_cambiada. Si no funciona y vuelve a ver este mensaje es que se rompio el programa o hay algo muy raro con el archivo que esta leyendo('+filename+').'+
            '\n Por favor saquelo de la carpeta y copielo manualmente(vuelva a iniciar el programa)')
            



#DESPUES DE ESE CANCER, EMPIEZA LO DECENTE


for filename in os.listdir('.'):
    if filename.endswith(' -TEMPORAL.docx'): 
        TITLE=['Cooperacion',
                'Sumario',
                'LP',
                'Suscribiente', 
                'Comisaria/Division/etc',
                'Fiscalia/Juzgado',
                'Magistrado interventor',
                'Secretaria', 
                'Secretario',
                'Imputado/s',
                'Perito',   
                'Fecha',
                'Hora de inicio',
                'Hora de finalizacion',
                'Peso de MDMA',
                'Peso de Marihuana',
                'Peso de Clorhidrato de Cocaina',
                'Peso de Base de Cocaina'
                ]
        try:
            doc=getText(filename).upper()
            docfull= docx.Document(filename)
        except:
            print('''
            Hubo un error inesperado (ni idea que puede ser) o el archivo que ingresaste NO ES un archivo .docx')
            Pruebe de Frase_cambiada. Si no funciona y vuelve a ver este mensaje es que se rompio el programa o hay algo muy raro con el archivo que esta leyendo{}.        
            Por favor saquelo de la carpeta y copielo manualmente(vuelva a iniciar el programa)
                '''.format(filename))

        #-------------------------------------------------------------------------------------------------------------------------------
        #Numero de Cooperacion y Sumario
        try:
            section = docfull.sections[0]
            header = section.header
            for paragraph in header.paragraphs:
                #Aca, escanea todos los parrafos en el encabezado y los almacena en las variables
                if 'COOPERACION' in paragraph.text:
                    Reglon_Coop=paragraph.text
                if 'SUMARIO' in paragraph.text:
                    Reglon_Sumario=paragraph.text
            Numero_un_solo_digito=re.compile(r'((\d)+)')
            Numero_COOP=Numero_un_solo_digito.search(Reglon_Coop)
            Numero_de_sumario=re.compile(r'(((\d)+)(/)?((\d)*)?)')
            Numero_SUMARIO=Numero_de_sumario.search(Reglon_Sumario)
            COOP=['ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA','ERROR EN LA CARGA',]
            Se_Encontro_Error_en_la_carga=False
            COOP[0]='SA'+Numero_COOP.group()
            COOP[1]=Numero_SUMARIO.group()
        except:
            print('''
            Error en la lectura del encabezado. Por favor, introduzcalo manualmente. Aca esta el fragmento donde DEBERIA estar los numeros necesarios:
            {}'''.format(spam))
            section = docfull.sections[0]
            header = section.header
            spam=header.paragraphs
            COOP[0]=purge(input('Numero de cooperacion:'))
            COOP[1]=purge(input('Numero de sumario:'))
        barra()
        print('ACTUALMENTE LEYENDO LA COOPERACION : '+COOP[0])
        barra()

        #-------------------------------------------------------------------------------------------------------------------------------
        #LP
        try:
            spam=getData(doc,'SUSCRIBE','NUMERARIO')
            Patron_numero_LP= re.compile(r'(\d)+')
            bugg=re.compile(r'(°|º|/s)?1(°|º|/s)?')
            if bugg.search(spam):
                bug=bugg.search(spam)
                spam=spam.split()
                spam.remove(bug.group())
                spam=' '.join(spam)
            COOP[2]=Patron_numero_LP.search(spam).group()
                    
        except:
            print('Error al leer el LP del suscribiente. Por favor, introduzcalo manualmente. Aca esta el fragmento donde DEBERIA estar los numeros necesarios\n.{}'.format(
                getData(doc,'EL FUNCIONARIO','A LOS FINES LEGALES')))
            COOP[2]=input('LP:   ')

        #-------------------------------------------------------------------------------------------------------------------------------            
        #Suscribiente    
        try:    
            spam=getData(doc,'SUSCRIBE','DEL NUMERARIO').split()
            eggs=re.compile(r'L(.)?P(.)?[^RIMERO]?')
            for i in range(len(spam)):
                try:
                    if eggs.search(spam[i]):
                        _=eggs.search(spam[i])
                        _=_.group()
                        spam.remove(_.strip(' '))
                except:
                    continue
            if COOP[2] in spam:
                spam.remove(Numero_de_LP)
            COOP[3]=' '.join(spam)
        except:
            COOP[3]=input('Error a leer el nombre de suscribiente. Por favor introduzcalo manualmente. Aca esta el fragmento donde DEBERIA estar los numeros necesarios\n{}'.format(getData(doc,'EL FUNCIONARIO','A LOS FINES LEGALES'))).upper()

        #-------------------------------------------------------------------------------------------------------------------------------
        #Comisaria/Division/etc
        try:
            spam=getData(doc,'NUMERARIO DE','A LOS FINES')
            Patron_Numero_de_comisaria=re.compile(r'((\d)?\d)(\w)?')
            if Patron_Numero_de_comisaria.search(spam):
                Numero_de_comisaria=Patron_Numero_de_comisaria.search(spam).group()
                Numero_de_comisaria='CC'+Numero_de_comisaria           
                COOP[4]=purge(spam)
            elif 'DIVIS' in spam:
                spam=spam.split('DIVIS')
                spam='DIVIS'+spam[1]
                COOP[4]=(purge(spam))
                Numero_de_comisaria=purge(spam)
            else:
                Numero_de_comisaria=spam
        except:
            Comisaria=input('Error al leer el nombre de la comisaria/division/etc. Por favor, introduzcalo manualmente:Aca esta el fragmento donde DEBERIA estar los numeros necesarios\n{}'.format(getData(doc,'EL FUNCIONARIO','A LOS FINES LEGALES'))).upper()
            if Patron_Numero_de_comisaria.search(Comisaria):
                Numero_de_comisaria=Patron_Numero_de_comisaria.search(Comisaria).group()
                Numero_de_comisaria='CC'+Numero_de_comisaria
            else:
                Numero_de_comisaria=Comisaria            
            COOP[4]=(purge(Comisaria))

        #-------------------------------------------------------------------------------------------------------------------------------            
        #Magisterio interventor
        try:
            Magisterio_Interventor=getData(doc,'CON INTERVENCIÓN DE','SECRETAR')
            if 'JUZGADO O FISCALIA' in Magisterio_Interventor:
                Magisterio_Interventor.replace('JUZGADO O FISCALIA','')
            Secretaria_Interventora=getData(doc,'SECRETA','EN LA QUE RESULTA')
            fiscalia=re.compile(r'((.)?F(.)?I(.)?S(.)?C(.)?A(.)?L(.)?.(.)?A(.)?)')
            juzgado=re.compile(r'((.)?J(.)?U(.)?Z(.)?G(.)?A(.)?D(.)?O(.)?)')
            Numero_de_Magisterio=re.compile(r'((\d)?(\d)?\d)')
            dr=re.compile(r'(DR(.*)?(\w)+)')
            if fiscalia.search(Magisterio_Interventor):
                Tipo_de_Magisterio='FISCALIA'       #Tipo de Magisterio. Si es Fiscalia o Juzgado (me sirve para distinguir)          
                if Numero_de_Magisterio.search(Magisterio_Interventor):
                    COOP[5]=Numero_de_Magisterio.search(Magisterio_Interventor).group()
                else:
                    COOP[5]=input('Eror: No se ha encontrado el numero del Magisterio, necesito que lo introduzcas manualmente.')
                Numero_de_Fiscalia='FISCALIA'+COOP[5]
                if dr.search(Magisterio_Interventor):
                    if ',' in (dr.search(Magisterio_Interventor).group()):
                        _=(dr.search(Magisterio_Interventor).group()).split(',')
                        _=_[0]
                    COOP[6]=(_)
                else:
                    print('Error:No se ha encontrado el nombre del fiscal. Introducilo Manualmente')
                    COOP[6]=input('Error:No se ha encontrado el nombre del fiscal. Introducilo Manualmente. Aca esta el fragmento donde DEBERIA estar el nombre del fiscal\n{}'.format(Magisterio_Interventor))
                COOP[7]=('ÚNICA')
                if dr.search(Secretaria_Interventora):
                    if ',' in dr.search(Secretaria_Interventora).group():
                        _=dr.search(Secretaria_Interventora).group().split(',')
                        _=_[0]
                    COOP[8]=_
                else:
                    COOP[8]=input('Error:No se ha encontrado el nombre del secretario. Introducilo Manualmente. Aca esta el fragmento donde DEBERIA estar el nombre del secretario\n'+'SECRE'+Secretaria_Interventora+' EN LA QUE RESULTA\n')
            elif juzgado.search(Magisterio_Interventor):
                Tipo_de_Magisterio='JUZGADO' #Tipo de Magisterio. Si es Fiscalia o Juzgado (me sirve para distinguir)
                if Numero_de_Magisterio.search(Magisterio_Interventor):
                    COOP[5]=Numero_de_Magisterio.search(Magisterio_Interventor).group()
                else:
                    COOP[5]=input('Eror: No se ha encontrado el numero del Magisterio, necesito que lo introduzcas manualmente.\n')
                Numero_de_Juzgado='J'+COOP[5]
                if dr.search(Magisterio_Interventor):
                    if ',' in dr.search(Magisterio_Interventor).group():
                        _=dr.search(Magisterio_Interventor).group().split(',')
                        _=_[0]
                    COOP[6]=_
                else:
                    print('Error:No se ha encontrado el nombre del juez. Introducilo manualmente ')
                    COOP[6]=input('Aca esta el fragmento donde DEBERIA estar el nombre del juez\n'+Magisterio_Interventor+'\n')
                if Numero_de_Magisterio.search(Secretaria_Interventora):
                    COOP[7]=Numero_de_Magisterio.search(Secretaria_Interventora).group()
                else:
                    COOP[7]=input('Eror: No se ha encontrado el numero de la secretaria, necesito que lo introduzcas manualmente. Si es una fiscalia y salio este mensaje, pone "UNICA"\n')
                Numero_de_Juzgado_y_Secretaria=Numero_de_Juzgado+' S'+COOP[7]
                if dr.search(Secretaria_Interventora):
                    if ',' in dr.search(Secretaria_Interventora).group():
                        _=dr.search(Secretaria_Interventora).group().split(',')
                        _=_[0]
                    COOP[8]=_
                else:
                    print('Error:No se ha encontrado el nombre del secretario. Introducilo Manualmente')
                    COOP[8]=input('Aca esta el fragmento donde DEBERIA estar el nombre del secretario\n'+'SECRE'+Secretaria_Interventora+' EN LA QUE RESULTA\n')
        except:
            print('Algo terrible paso con la parte del magisterio interventor. Rellenalo manualmente')
            Tipo_de_Magisterio=input('Primero introduzca J o F si es Juzgado o Fiscalia. Si es cualquier otra cosa, este programa no lo soporta. Saque esa Cooperacion y vuelva a correr este programa')
            COOP[5]=(input('Numero de Fiscalia/Juzgado?   '))
            COOP[6]=(input('Nombre del Fiscal/Juez?   '))
            COOP[7]=(input('Numero de Secretaria?(en caso de ser fiscalia poner "UNICA"   '))
            COOP[8]=(input('Nombre del secretario?   '))
        #-------------------------------------------------------------------------------------------------------------------------------            
        #Imputado
        try:
            print('Por ahora no se como distinguir el imputado, necesito que lo escribas manualmente.\n'+
                'Para ayudarte, aca tenes el fragmento del acta que habla de los imputados. Gracias y disculpas!\n{}'.format(getData(doc,'IMPUTADO','A FIN DE')))
            COOP[9]=purge(input())
        except:
            COOP[9]=(input('Error al introducir imputados. Mira el acta e introducilo manualmente porque ocurrio un problema bastante grave a la hora de leer el acta\n'))

        #-------------------------------------------------------------------------------------------------------------------------------            
        #Peritos
        try:
            spam=getData(doc,'AGREGANDO QUE PARA','SE UTILIZARÁN LOS SIGUIENTES REACTIVOS')
            if 'FERNÁNDEZ' in spam:
                spam=spam.replace('FERNÁNDEZ','FERNANDEZ')
            if 'ALARCÓN' in spam:
                spam=spam.replace('ALARCÓN','ALARCON')
            Se_encontro_perito=False
            for k in Nums.keys():
                if k in spam:
                    COOP[10]=(k)
                    Se_encontro_perito=True
            if Se_encontro_perito==False:
                spam=(input('Ocurrio un error a la hora de leer el nombre del perito, por favor introduzcalo, con el LP/DNI separado por un guion. Por ejemplo, Fernandez-35417114 \n')).split('-')
                COOP[10]=spam[0]
                Nums[COOP[10]]=spam[1]
        except:
            COOP[10]=(input('Error CATASTROFICO a la hora de leer el acta en la parte del perito. Debe haber algun error importante en el modelo. Por favor, introduzca el perito manualmente'))

        #-------------------------------------------------------------------------------------------------------------------------------                            
        #Fecha, Hora de inicio y finalizacion
        ##Fecha
        try:
            spam=getData(doc,'AIRES, HOY',' SIENDO LAS').split()
            dia=spam[0]
            for k,v in Meses.items():
                if k == spam[4]:
                    mes = v
            año=spam[7]
            COOP[11]=dia+'/'+mes+'/'+año
        except:
            COOP[11]=input('Ha habido un error debido a que la fecha se encuentra mal escrita, para que la lea el programa, expresarla de la siguiente manera:'+
            'hoy XX del Mes de XXXX del año XXXX, introduci la fecha manualmente en formato digital, Por ejemplo: 09/07/1816. Aca te dejo el fragmento del acta donde deberia esta la informacion\n{}').format(getData(doc,'AIRES, HOY',' SIENDO LAS'))
        ##Hora inicial y final
        try:
            Hora=re.compile(r'((\d)?\d[,:;.]?\d{2})')
            if Hora.search(getData(doc,'SIENDO LAS','HORAS')):
                COOP[12]=Hora.search(getData(doc,'SIENDO LAS','HORAS')).group()
            else:
                COOP[12]=(input('Error en el formato de la hora inicial. El programa solo admite XX:XX. Introduzca la hora inicial'    ))
            if Hora.search(getData(doc,'TERMINADO EL ACTO, SIENDO','SE DA POR FINALIZADA')):
                COOP[13]=Hora.search(getData(doc,'TERMINADO EL ACTO, SIENDO','SE DA POR FINALIZADA')).group()
            else:
                COOP[13]=(input('Error en el formato de la hora final. El programa solo admite XX:XX. Introduzca la hora de finalizacion    '))
        except:
            print('Error al leer la hora de inicio o finalizacion en el acta.\n'+
                'Por favor, introduzcalas manualmente en el formato XX:XX')
            COOP[12]=input('Hora inicial:   ')
            COOP[13]=input('Hora finalizacion:   ')
        #-------------------------------------------------------------------------------------------------------------------------------            
        #Peso de las Drogas. 
        pesos=re.compile(r'\d*,?\.?\d{3}')
        Pericia=reemplazar_palabra(getData(doc,'ACTO SEGUIDO SE PROCEDE A REALIZAR LA APERTURA','FINALIZADO EL PROCEDIMIENTO'),'GRAMOS','SUSTANCIA, MATERIAL')
        if 'GRS' in spam:
            bacon=spam.replace('GRS','GRAMOS\n')
        if 'SUBSTANCIA' in bacon:
            bacon=bacon.replace('SUBSTANCIA','\nSUSTANCIA')
        if 'CIGARRILLO' in bacon:               
            bacon=bacon.replace('CIGARRILLO','\nSUSTANCIA')  
        if 'SEMILLA' in bacon:
            bacon=bacon.replace('SEMILLA','\nSUSTANCIA VEGETAL') 

        bacon=bacon.split('\n') #CON ESTO LOGRO SEPARAR EN FRASES QUE EMPIECEN CON 'SUSTANCIA/MATERIAL' Y TERMINEN CON 'GRAMOS'. SUPONGO QUE ESTO CONTENDRA EL PESO DE LA DROGA
        drg= {}
        if 'MDMA' in doc:
            COOP[14]=input('Hay MDMA, tenes que introducir manualmente el peso\n')
            drg['MDMA']=COOP[14]
        else:
            COOP[14]=('NO HAY')
        Mari=0
        Coca=0
        Paco=0
        for i in range(0,len(bacon)-1):
            if 'SUSTANCIA' in bacon[i] or 'MATERIAL' in bacon[i]:
                if 'PESA' in bacon[i] or 'PESO' in bacon[i]:
                    if pesos.findall(bacon[i]):
                        Peso_de_sustancia=float(str(pesos.findall(bacon[i])[0]).replace(',','.'))
                        mariajuana=['VEGETAL','VERDE AMARRONAD','VERDE','VERDEAMARRONAD','VERDE MARRON']
                        if Peso_de_sustancia>1000:
                            meme=input('Este numero, es un peso de droga o algun otro numero?\n{}. Si esta bien, apreta enter. Si no cualquier cosa y el programa descartara el numero y seguira leyendo\n'.format(str(_)))
                            if meme != '':
                                Peso_de_sustancia=0
                        for Descripciones_de_marihuana in mariajuana:
                            if Descripciones_de_marihuana in bacon[i]:
                                Mari += Peso_de_sustancia
                                Mari=round(Mari,3)
                                bacon[i]='wea'
                        if 'BLANCA' in bacon[i] or 'PULVERULENTA' in bacon[i]:
                            Coca+=Peso_de_sustancia
                            Coca=round(Coca,3)
                            bacon[i]='wea'
                        if 'AMARI' in bacon[i]:
                            Paco+=Peso_de_sustancia
                            Paco=round(Paco,3)
                            bacon[i]='wea'
        #ESTA PARTE SIRVE PARA DETECTAR CUANDO HAY LIMPIO O ESO
        foo=getData(doc,'FINALIZADO EL PROCEDIMIENTO','EL MATERIAL REMANENTE DEVUELTO')
        if 'PLANT' in foo:
            print('Se detecto la presencia de plantas. No hay chance de que pueda automatizar eso, asi que voy a meterlo como marihuana y modificalo en donde sea necesario(necesitaria introducir MUCHAS lineas de codigo para algo que pasa 1 vez por mes, y tengo hambre)')
            COOP[15]=('420')
            drg['MARIHUANA']='420'   
        Drogas_No_Soportadas=['NO CONCLUYENTE','LIMPIO','BICARBONATO','ALMIDON','KETAMINA']
        Se_encontraron_drogas_no_soportadas=False            
        for weas in Drogas_No_Soportadas:
            if weas in foo: 
                print(('Se ha detectado la presencia de algun tipo de sustancia {} la cual NO es soportada por este programa "de manera automatica"\n'+
                'Por favor, introduzca el nombre de la sustancia (asi aparecera en la hoja de ruta y eso) y el peso\n'+
                'Aca tenes el resultado del trunarc para guiarte\n'+foo).format(weas))                   
                Se_encontraron_drogas_no_soportadas=True
                Drogas_No_Soportada=input('Nombre de la droga?\n')
                drg[Drogas_No_Soportada]=input('Peso de la droga?\n')
                break
        if 'CLORHIDRATO DE COCAINA' not in foo:
            if Se_encontraron_drogas_no_soportadas== True:
                Coca=0
        if 'BASE DE COCAINA' not in foo:
            if Se_encontraron_drogas_no_soportadas== True:
                Paco=0
        if Mari != 0:
            Mari=str(Mari)
            Mari=Mari.replace('.',',')
            COOP[15]=(Mari)
            drg['MARIHUANA']=Mari
        else:
            COOP[15]=('NO HAY')
        if Coca != 0:                               
            Coca=str(Coca)
            Coca=Coca.replace('.',',')
            COOP[16]=(Coca)
            drg['CLORHIDRATO DE COCAINA']=Coca 
        else:
            COOP[16]=('NO HAY')
        if Paco !=0:
            Paco=str(Paco)
            Paco=Paco.replace('.',',')
            COOP[17]=(Paco)
            drg['BASE DE COCAINA']=Paco    
        else:
            COOP[17]=('NO HAY')
                                    
        #-------------------------------------------------------------------------------------------------------------------------------
        #CANTIDAD DE SOBRES
        try:
            spam=getData(doc,'ACTO SEGUIDO SE PROCEDE A REALIZAR LA APERTURA','FINALIZADO EL PROCEDIMIENTO')
            Sobr=re.compile(r'(SOBRE(S)?)')
            if Sobr.search(spam):
                _=Sobr.search(spam).group()
                bacon=spam.replace(_,_+'WEA')
                Sobre=getData(bacon,'DE','WEA')
        except:
            Sobre=input('Hubo un error. Cuantos sobres son en esta pericia? Introduzcalo a continuacion, por ejemplo "DOS SOBRES"    ')

        #-------------------------------------------------------------------------------------------------------------------------------            
        #CHECKEO QUE ESTE TODO BIEN. RE HACER TODO. VER COMO HAGO PARA QUE NO       
        for i in range(len(COOP)):
            if len(COOP[i])>70:
                print('El programa detecto que el siguiente dato es muy largo, se puede deber a un problema de lectura o por ahi esta bien. Reviselo por las dudas:')
                print(COOP[i])
                print('Si esta todo bien, presione enter. Si desea cambiarlo, presione cualquier otro boton')
                if input()!='':
                    COOP[i]=input('Introduzca por favor el valor correcto (o una version mas resumida, sin omitir los datos IMPORTANTES)')

        #-------------------------------------------------------------------------------------------------------------------------------            
        #Checkeo al final SI ESTA EN "MODO SEGURO"
        if Se_encontraron_drogas_no_soportadas==False:
            if 'ERROR EN LA CARGA' in COOP:
                barra()
                print('Se detecto un error en la carga de los datos, por favor revise los datos introducidos que se muestras a continuacion')
                Se_Encontro_Error_en_la_carga=True
                barra()
        if (Modo_Seguro==True or 
            Se_Encontro_Error_en_la_carga==True):
            while True:
                mostrarlista()
                spam=input('Desea Modificar algo? Si es asi, introduzca el indice (Numero a la izquierda). En el caso de que este todo bien, apretar enter\n')
                if spam != '':
                    spam=int(spam)
                    COOP[spam]=input('Que valor deberia tener esto?\n')
                else:
                    break
        #-------------------------------------------------------------------------------------------------------------------------------                        
        #CREACION DE LISTAS Y DATAFRAMES Y OTRAS COSAS QUE NECESITE HACER PARA QUE FUNCIONE ESTA WEA QLA
        if Tipo_de_Magisterio=='FISCALIA': #SI ES FISCALIA, VA F+ EL NUMERO SIN SECRETARIA
            PRODUCTIVIDAD=[COOP[11],COOP[0],'DIVISION QUIMICA INDUSTRIAL Y ANALISIS FISICOS Y QUIMICOS', COOP[1], 'INF. LEY 23737',Numero_de_comisaria,Numero_de_Fiscalia,'EDIFICIO CENTRAL',COOP[10],'TEST ORIENTATIVO Y PESAJE' ]
            ODI92=[COOP[0],COOP[11],COOP[12],Numero_de_Fiscalia,'S1',Numero_de_comisaria,COOP[1],]
            LIBRO_VIRTUAL=[COOP[0],'',COOP[11],COOP[10],Numero_de_Fiscalia,COOP[6],COOP[7], COOP[8], COOP[1],'INFRACCION A LA LEY 23.737','LEY Y SOCIEDAD',COOP[9],Sobre,'LEY 23.737', COOP[10],COOP[11],Numero_de_comisaria]
            PERICIAS_ADEUDADAS=[COOP[11],COOP[11],COOP[11],COOP[10],Sobre, 'TEST ORIENTATIVO Y PESAJE','LEY 23737', '',COOP[1],Numero_de_comisaria,Numero_de_Fiscalia]
        elif Tipo_de_Magisterio=='JUZGADO': #SI ES JUZGADO, VA J+NUMERO+S+NUMERO
            PRODUCTIVIDAD=[COOP[11],COOP[0],'DIVISION QUIMICA INDUSTRIAL Y ANALISIS FISICOS Y QUIMICOS', COOP[1], 'INF. LEY 23737',Numero_de_comisaria,Numero_de_Juzgado_y_Secretaria,'EDIFICIO CENTRAL',COOP[10],'TEST ORIENTATIVO Y PESAJE' ]
            ODI92=[COOP[0],COOP[11],COOP[12],Numero_de_Juzgado,COOP[7],Numero_de_comisaria,COOP[1]]
            LIBRO_VIRTUAL=[COOP[0],'',COOP[11],COOP[10],Numero_de_Juzgado,COOP[6],COOP[7], COOP[8], COOP[1],'INFRACCION A LA LEY 23.737','LEY Y SOCIEDAD',COOP[9],Sobre,'LEY 23.737',COOP[10],COOP[11],Numero_de_comisaria]
            PERICIAS_ADEUDADAS=[COOP[11],COOP[11],COOP[11],COOP[10],Sobre, 'TEST ORIENTATIVO Y PESAJE','LEY 23737', '',COOP[1],Numero_de_comisaria,Numero_de_Juzgado_y_Secretaria]
        i=0
        ODI92_EXTRA_DROGA=[]
        for k,v in drg.items(): #con esto agrego las drogas encontradas
            if i==0:
                ODI92.extend([k,'',v,'GRAMOS'])
                i+=1                                    
            else:
                ODI92_EXTRA_DROGA=['','','','','','','',k,'',v,'GRAMOS']

        #-------------------------------------------------------------------------------------------------------------------------------    
        #Pasaje a excels
        wb = openpyxl.load_workbook('info.xlsx')                
        ws = wb.active
        sheet=wb['LIBRO']
        for ColumnNum in range(len(Titulo_LIBRO)):
            sheet.cell(row=CANTIDAD_DE_COOPERACIONES_HECHAS+2, column=ColumnNum+1).value =LIBRO_VIRTUAL[ColumnNum]
        sheet=wb["PRODUCTIVIDAD"]
        for ColumnNum in range(len(Titulo_PROD)):
            sheet.cell(row=CANTIDAD_DE_COOPERACIONES_HECHAS+2, column=ColumnNum+1).value =PRODUCTIVIDAD[ColumnNum]
        sheet=wb["ODI92"]
        for ColumnNum in range(len(Titulo_ODI92)):
            sheet.cell(row=CONTADOR_DE_FILAS_PARA_ODI92+2, column=ColumnNum+1).value =ODI92[ColumnNum]
        if ODI92_EXTRA_DROGA: #SI HAY MAS DE UNA DROGA
            CONTADOR_DE_FILAS_PARA_ODI92+=1
            for ColumnNum in range(len(Titulo_ODI92)):
                sheet.cell(row=CONTADOR_DE_FILAS_PARA_ODI92+2, column=ColumnNum+1).value =ODI92_EXTRA_DROGA[ColumnNum]
        sheet=wb['PERICIAS ADEUDADAS']
        for ColumnNum in range(len(Titulo_Pericias)):
            sheet.cell(row=CANTIDAD_DE_COOPERACIONES_HECHAS+2, column=ColumnNum+1).value =PERICIAS_ADEUDADAS[ColumnNum]
        wb.save('info.xlsx')     

        #-------------------------------------------------------------------------------------------------------------------------------    
        #Exportacion de los datos para para cargar a GAP/SADE
        COOP_PARA_ARCHIVO_CSV=COOP
        COOP_PARA_ARCHIVO_CSV.append(Tipo_de_Magisterio)
        Perito=Nums[COOP[10]]
        COOP_PARA_ARCHIVO_CSV.append(Perito)
        if Se_encontraron_drogas_no_soportadas==True:
            COOP_PARA_ARCHIVO_CSV.append('NS')
            COOP_PARA_ARCHIVO_CSV.append(Drogas_No_Soportada)
            COOP_PARA_ARCHIVO_CSV.append(drg[Drogas_No_Soportada])
        else:
            COOP_PARA_ARCHIVO_CSV.append('')
            COOP_PARA_ARCHIVO_CSV.append('')
            COOP_PARA_ARCHIVO_CSV.append('')
        with open((COOP[0]+'.csv'), 'w', newline='') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
            wr.writerow(COOP_PARA_ARCHIVO_CSV)
            myfile.close()

        #-------------------------------------------------------------------------------------------------------------------------------            
        #ACA VA LA HOJA DE RUTA
        shutil.copy('HDR.docx', 'HOJA DE RUTA DE {}.docx'.format(COOP[0]))
        hdr= docx.Document('HOJA DE RUTA DE {}.docx'.format(COOP[0]))
        replace_string(hdr,'!',COOP[0])
        replace_string(hdr,'$',COOP[11])
        replace_string(hdr,'%',COOP[1])
        replace_string(hdr,'&',Numero_de_comisaria)
        if Tipo_de_Magisterio=='FISCALIA':
            replace_string(hdr,'*',('FISCALIA PENAL CONTRAVENCIONAL Y DE FALTAS N°'+COOP[5]+' A CARGO DE '+COOP[6]))
            replace_string(hdr,'ç',(COOP[7]+' A CARGO DE '+COOP[8]))
        elif Tipo_de_Magisterio=='J':
            replace_string(hdr,'*',('JUZGADO NACIONAL EN LO CRIMINAL Y CORRECCIONAL N°'+COOP[5]+' A CARGO DE '+COOP[6]))
            replace_string(hdr,'ç',(COOP[7]+' A CARGO DE '+COOP[8]))
        else:
            print('Error importante, no puedo identificar si es Fiscalia o Juzgado, se dejara esa parte de la hoja de ruta en blanco')
        replace_string(hdr,'¨',COOP[9])
        replace_string(hdr,'<',(COOP[10]+'  DNI/LP: '+ Nums[COOP[10]]))
        replace_string(hdr,'>',(COOP[3]+' LP: '+COOP[2]))
        replace_string(hdr,'+',COOP[12])
        replace_string(hdr,'-',COOP[13])
        if COOP[14]!='NO HAY':
            replace_string(hdr,'DR1',('MDMA: '+COOP[14]))
        else:
            replace_string(hdr,'DR1','')
        if COOP[15]!='NO HAY':
            replace_string(hdr,'DR2',('M: '+COOP[15]))
        else:
            replace_string(hdr,'DR2','')
        if COOP[16]!='NO HAY':
            replace_string(hdr,'DR3',('CC: '+COOP[16]))
        else:
            replace_string(hdr,'DR3','')
        if COOP[17]!='NO HAY':
            replace_string(hdr,'DR4',('BC: '+COOP[17]))
        else:
            replace_string(hdr,'DR4','')
        try:
            if Se_encontraron_drogas_no_soportadas==True:
                replace_string(hdr,'DR5',(Drogas_No_Soportada+' :'+drg[Drogas_No_Soportada]))
            else:
                replace_string(hdr,'DR5','')
        except:
            print('Estas usando una HDR.docx viejita. Usa la que tiene la fila DR5 asi el programa sirve para drogas no soportadas')
        hdr.save('HOJA DE RUTA DE {}.docx'.format(COOP[0]))                
        CANTIDAD_DE_COOPERACIONES_HECHAS+=1
        CONTADOR_DE_FILAS_PARA_ODI92+=1
        Cantidad_de_cooperaciones_leidas.append(COOP[0])
        continue

#-------------------------------------------------------------------------------------------------------------------------------
#BORRO LOS DOCX CREADOS AL PEDO
for filename in os.listdir():
    if filename.endswith(' -TEMPORAL.docx'):
        os.unlink(filename)
#Cierre y besis
barra()
print('Programa terminado. :^D')
Cantidad_de_cooperaciones_leidas=', '.join(Cantidad_de_cooperaciones_leidas)    
print('Se leyeron las siguientes cooperaciones: {}'.format(Cantidad_de_cooperaciones_leidas))    
barra()
input('Nos vemos la proxima! Besis.\nPresione enter para salir')
