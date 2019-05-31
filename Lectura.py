#! python3
Data={}
import docx,os,pyperclip,re,shelve,math

#Gloriosas funciones
def purge(x):#CAMBIAR PARA QUE ACEPTE PARAMETROS
    sacar=[' ',',']
    x=x.upper()
    for i in range(len(x)):
        for i in range(len(sacar)):
            x=x.strip(sacar[i])
    return x
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return ''.join(fullText)
def mostrarlista():
    print()
    for k,v in DataFull.items():
        k=str(k)
        v=str(v)
        print(k.ljust(40,'.')+v.rjust(60, '.'))
    print('Elegi lo que quieras copiar(El numero de la fila)')
def purge_nums(a,b,c):
    a=str(a)
    a=a.split(b)
    a=a[1]
    a=a.split(c)
    a=a[0]
    for _ in range(len(a)):
        a=a.strip()
    return a
def getData(a,b,c):
    a=a.split(b)
    a=a[1]
    a=a.split(c)
    a=a[0]
    return purge(a)
def getNum(bacon,b,c):
    for i in range(len(bacon)-1):
        if (bacon[i].isdigit())==True:
            if len(bacon[i])<c:
                bacon.remove(bacon[i])
            else:
                Data[b]=bacon[i]
            return bacon[i]
def searchanddestroy(eggs,b):
    for i in range((len(eggs)-1)):
        if eggs[i] in b:
            eggs.remove(eggs[i])
    return eggs
    
#Datos necesarios

Nums= {'DI BLASI':'29076875',
       'SALBATIERRA': '35096157',
       'FERNANDEZ': '35417114',
       'REHAK': '31453',
       'ZORNETTA' : '35246371',
       'RODRIGUEZ': '4032',
       'ALARCON' : '34270825'}

Drogas=['CLORHIDRATO',
 'BASE', 
 'MARIHUANA',
 'MDMA', 
 'KETAMINA', 
 'METANFETAMINA']

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
       'DICIEMBRE':'12'}

#Mensaje de bienvenida
print('''Bienvenido al PATA (Programa de Automatizacion de Tareas Administrativas) version 0.
Este programa sirve para leer Actas en formado .DOCX y resumir los datos utilies para que sea mas facil completar tareas administrativas
Eventualmente este programa hara casi todo, pero por ahora solo lee actas y te muestra los datos. Dame mas tiempo!
ACLARACION: solo funciona para archivos que se encuentren en EL MISMO DIRECTORIO que este programa
Cada vez que haya un error(puede ser por error del programa, o porque los oficiales escriben CUALQUIER COSA),
el programa mostrara la seccion donde DEBERIA estar la informacion y te va a pedir que la introduzcas manualmente.
Por favor, ante cualquier duda, consulta 


Por favor introduzca el nombre del archivo a leer. Sin la extension(que deberia ser siempre .docx)
''')

#Aca es el bloque para obtener todo el texto del acta
##Texto completo se llama doc, y el archivo docfull
while True:
    try:
        eggs=input()+'.docx'
        doc=getText(eggs)
        doc=doc.upper()
        docfull= docx.Document(eggs)
        break
    except:
        print('Hubo un error inesperado (ni idea que puede ser) o el archivo que ingresaste NO ES un archivo .docx')
        print('Pruebe de nuevo. Si no funciona y vuelve a ver este mensaje es que se rompio el programa')
        continue

#DESDE EL ENCABEZADO
#Numero de Cooperacion y Sumario

try:
    section = docfull.sections[0]
    header = section.header
    spam=header.paragraphs
    bacon=0
    if bacon<2:
        for paragraph in header.paragraphs:
            if 'COOPERACION' in paragraph.text:
                spam=paragraph.text
            if 'SUMARIO' in paragraph.text:
                eggs=paragraph.text
                bacon+=1

    spam=purge_nums(spam,':','/')
    sumario=re.compile(r'(\d)*/((\d)*)?')
    _=sumario.search(eggs)
    Data['Cooperacion']=purge(spam)
    Data['Sumario']=_.group()
except:
    print('----'*5)
    print('Error en la lectura del encabezado. Por favor, introduzcalo manualmente')
    section = docfull.sections[0]
    header = section.header
    spam=header.paragraphs
    print(spam)
    print('Numero de cooperacion:')
    spam=input()
    print('Numero de sumario:')
    eggs=input()
    Data['Cooperacion']=purge(spam)
    Data['Sumario']=purge(eggs)

#DESDE EL TEXTO

#Suscribiente, LP y Comiseria/Division/etc

##LP
try:
    spam=getData(doc,'SUSCRIBE','DEL NUMERARIO')
    numLP= re.compile(r'(\d)+')
    LP=numLP.search(spam).group()
    Data['LP']= LP
      
except:
    print('Error al leer el LP del suscribiente. Por favor, introduzcalo manualmente')
    print('LP:')
    Data['LP']=input()

##Suscribiente
try:    
    spam=getData(doc,'SUSCRIBE','DEL NUMERARIO')
    eggs=re.compile(r'L(.)?P(.)?')
    _=eggs.search(spam)
    _=_.group()
    _=_.strip(' ')
    spam=spam.split()
    spam.remove(LP)
    spam.remove(_)   
    Data['Suscribiente']=' '.join(spam)
except:
    print('Error al leer el nombre del suscribiente. Por favor, introduzcalo manualmente:')
    spam=input()
    spam=spam.upper()
    Data['Suscribiente']=spam

##Comiseria/Division/etc
try:
    spam=getData(doc,'NUMERARIO DE','A LOS FINES')
    Data['Comiseria/Division/etc']=purge(spam)
except:
    print('Error al leer el nombre del la comiseria/division/etc. Por favor, introduzcalo manualmente:')
    spam=input()
    spam=spam.upper()
    Data['Comiseria/Division/etc']=spam

#Magisterio interventor
MagInterventor=getData(doc,'CON INTERVENCIÓN DE','POR ANTE LA')
SecrInterventora=getData(doc,'POR ANTE LA','EN LA QUE RES')
fiscalia=re.compile(r'(FISCAL.A)')
juzgado=re.compile(r'(JUZGADO)')
NumMagisterio=re.compile(r'((\d)?\d)')
dr=re.compile(r'(DR(.*)?(\w)+)')
if fiscalia.search(MagInterventor):
    Data['Fiscalia nº']=NumMagisterio.search(MagInterventor).group()
    if dr.search(MagInterventor):
        Data['Magistrado interventor']=dr.search(MagInterventor).group()
    else:
        print('Error:No se ha encontrado el nombre del fiscal. Introducilo al final ')
    Data['Secretaria']='UNICA'
    if dr.search(SecrInterventora):
        Data['Secretario']=dr.search(SecrInterventora).group()
elif juzgado.search(MagInterventor):
    Data['Juzgado nº']=NumMagisterio.search(MagInterventor).group()
    if dr.search(MagInterventor):
        Data['Magistrado interventor']=dr.search(MagInterventor).group()       
    else:
        print('Error:No se ha encontrado el nombre del juez. Introducilo al final ')
    Data['Secretaria nº']=NumMagisterio.search(SecrInterventora).group()
    if dr.search(SecrInterventora):
        Data['Secretario']=dr.search(SecrInterventora).group()
else:
    print('Algo terrible paso con la parte del magisterio interventor. Rellenalo al final')

#Imputado VER COMO RESOLVER ESTO.

print('''

Por ahora no se como distinguir el imputado. 
Por ahora necesito que lo escribas manualmente.
Para ayudarte, aca tenes el fragmento del acta que habla de los imputados. Gracias y disculpas!
''')
try:
    spam=getData(doc,'IMPUTAD','A FIN DE')
    print(spam)
    eggs=input()
    eggs=purge(eggs)
    Data['Imputado/s']=eggs
except:
    print('''Error al introducir imputados. 
Mira el acta e introducilo manualmente''')
    Data['Imputado/s']=input()

#Peritos

for k,v in Nums.items():
    if k in doc:
        Data['Nombre de Perito']=k
    if v in doc:
        Data['DNI/LP del Perito']=v
            
#Peso de las Drogas. SOLUCION TEMPORAL POR FALTA DE IDEAS

spam=getData(doc,'ACTO SEGUIDO SE PROCEDE A REALIZAR LA APERTURA DE','FINALIZADO EL PROCEDIMIENTO')
pesos=re.compile(r'\d*,?\.?\d{3}')
bacon=spam.replace('GRAMOS','\n')
bacon=bacon.split('\n')
if 'MDMA' in doc:
	print('Hay MDMA, tenes que introducir manualmente el peso')
	Data['MDMA']=input()

Mari=0
Coca=0
Paco=0
for i in range(0,len(bacon)-1):
	if 'SUSTANCIA' in bacon[i]:
		_=pesos.findall(bacon[i])
		_=_[0]
		_=str(_)
		_=_.replace(',','.')
		_=float(_)
		if 'VEGETAL' in bacon[i]:
			Mari=Mari+_
			Mari=round(Mari,3)
		elif 'BLANCA' in bacon[i]:
			Coca=Coca+_
			Coca=round(Coca,3)
		elif 'AMARI' in bacon[i]:
			Paco=Paco+_
			Paco=round(Paco,3)
if Mari !=0:
	Data['Peso de Marihuana']=str(Mari)
if Coca != 0:
	Data['Peso de Cocaina']=str(Coca)
if Paco != 0:
	Data['Peso de Base de Cocaina']=str(Paco)

#Fecha, Hora de inicio y finalizacion
##Fecha
try:
    spam=getData(doc,'AIRES, HOY',' SIENDO LAS')
    spam=spam.split()
    dia=spam[0]
    for k,v in Meses.items():
        if k == spam[4]:
            mes = v
    año=spam[7]
    Data['Fecha']=(dia+'/'+mes+'/'+año)
except:
    print('''Ha habido un error debido a que la fecha se encuentra mal escrita
Para que la lea el programa, expresarla de la siguiente manera:
hoy XX del Mes de XXXX del año XXXX
Igual, te la dejo pasar por esta vez, introduci la fecha manualmente en formato digital
Por ejemplo: 09/07/1816
Aca te dejo el fragmento del acta donde deberia esta la informacion''')
    eggs=getData(doc,'AIRES, HOY',' SIENDO LAS')
    print(eggs)
    spam=input()
    Data['Fecha']=spam

#Hora inicial y final
try:
	spam=getData(doc,'SIENDO LAS','HORAS')
	if len(spam)==5:
		spam=spam.replace(',',':')
		spam=spam.replace('.',':')
		Data['Hora de inicio']=spam
	else:
		print('''Error en el formato de la hora inicial. El programa solo admite XX:XX
      Introduzca la hora inicial''')
		Data['Hora de inicio']=input()
	spam=getData(doc,'TERMINADO EL ACTO, SIENDO LAS','HORAS')
	if len(spam)==5:
		spam=spam.replace(',',':')
		spam=spam.replace('.',':')
		Data['Hora de finalizacion']=spam
	else:
		print('''Error en el formato de la hora final. El programa solo admite XX:XX 
        Introduzca la hora de finalizacion''')
		Data['Hora de finalizacion']=input()
except:
    print('''Error al leer la hora de inicio o finalizacion en el acta.
    Por favor, introduzcalas manualmente en el formato 
    XX:XX''')
    print('Hora inicial:')
    Data['Hora de inicio']=input()
    print('Hora finalizacion:')
    Data['Hora de finalizacion']=input()

#CHECKEO QUE ESTE TODO BIEN

for k,v in Data.items():
        if len(str(v))>50:
            print('PARECERIA HABER UN PROBLEMA CON EL VALOR '+'""'+str(k)+'"" ya que el valor leido es '+'\n[[['+str(v)+']]]')
            print('Si desea modificarlo, introduzca el valor correcto ahora, Si no presione enter')
            spam=input()
            if spam != '':
                Data[k]=spam
            else:
                continue
        if len(str(v))==0:
            print('PARECERIA HABER UN PROBLEMA CON EL VALOR '+'""'+str(k)+'""ya que el valor leido es '+'\n[[['+str(v)+']]]')
            print('Si desea modificarlo, introduzca el valor correcto ahora, Si no presione enter')
            spam=input()
            if spam != '':
                Data[k]=spam
            else:
                continue
print(Data)
#MAGIA DE LISTA
NewData = [*Data]
IndData=[*Data]
spam= len(NewData)
for eggs in range(0,spam):
    NewData[eggs]= str(eggs+1)+'. '+NewData[eggs]
    IndData[eggs]=str(eggs+1)
DataFull=dict(zip(NewData, list(Data.values())))
IndDataFull=dict(zip(IndData, list(Data.values())))

#Guardar info
Coop=Data['Cooperacion']
shelfFile= shelve.open(Coop)
shelfFile[Coop]=Data
shelfFile.close

#MOSTRAR LISTA

while True:
    print()
    mostrarlista()
    spam= input()
    if spam in IndDataFull:
        pyperclip.copy(IndDataFull[spam])
        while True:
                Repe=input()
                mostrarlista()
                if Repe == '':
                    pyperclip.copy(IndDataFull[spam])
                elif Repe in IndData:
                    pyperclip.copy(IndDataFull[Repe])
                else:
                    print('Por favor, introduzca un numero valido')
                    continue
