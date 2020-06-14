import os,sys
PYTHONPATH =  r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
sys.path.append(PYTHONPATH)
os.environ['PATH'] += ';' + PYTHONPATH
import psse34
import csv
import excelpy
import psspy
import redirect

from os      import chdir, getcwd, system
raiz = getcwd()

# ----------------------------------------------------------------------------------------------------
def array2dict(dict_keys, dict_values):
    '''Convert array to dictionary of arrays.
    Returns dictionary as {dict_keys:dict_values}
    '''
    tmpdict = {}
    for i in range(len(dict_keys)):
        tmpdict[dict_keys[i]] = dict_values[i]
    return tmpdict

def crear_bloc_casos(name_bloc):
    from os import chdir, getcwd, system
    raiz = getcwd()
    from glob   import glob1
    cases = glob1(raiz, '*.sav')
    from sys     import path
    psspy.path(raiz)
    from os import remove
    try:
        remove(name_bloc)
    except:
        pass
    f = open (name_bloc,'w')
    for h in range(0,len(cases),1):
        pass
        texto1= raiz + "\\" + cases[h] 
        f.write(texto1 + '\n')
    f.close()
# ----------------------------------------------------------------------------------------------------
name_bloc="casos.txt"
crear_bloc_casos(name_bloc)

casos = open('casos.txt','r')
lineas = casos.readlines()
enlaces = open('enlaces.txt','r')
enl = enlaces.readlines()

for i in range(len(lineas)):
    lineas[i] = lineas[i].replace(".sav\n",".sav")
    enl[i] = enl[i].replace("\n","")

try:
    remove('Test.xlsx')
except:
    pass
x1 = excelpy.workbook()
x1.show()
x1.worksheet_rename('Interfases', 'Hoja1', overwritesheet=True)
x1.set_cell((1,1),'NOM INTER')
x1.set_cell((1,2),'FLUJO MW')
x1.set_cell((1,3),'LIM MAX MW')
x1.set_cell((1,4),'LIM MIN MW')
x1.set_cell((1,5),'CASO FLUJOS')
x1.set_cell((1,6),'ARCHIVO MONITOR')
num_inter = 0

for k in range(len(lineas)):
    interfases = open(enl[k],'r')
    reng = interfases.readlines()
    inter = []
    interfases = []
    enlaces = []
    enlaces2 = []
    lim1 = []
    lim2 = []
    pos = []
    cont = 0
    fila = 1
    columna = 1
    for i in range(len(reng)):
        if not ("END" in reng[i] or "COM" in reng[i] or "BRANCH" in reng[i]):
            inter.append(reng[i])
            inter[cont] = inter[cont].replace("\n","")
            cont = cont +1
    for i in range(len(inter)):
        if "MONITOR" in inter[i]:
            aux = inter[i].split()
            interfases.append(aux[2])
            lim1.append(aux[4])
            lim2.append(aux[7])
            pos.append(i)                      #indica la posicion de los renglones que dicen "MONITOR"
        else:
            aux = inter[i].split()
            enlaces.append(str(aux[0]) + str(aux [1]) + "|" + str(aux[2]) + "|")
            enlaces2.append(str(aux[0]) + str(aux [1]) + str(aux[2]))
    b = 0
    numctos = []
    while b < (len(pos)-1):                    #Calcula el numero de circuitos por cada enlace y lo coloca en la lista 'numctos'
        b += 1
        a = (pos[b]-pos[b-1])-1
        numctos.append(a)

    tot = 0

    for i in numctos:                          #Calcula el numero de circuitos de todos los enlace
        tot = i + tot
    numctos.append((len(enlaces)-tot))         #Agrega el numero de circuitos del ultimo enlace

    pos1 = [0]
    pos2 = [numctos[0]]
    a = 0
    b = 1

    while b < (len(numctos)):                  #Crea las listas con las posiciones para poder extraer los circuitos que conforman cada enlace
        pos1.append(pos1[a]+numctos[a])
        pos2.append(pos2[a]+numctos[b])
        a += 1
        b += 1

    circxenl = []

    for i in range(len(numctos)):                       #Se forman las listas con los circuitos que forman cada enlace "circxenl"
        circxenl.append(enlaces2[pos1[i]:pos2[i]])

    redirect.psse2py()
    psspy.psseinit(10000)
    caso = lineas[k]
    psspy.case(caso)

    sid = 0
    ierr = psspy.bsys(sid,
                  1,[0.0,400.0],             #kv    filter
                  0,[],                       #area  filter
                  0,[],                        #bus   filter
                  0,[],                        #owner filter
                  0,[])                        #zone  filter

    ierr, ids = psspy.aflowchar(sid,1,6,2, string=["ID"])
    ierr, numeros = psspy.aflowint(sid,1,6,2, string=["FROMNUMBER","TONUMBER"])
    ierr, flujos = psspy.aflowreal(sid,1,6,2, string=["P","Q","MVA","PLOSS"])

    ctosid = []
    for i in range(len(ids[0])):
        clave = str(numeros[0][i]) + str(numeros[1][i]) + str(ids[0][i])
        ctosid.append(clave)

    valorescto = zip(*flujos)

    dictval = array2dict(ctosid, valorescto)

    flujototal = []

    for i in range(len(circxenl)):
        fluxinter = 0
        for j in range(len(circxenl[i])):
            fluxinter = fluxinter + dictval[circxenl[i][j]][0]
        flujototal.append(fluxinter)

    x1.set_cell((num_inter+2,5), caso)
    x1.set_cell((num_inter+2,6), enl[k])
    x1.set_range(num_inter+2,1,zip(interfases))
    x1.set_range(num_inter+2,2,zip(flujototal))
    x1.set_range(num_inter+2,3,zip(lim1))
    x1.set_range(num_inter+2,4,zip(lim2))
    num_inter = num_inter + len(interfases)

x1.save(raiz + '\Test.xls')
