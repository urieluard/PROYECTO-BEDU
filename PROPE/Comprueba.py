import os,sys
import csv
import psse34

PYTHONPATH = r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
sys.path.append(PYTHONPATH)
os.environ['PATH'] += ';' + PYTHONPATH

import excelpy
import psspy
import redirect

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

if (len(lineas) != len(enl)):
    print 'no coinciden los archivos.sav con los archivos de monitoreo de enlaces.txt'
else:

    for i in range(len(lineas)):
        lineas[i] = lineas[i].replace(".sav\n",".sav")
        enl[i] = enl[i].replace("\n","")

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
            else:
                aux = inter[i].split()
                enlaces2.append(str(aux[0]) + str(aux [1]) + str(aux[2]))
    
        redirect.psse2py()
        psspy.psseinit(10000)
        caso = lineas[k]
        psspy.case(caso)

        sid = 0
        ierr = psspy.bsys(sid,
                      1,[230.0,400.0],             #kv    filter
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

        for i in range(len(enlaces2)):
            if not dictval.has_key(enlaces2[i]):
               print 'el enlace', enlaces2[i] , 'del archivo', enl[k], 'NO esta en el caso', caso   