#GERA CSV PARA O GRAFO
import xml.etree.ElementTree as ET

tree = ET.parse('DMRE.xml')
root = tree.getroot()
nome_job = ""
texto = ""
split_texto = []
lista_in = []
lista_out = []
lista_dependencias = []

for job in root.findall("./FOLDER/JOB"):
    lista_in = list(job.iter('INCOND'))
    lista_out = list(job.iter('OUTCOND'))
    nome_job = job.attrib['JOBNAME'].upper().replace('#','_')

    if (len(lista_in) == 0) and (len(list(lista_out)) == 0):#caso o job nao tenha dependentes, imprime somente o nome do job
        lista_dependencias.append(nome_job + ';'+ ' ->' + ';'+ '')
    
    else:
        for incond in lista_in:
            texto = incond.attrib['NAME'].upper().replace('-OK', '').replace('#', '_')
            split_texto = texto.split('-')
            if len(split_texto) != 2:
                lista_dependencias.append(split_texto[0] + ';' + " INCOND INVALIDO" + ';' + nome_job)
            else:
               lista_dependencias.append(split_texto[0] + ';' + " -> " + ';' + nome_job)#printa o nome do job com somente o job a esquerda do hifen de INCOND
        
        for outcond in lista_out:
            texto = outcond.attrib['NAME'].upper().replace('-OK', '').replace('#', '_')
            split_texto = texto.split('-')
            if len(split_texto) != 2:
                lista_dependencias.append(nome_job + ';' + " OUTCOND INVALIDO" + ';' + texto)
            
            else:
                if nome_job != split_texto[1]: #se o nome do job for diferente do outcond a direita do hifen, printa o nome do job com somente o job a direita do hifen de OUTCOND
                    lista_dependencias.append(nome_job + ';' + " -> " + ';' + split_texto[1])

lista_dependencias = list(dict.fromkeys(lista_dependencias))#remove duplicadas
for celula in lista_dependencias:
    print(celula)