
import openpyxl
doc = openpyxl.load_workbook('basededatos.xlsx') #suponiendo que el archivo esta en el mismo directorio del script
hoja = doc.get_sheet_by_name('base')
fichier = open("creabase.json", "a")



base=[]
for i in range(1,3089):
    nombre="A{}".format(i)
    tutor="B{}".format(i)
    matricula="C{}".format(i)
    Emoción="D{}".format(i)
    Académico="E{}".format(i)
    mail="F{}".format(i)
    whatsap="G{}".format(i)
    
    base.append('{'+ '"id"'+':'+'"{}"'.format(i)+','+'"nombre"'+':'+'"{}"'.format(hoja[nombre].value)+','+ '"matricula"'+':'+'"{}"'.format(hoja[matricula].value)+
    ','+ '"tutor"'+':'+'"{}"'.format(hoja[tutor].value)+','+ '"color"'+':'+'"{}"'.format(hoja[Emoción].value)+
    ','+ '"coco"'+':'+'"{}"'.format(hoja[Académico].value)+','+ '"mail"'+':'+'"{}"'.format(hoja[mail].value)+','+ '"whatsap"'+':'+'"{}"'.format(hoja[whatsap].value)+"}")

fichier.write('{"alumnos":[ \n ')
for element in base:
        fichier.write(element+', \n ')
fichier.write('] \n')
fichier.write('}')


    














        

           
