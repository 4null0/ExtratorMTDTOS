#!/usr/bin/env python
# -*- encoding: utf-8 -*- 

import zipfile
import argparse
import datetime
import docx # Module for DOCX manipulation

"""
----------  CREATED BY 4NULL0   ----------
"""
#Colores
rd='\033[1;31m'
gr='\033[1;32m'
yw='\033[1;33m'
bl='\033[1;34m'
pp='\033[1;35m'
cy='\033[1;36m'
xx='\033[0m'

print rd+"\n|-------------------------------|"+xx
print rd+"|  ExtractorMTDTOS by 4NULL0    |"+xx
print rd+"|-------------------------------|"+xx

# Control de parametros
parser = argparse.ArgumentParser()
parser.add_argument("-f", "--fichero", help="Fichero del cual se van a comprobar los metadatos")

args = parser.parse_args() 

if (args.fichero):
	
	nombre = args.fichero
	
	ext = nombre.lower().rsplit(".", 1)[-1]
	
	if ext == "docx" or ext == "xlsx":
		
		#-----Extraemos los metadatos del documento-----
		print yw+"\n[+] Metadata for file:"+gr+" %s" % (nombre)+xx
		# Open the file
		docxFile = docx.Document(file(nombre, "rb"))
		# Data structure with document information
		docxInfo = docxFile.core_properties
		# Print metadata
		attrs = ["author", "category", "comments", "content_status", 
		"created", "identifier", "keywords", "language", 
		"last_modified_by", "last_printed", "modified", 
		"revision", "subject", "title", "version"]
		for attr in attrs: 
			value = getattr(docxInfo,attr)
			if value:
				if isinstance(value, unicode): 
					print "\t-" + cy + str(attr) + ": " + gr + value + xx
				elif isinstance(value, datetime.datetime):  
					print "\t-" + cy + str(attr) + ": " + gr +str(value) + xx
	
		#-----Buscamos metadatos en la información insertada/vinculada-----
		zf = zipfile.ZipFile(args.fichero, "r")
		#Definimos el array que contendrán los metadatos encontrados
		recursos = []	
		#Revisaremos el contenido de cada archivo contenido en el archivo ofimático
		for info in zf.infolist():
			#Determinamos la extensión del archivo a tratar
			ext = info.filename.lower().rsplit(".", 1)[-1]
			
			#Sólo revisaremos los archivos con extensión ".xml"
			if ext == "xml":
				#Obtenemos el contenido del archivo
				datos = zf.read(info.filename).decode('utf-8')
				#Buscamos la primera posición donde aparece el texto: "descr"
				pp = datos.find ("descr=")
				
				#Proceso por el que buscamos todos las variables: "descr", dentro del archivo
				#pp = -1, si no se encuentra la cadena buscada				
				while pp != -1:
					#Extraemos los datos desde la posisición indicada por la variable: pp
					restos = datos[pp:]
					#Buscamos la cadena: />, dentro de la cadena de datos filtrada
					pf = restos.find ("/>")
					#Parseamos la cadena obtenida para quedarnos con los datos de interes
					recurso = restos[7:pf-1]
					
					#Si es el primer dato extraído, lo almacenamos en el array de resultados
					if len(recursos) < 1:
						recursos.append (str(recurso))
						
						print "\t-" + cy + "Descripciones encontradas en las imagenes insertadas/vinculadas: "+str(info.filename)
					else:#Comparamos los nuevos datos con los datos que ya tenemos para no tener datos repetidos
						comparacion = 0
						for i in range (len(recursos)):
							if str(recurso) == str(recursos[i]):
								comparacion = 1
							elif recurso != recursos[i] and comparacion == 1:
								comparacion = 1
							else:
								comparacion = 0
						
						if comparacion == 0:
							recursos.append (str(recurso))
				
					
					restos = restos [pf+1:]
					pp = restos.find ("descr=")
					datos = restos
				#Visualizamos el contenido del array: recursos
				for i in range (len(recursos)):
						print "\t\t\t"+gr+recursos[i]+xx
				#Reiniciamos el array: recursos, para el resto de archivos		
				recursos = []

		zf.close()
	
	print "\n"
else:
	print "\n"+rd+"El argumento: [-f | --file], es obligatorio\n"+xx
