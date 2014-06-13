# -*- coding: utf-8 -*-

"""
finale Version vom 22.03.2014,
auf Python 3 umgestellt am 13.06.2014
"""

import jinja2 
import os
import mysql.connector
import codecs

# siehe http://e6h.de/post/11/ (leider offline)
# Im LaTeX Dokument müssen Variablen wie folgt 
# deklariert werden, Beispiele
# \BLOCK{for item in liste}
# \VAR{item[0]}

latex_jinja_env = jinja2.Environment(
    block_start_string = '\BLOCK{',
    block_end_string = '}',
    variable_start_string = '\VAR{',
    variable_end_string = '}',
    comment_start_string = '\#{',
    comment_end_string = '}',
    line_statement_prefix = '%-',
    line_comment_prefix = '%#',
    trim_blocks = True,
    autoescape = False,
    loader = jinja2.FileSystemLoader(os.path.abspath('.'))
)

# Laden des Templates aus einer Datei
template = latex_jinja_env.get_template('Sammelbestaetigung_Geldzuwendung.tex')

def getDBsettings():
	""" 
	Holt sich die Datenbank-Einstellungen aus einer externen Datei (kann man auch hart reincoden)
	Hinweis: wird ausgetauscht werden, dies setzt voraus, dass in der letzten Datenzeile auch ein Umbruch ist
	"""
	settings = {}
	with open("g:/dbsettings.txt") as myfile:		
		for line in myfile:
			name, var = line.partition("=")[::2]
			settings[name.strip()] = var[:-1]
	return settings

settings= getDBsettings()

db = mysql.connector.connect(host=settings["server"],user=settings["login"],password=settings["password"], 
db = settings["database"],charset="utf8",use_unicode=True)

# definiert drei verschiedene Cursors
# a) für Stammdaten
# b) für Adressen
# c) für die Buchungen

master_cursor = db.cursor()
address_cursor = db.cursor()
transactions_cursor = db.cursor()

# holt alle IDs aus der Stammdaten-Tabelle, die auch in der Buchungstabelle vorhanden sind.
master_cursor.execute("select ID from Stammdaten where ID in (Select distinct(Klasse) from Buchungen)") 

# definiert das Verzeichnis für die TeX und PDF Dateien
os.chdir("E:/_output/")

# Welche Kategorien sollen mit ausgegeben werden
kategorien = "('Aufnahmegebühr','Zweckspende','Mitgliedsbeitrag','Spende')"

# Für jeden Eintrag in den Masterdaten tue folgendes
for (ID) in master_cursor.fetchall():
	ID = ID[0] # hole die ID aus der Liste raus
	print("\nID = " + str(ID)) # Ausgabe für Logging Zwecke

	# holt die Adresse ab über eine selbstdefinierte Funktion
	# kann noch verbessert werden, gibt ",," aus wenn Feld nicht gefüllt.
	address = address_cursor.execute("select fs_getAddress(" + str(ID) + ")")
	address =  address_cursor.fetchall()[0][0]
	
	# Hole alle Buchungen, die zu der Mitglieds-ID gehört und die zu den o.e. Kategorien gehört
	transactions_cursor.execute("select DATE_FORMAT(Datum, '%d.%m.%Y'), Kategorie, replace(ROUND(Betrag, 2),\".\",\",\") from Buchungen where Kategorie in " + kategorien + " and Klasse=" + str(ID))
	transactions =  transactions_cursor.fetchall()
	# 'transactions' enthält jetzt eine Liste aller Buchungen

	# Hier recycle ich den transactions Cursor um die Summe der Buchungen abzuholen und in 'summe' zu speichern 
	transactions_cursor.execute("select replace(ROUND(sum(Betrag), 2),\".\",\",\") from Buchungen where Kategorie in " + kategorien +  " and Klasse=" + str(ID))
	summe = transactions_cursor.fetchall()[0][0]
	# print (summe)

	# erneutes Recycling des Cursors, um die Kardinalzahl für die Summe zu holen
	transactions_cursor.execute("select ID, kardinal from kardinal where ID = (select round(sum(Betrag),0) from Buchungen where Kategorie in " + kategorien +  " and Klasse=" + str(ID) + ")" )
	for entry in transactions_cursor.fetchall():
		Nummer = entry[0]
		kardinal = entry[1]
	
	# jetzt wird in 'dokument' der TeX Code abgelegt
	dokument = template.render(Spender=address, ID=ID,Summe=summe,kardinal=kardinal,liste=transactions)
	# das Dokument wird gespeichert
	with codecs.open (str(ID) + ".tex", "w","utf-8") as letter:
		letter.write(dokument);
		letter.close();
		# Aufruf von pdflatex
		os.system("pdflatex " + str(ID) + ".tex")
