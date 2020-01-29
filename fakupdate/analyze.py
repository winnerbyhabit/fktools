#! /usr/bin/env python3

import csv
import argparse
import os
import sys

def main():
	parser = argparse.ArgumentParser(prog='FAKanalyze')
	parser.add_argument('faelle_csv', help='csv-Datei mit Fällen oder Personen')
	parser.add_argument('personen_csv', help='csv-Datei mit Personen oder Fällen')
	parser.add_argument('fachschaftenliste_md', help='Markdown-Datei mit der aktuellen Fachschaftenliste')
	args = parser.parse_args()
	
	if(not os.path.isfile(args.faelle_csv)):
		print("Datei {} existiert nicht".format(args.faelle_csv))
		sys.exit()
	if(not os.path.isfile(args.personen_csv)):
		print("Datei {} existiert nicht".format(args.personen_csv))
		sys.exit()
	if(not os.path.isfile(args.fachschaftenliste_md)):
		print("Datei {} existiert nicht".format(args.fachschaftenliste_md))
		sys.exit()
	analyze(args)

def analyze(args):

	faks = set()
	with open(args.faelle_csv, "r") as a, open(args.personen_csv, "r") as b:
		for x in (a,b):
			csvfile = csv.DictReader(x, delimiter=",")
			for row in csvfile:
				faks.add("{} ({})".format(row['Studienfach'], row['angestrebter Abschluss']))


	data = {}
	with open(args.fachschaftenliste_md, "r") as f:
		current_fs = "None"
		
		# skip title
		f.readline()
		f.readline()
		f.readline()
		
		for line in f:
			# fak
			if(line[0:2] == "  "):
				if current_fs not in data:
					data[current_fs] = set()
				data[current_fs].add(line[4:].strip())
			# fs
			elif(line[0] != "-"):
				current_fs = line.strip()
			# else: --- or newline

	current = {}

	for fs in sorted(data):
		for fak in sorted(data[fs]):
			if fak not in current:
				current[fak] = set()
			current[fak].add(fs)
		

	allfaks = set()
	remove_for_fs = {}

	for fak in current:
		allfaks.add(fak)

	for fak in faks:
		allfaks.add(fak)

	with open("FAKDIFF.txt", "w") as f, open("FAKNEW.txt", "w") as faknew, open("FAKREMOVED.txt", "w") as fakremoved:
		for fak in sorted(allfaks):
			if fak in current and fak not in faks:
				print("REMOVED\t{} (was: {})".format(fak, current[fak]), file=f)
				print(fak, file=fakremoved)
				
				for fs in current[fak]:
					if fs not in remove_for_fs:
						remove_for_fs[fs] = set()
					remove_for_fs[fs].add(fak)
				
			if fak in faks and fak not in current:
				print("NEW:   \t{}".format(fak), file=f)
				print(fak, file=faknew)

	with open("FS-REMOVED.md", "w") as f:
		title = "Zu entfernende FAKs je Fachschaft"
		print(title, file=f)
		print("="*len(title), file=f)
		print("", file=f)
		for fs in sorted(remove_for_fs):
			print(fs, file=f)
			print('-'*len(fs), file=f)
			for fak in sorted(remove_for_fs[fs]):
				print("  * {}".format(fak), file=f)
			print("", file=f)
				
	with open("FAKLISTE.txt", "w") as f:
		print("Anzahl der Fach-Abschluss-Kombinationen: {}\n".format(len(faks)), file=f)
		for x in sorted(faks):
			print(x, file=f)

if __name__ == "__main__":
    main()
