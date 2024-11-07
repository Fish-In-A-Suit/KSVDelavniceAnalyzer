"""
Prebere excel tabelo delavnic in generira tekst tabelo z grupiranimi udeleženci za posamezno delavnico.
"""
from goreverselookup import JsonUtil
from datetime import datetime

import pandas as pd

translation_table = str.maketrans({
	"Č": "C",
	"č": "c",
	"Ć": "C",
	"ć": "c",
	"Ž": "Z",
	"ž": "z",
	"Š": "S",
	"š": "s",
	"Đ": "D",
	"đ": "d"
})

# load the excel file
filepath = "excel.xlsx"
df = pd.read_excel(filepath)

class Termin:
	def __init__(self, termin_label, timestamp):
		self.termin_label = termin_label
		self.timestamp = timestamp

# define the student class
class Student:
	def __init__(self, ime_in_priimek:str, email:str, letnik_studija:int, termini:list):
		self.ime_in_priimek = ime_in_priimek
		self.email = email
		self.letnik_studija = letnik_studija
		self.termini = termini

	def dodaj_termin(self, termin_in:Termin):
		"""
		Doda unikaten termin v listo terminov študenta. Uporabi to funkcijo, da se termini ne podvajajo.
		Če termin_in že obstaja v self.termini, potem se ustrezen termin posodobi z najbližjim timestampom
		"""
		is_unique = True
		for termin_obstojeci in self.termini:
			assert isinstance(termin_obstojeci, Termin)
			if termin_obstojeci.termin_label == termin_in.termin_label:
				is_unique = False

				# update the soonest timestamp za ta termin
				if termin_in.timestamp < termin_obstojeci.timestamp: # timestamps are in the format pandas.Timestamp, thus compare using <
					termin_obstojeci.timestamp = termin_in.timestamp
		
		if is_unique:
			self.termini.append(termin_in)
		

# parse the DataFrame to create Student objects
students = []
termini_set = set()
for _,row in df.iterrows():
	timestamp = row["Timestamp"]
	ime_in_priimek = row["Ime in priimek"].translate(translation_table)
	email = row["Email"].translate(translation_table)
	letnik_studija = int(row["Letnik študija"])
	termini = row["Termini"].translate(translation_table)

	termini_arr = termini.split(', ') if ',' in termini else [termini]
	termini_set.update(termini_arr)

	# convert to Termin classes with timestamps
	termini_arr2 = []
	for t in termini_arr:
		ter = Termin(t, timestamp)
		termini_arr2.append(ter)

	student = Student(ime_in_priimek,email,letnik_studija,termini_arr2)

	# preveri če že obstaja... npr. če se je en prijavil 2x različno za 2 delavnici
	already_exists = False
	for s in students:
		assert isinstance(s, Student)
		if s.ime_in_priimek == student.ime_in_priimek or s.email == student.email:
			for termin in termini_arr2:
				s.dodaj_termin(termin)
			already_exists = True

	# če ne obstaja > dodaj unikatnega študenta
	if already_exists == False:
		students.append(student)

# grupiraj po delavnicah
termin_vs_studenti = {}
for termin_label in termini_set:
	termin_vs_studenti[termin_label] = []

for student in students:
	assert isinstance(student, Student)
	for termin in student.termini:
		assert isinstance(termin,Termin)
		to_append = f"{student.ime_in_priimek} {student.email}"
		to_append = f"{termin.timestamp} | {to_append}"
		termin_vs_studenti[termin.termin_label].append(to_append)

print(f"Printing results:")
for termin,studenti in termin_vs_studenti.items():
	print(f"Termin: {termin} | {len(studenti)} studentov.")
	for student in studenti:
		print(f"  - {student}")

JsonUtil.save_json(termin_vs_studenti, "delavnice_razporeditev.json")


# generate excel files for workshops
for termin_label in termini_set:
	data = []
	student_infos = []
	for student in students:
		assert isinstance(student,Student)
		for student_termin in student.termini:
			assert isinstance(student_termin,Termin)
			if student_termin.termin_label == termin_label:
				# add student data
				student_info = {
					"Timestamp": student_termin.timestamp,
					"Ime in priimek": student.ime_in_priimek,
					"Email": student.email
				}
				student_infos.append(student_info)

	# SORT BY DESCENDING VALUE OF "Timestamp", which is in pandas.Timestamp format
	student_infos = sorted(student_infos, key=lambda x: x["Timestamp"], reverse=False)

	data = student_infos

	# ustvari dataframe za ta termin delavnice
	df_termin = pd.DataFrame(data, columns=["Timestamp", "Ime in priimek", "Email"])
	safe_termin_label = "".join(c if c.isalnum() else "_" for c in termin_label)
	outfilepath = f"delavnice/{safe_termin_label}.xlsx"
	# save to excel file
	df_termin.to_excel(outfilepath, index=False)