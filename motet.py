
import openpyxl
import polyglot
from pprint import pprint



class Motet:
	def __init__(self, composer, title, triplum, motetus, tenor, triplum_english, motetus_english, tenor_english,	source_link):
		self.composer = composer
		self.title = title
		self.triplum = triplum
		self.motetus = motetus
		self.tenor = tenor
		self.triplum_english = triplum_english
		self.motetus_english = motetus_english
		self.tenor_english = tenor_english
		self.source_link = source_link



wb = openpyxl.load_workbook("Motet Data.xlsx")
sheet = wb['Data']

motets = []
for row in range(2, sheet.max_row + 1):
	composer = sheet['B' + str(row)].value
	title = sheet['C' + str(row)].value
	triplum = sheet['D' + str(row)].value
	motetus = sheet['E' + str(row)].value
	tenor = sheet['F' + str(row)].value
	triplum_english = sheet['G' + str(row)].value
	motetus_english = sheet['H' + str(row)].value
	tenor_english = sheet['I' + str(row)].value
	source_link = sheet['J' + str(row)].value

	NewMotet = Motet(composer, title, triplum, motetus, tenor, triplum_english, motetus_english, tenor_english,	source_link)
	motets.append(NewMotet)

#for row in range(2, sheet.max_row + 1):
	#pprint(sheet['B' + str(row)].value)
	



