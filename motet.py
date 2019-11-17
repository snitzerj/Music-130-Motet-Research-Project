
import openpyxl
import polyglot
import math
import statistics
from polyglot.downloader import downloader
from polyglot.text import Text
from polyglot.transliteration import Transliterator
from pprint import pprint


def calculate_polarity(text, language_code):
	if not text:
		return 0
	
	analysis_text = Text(str(text), hint_language_code=language_code)
	try:
		polarity = analysis_text.polarity
		return polarity
	except:
		return 0

def show_polarity(text):
	text = Text(str(text))
	print("{:<16}{}".format("Word", "Polarity")+"\n"+"-"*30)
	for w in text.words:
		print("{:<16}{:>2}".format(w, w.polarity))

def show_sentiment_words(text, text_type=""):
	text = Text(str(text))
	print("{:<16}{}".format("Words", "Polarity")+"\n"+"-"*30)	
	for w in text.words:
		polarity = w.polarity
		if polarity != 0:
			print("{:<16}{:>2}".format(w, w.polarity))

def negative_word_count(text, language_code="la"):
	text = Text(str(text), hint_language_code=language_code)
	negative_words = 0
	for w in text.words:
		if w.polarity == -1:
			negative_words += 1
	return negative_words


def positive_word_count(text, language_code="la"):
	text = Text(str(text), hint_language_code=language_code)
	positive_words = 0
	for w in text.words:
		if w.polarity == 1:
			positive_words += 1
	return positive_words


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

	def triplum_polarity(self):
		return calculate_polarity(self.triplum, 'la')

	def motetus_polarity(self):
		return calculate_polarity(self.motetus, 'la')

	def tenor_polarity(self):
		return calculate_polarity(self.tenor, 'la')

	def triplum_english_polarity(self):
		return calculate_polarity(self.triplum_english, 'en')

	def motetus_english_polarity(self):
		return calculate_polarity(self.motetus_english, 'en')

	def tenor_english_polarity(self):
		return calculate_polarity(self.tenor_english, 'en')

	def negative_word_count(self, text_type, language_code='la'):
		if text_type == "triplum":
			if language_code == 'la':
				my_text = self.triplum
			elif language_code == 'en':
				my_text = self.triplum_english
		elif text_type == "motetus":
			if language_code == 'la':
				my_text = self.motetus
			elif language_code == 'en':
				my_text = self.motetus_english
		elif text_type == "tenor":
			if language_code == 'la':
				my_text = self.tenor
			elif language_code == 'en':
				my_text = self.tenor_english

		return negative_word_count(my_text, language_code)

	def positive_word_count(self, text_type, language_code='la'):
		if text_type == "triplum":
			if language_code == 'la':
				my_text = self.triplum
			elif language_code == 'en':
				my_text = self.triplum_english
		elif text_type == "motetus":
			if language_code == 'la':
				my_text = self.motetus
			elif language_code == 'en':
				my_text = self.motetus_english
		elif text_type == "tenor":
			if language_code == 'la':
				my_text = self.tenor
			elif language_code == 'en':
				my_text = self.tenor_english

		return positive_word_count(my_text, language_code)

	def sentiment_difference(self, language_code='la'):
		triplum_rank = self.positive_word_count("triplum", language_code) - self.negative_word_count("triplum", language_code) 
		motetus_rank = self.positive_word_count("motetus", language_code) - self.negative_word_count("motetus", language_code)

		#print(f"Triplum Score: {triplum_rank}  Motetus Score: {motetus_rank}  Total Score: {abs(triplum_rank - motetus_rank)}")
		return abs(triplum_rank - motetus_rank)

	def sentiment_average(self, language_code='la'):
		triplum_scores = (self.positive_word_count("triplum", language_code), self.negative_word_count("triplum", language_code)) 
		motetus_scores = (self.positive_word_count("motetus", language_code), self.negative_word_count("motetus", language_code))

		triplum_rank = statistics.mean(triplum_scores)
		motetus_rank = statistics.mean(motetus_scores)

		#print(f"Triplum Score: {triplum_rank}  Motetus Score: {motetus_rank}  Total Score: {abs(triplum_rank - motetus_rank)}")
		return abs(triplum_rank - motetus_rank)


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



motets.sort(key= lambda x: x.sentiment_difference())

for motet in motets:
	print(f"Title: {motet.title}")
	show_sentiment_words(motet.triplum)
	show_sentiment_words(motet.motetus)

	


















	



