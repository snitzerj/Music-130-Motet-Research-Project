import polyglot
from polyglot.downloader import downloader
from polyglot.text import Text
import statistics
import math
import openpyxl
from openpyxl import Workbook
from cltk.lemmatize.latin.backoff import BackoffLatinLemmatizer
from cltk.corpus.utils.importer import CorpusImporter
from cltk.stem.latin.j_v import JVReplacer

corpus_importer = CorpusImporter('latin')
corpus_importer.import_corpus('latin_models_cltk')
lemmatizer = BackoffLatinLemmatizer()
j = JVReplacer()


def lemmatize(text):
	text = j.replace(text)
	tokens = [token for token in text.split()]
	lemmatized = lemmatizer.lemmatize(tokens)
	lemmatized_text = " ".join([token[1] for token in lemmatized])
	return lemmatized_text


def motets_ordered_by_difference(motets):
	motets.sort(key= lambda x: x.sentiment_difference())
	book = Workbook()
	sheet = book.active
	sheet['A1'] = "Title"
	sheet['B1'] = "Composer"
	sheet['C1'] = "Total Difference Score"

	row = 2
	for motet in motets:
		sheet['A' + str(row)] = motet.title
		sheet['B' + str(row)] = motet.composer
		sheet['C' + str(row)] = motet.sentiment_difference()
		row += 1

	book.save(f"Motets Ordered by Difference.xlsx")

def motets_ordered_by_average(motets):
	motets.sort(key= lambda x: x.sentiment_average())
	book = Workbook()
	sheet = book.active
	sheet['A1'] = "Title"
	sheet['B1'] = "Composer"
	sheet['C1'] = "Total Average Score"

	row = 2
	for motet in motets:
		sheet['A' + str(row)] = motet.title
		sheet['B' + str(row)] = motet.composer
		sheet['C' + str(row)] = motet.sentiment_average()
		row += 1

	book.save(f"Motets Ordered by Average.xlsx")




def calculate_polarity(text, language_code="la"):
	if not text:
		return 0
	
	analysis_text = Text(lemmatize(str(text)), hint_language_code=language_code)
	try:
		polarity = analysis_text.polarity
		return polarity
	except:
		return 0

def show_polarity(text):
	text = Text(lemmatize(str(text)))
	print("{:<16}{}".format("Word", "Polarity")+"\n"+"-"*30)
	for w in text.words:
		print("{:<16}{:>2}".format(w, w.polarity))
	try:
		polarity = text.polarity
	except:
		polariy = 0
	print("{:<16}{:>2}".format("total", polarity))

def show_sentiment_words(text, text_type=""):
	text = Text(lemmatize(str(text)))
	print("{:<16}{}".format("Words", "Polarity")+"\n"+"-"*30)	
	for w in text.words:
		polarity = w.polarity
		if polarity != 0:
			print("{:<16}{:>2}".format(w, w.polarity))
	try:
		polarity = text.polarity
	except:
		polarity = 0
	print("{:<16}{:>2}".format("total", polarity))


def negative_word_count(text, language_code='la'):
	text = Text(lemmatize(str(text)), hint_language_code=language_code)
	negative_words = 0
	for w in text.words:
		if w.polarity == -1:
			negative_words += 1
	return negative_words


def positive_word_count(text, language_code='la'):
	text = Text(lemmatize(str(text)), hint_language_code=language_code)
	positive_words = 0
	for w in text.words:
		if w.polarity == 1:
			positive_words += 1
	return positive_words


def get_sentiment_words(text, language_code='la'):
	words = {}
	polytext = Text(lemmatize(str(text)))
	sum = 0
	for w in polytext.words:
		polarity = w.polarity
		if polarity != 0:
			words[w] = polarity if w not in words else words[w] + polarity
			sum += polarity
	words['Total Average'] = calculate_polarity(text)
	words['Total Sum'] = sum

	return words


		


