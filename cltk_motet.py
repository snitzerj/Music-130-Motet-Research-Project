from cltk.lemmatize.latin.backoff import BackoffLatinLemmatizer
from cltk.corpus.utils.importer import CorpusImporter
from cltk.stem.latin.j_v import JVReplacer
import polyglot
from polyglot.text import Text
from motet_utils import show_polarity
from motet_utils import show_sentiment_words
from motet_utils import negative_word_count
from motet_utils import positive_word_count


corpus_importer = CorpusImporter('latin')
corpus_importer.import_corpus('latin_models_cltk')


lemmatizer = BackoffLatinLemmatizer()
j = JVReplacer()

text = j.replace('''Impudenter circumivi
solum quod mare terminat
indiscrete concupivi
quidquid amantem inquinat
si amo forsan nec amor
tunc pro mercede crucior
aut amor nec in me amor
tunc ingratus efficior
porro cum amor et amo
mater Aeneae media
in momentaneo spasmo
certaminis materia
ex quo caro longe fetet
ad amoris aculeos
quis igitur ultra petet
uri amore hereos?
fas est vel non est mare
fas est. quam ergo virginem?
que meruit bajulare
verum deum et hominem
meruit quod virtuosa
pre cunctis plena gratia
potens munda speciosa
dulcis humilis et pia.
cum quis hanc amat amatur
est ergo grata passio
sui amor quo beatur
amans amoris basio
O maria virgo parens
meum sic ure spiritum
quod amore tuo parens
amorem vitem irritum.''').lower()


lemmatizer = BackoffLatinLemmatizer()

tokens = [token for token in text.split()]

lemmatized = lemmatizer.lemmatize(tokens)

lemmatized_text = " ".join([token[1] for token in lemmatized])

print(lemmatized_text)

show_sentiment_words(lemmatized_text)
print(positive_word_count(lemmatized_text))
print(negative_word_count(lemmatized_text))