from motet import Motet
from motet import import_motet_data
from motet_utils import motets_ordered_by_difference
from motet_utils import motets_ordered_by_average


motets = import_motet_data("Motet Data.xlsx")


for motet in motets:
	print(motet.title)
	motet.write_to_table()


motets_ordered_by_difference(motets)
motets_ordered_by_average(motets)