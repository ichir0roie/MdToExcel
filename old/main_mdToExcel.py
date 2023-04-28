import mdToArray
import arrayToExcel

from settings import mdToExcel as stg

mta = mdToArray.MdToArray()
ate = arrayToExcel.ArrayToExcel()

mta.read(stg.input_filepath)
ate.setBook(mta.book)
ate.generate(stg.output_filepath)
