
import urllib
import subprocess 
from pyExcelerator import *

url = 'http://app.luis.steiermark.at/luft2/export.php'
values = {
    'station1' : '138',
    'station2' : '',
    'komponente1' : '125',
    'station3' : '164',
    'station4' : '',
    'komponente2' : '114',
    'von_tag' : '1',
    'von_monat' : '6', 
    'von_jahr' : '2013',
    'bis_tag' : '1',
    'bis_monat' : '6', 
    'bis_jahr': '2013',
    'mittelwert' : '1'
  }
data = urllib.urlencode(values)
data = data.encode('utf-8') # data should be bytes

local_filename, headers = urllib.urlretrieve(url, '/tmp/foo2.xls', None, data)
print local_filename

for sheet_name, values in parse_xls(local_filename, 'cp1251'):
  matrix = [[]]
  print 'Sheet = "%s"' % sheet_name.encode('cp866', 'backslashreplace')
  print '----------------'
  for row_idx, col_idx in sorted(values.keys()):
      v = values[(row_idx, col_idx)]
      if isinstance(v, unicode):
          v = v.encode('cp866', 'backslashreplace')
      else:
          v = str(v)
      last_row, last_col = len(matrix), len(matrix[-1])
      while last_row < row_idx:
          matrix.extend([[]])
          last_row = len(matrix)

      while last_col < col_idx:
          matrix[-1].extend([''])
          last_col = len(matrix[-1])

      matrix[-1].extend([v])

  for row in matrix:
      csv_row = ','.join(row)
      print csv_row


