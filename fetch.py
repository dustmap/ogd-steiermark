import os
import urllib
import subprocess 
from pyExcelerator import *
import simplejson as json
from datetime import datetime
import pytz

node_id = '1234'
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
local_filename, headers = urllib.urlretrieve(url, '/run/shm/ogd-graz.xls', None, data)


data = {}

for sheet_name, values in parse_xls(local_filename, 'cp1251'):
  matrix = [[]]
  print '"%s"' % sheet_name.encode('cp866', 'backslashreplace')
  print ''
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

  for idx,row in enumerate(matrix):
    # skip xls-header
    if idx<4:
      continue
    csv_row = ','.join(row)
    #print csv_row
    # convert date
    date_str = row[0] + row[1]
    date = datetime.strptime(date_str, '%d.%m.%y%H:%M')
    local = pytz.timezone("Europe/Vienna")
    local_dt = local.localize(date, is_dst=None)
    date_utc = local_dt.astimezone(pytz.utc)
    timestamp_utc = date_utc.strftime("%s")
    #print date_utc.strftime("%Y-%m-%d %H:%M:%S")
    data.update( { timestamp_utc : [ { 'value' : row[2], 'type' : 'dust_concentration' } ] } ) 

os.remove(local_filename)

print json.dumps({ node_id: data })

#eof
