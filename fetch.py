#!/usr/bin/python
# -*- coding: utf-8 -*-


import optparse
import textwrap
import os
import urllib
import subprocess 
from pyExcelerator import *
import simplejson as json
from datetime import datetime
import pytz

DESCRIPTION="""Donwload der Online Luftguetedaten vom Land Steiermark.
Abfrageformular: http://app.luis.steiermark.at/luft2/suche.php
Datenquelle: CC-BY-3.0: Land Steiermark - data.steiermark.gv.at
"""

VERSION='%prog version 1.0'


class IndentedHelpFormatterWithNL(optparse.IndentedHelpFormatter):
  def format_description(self, description):
    if not description: return ""
    desc_width = self.width - self.current_indent
    indent = " "*self.current_indent
# the above is still the same
    bits = description.split('\n')
    formatted_bits = [
      textwrap.fill(bit,
        desc_width,
        initial_indent=indent,
        subsequent_indent=indent)
      for bit in bits]
    result = "\n".join(formatted_bits) + "\n"
    return result

  def format_option(self, option):
    # The help for each option consists of two parts:
    #   * the opt strings and metavars
    #   eg. ("-x", or "-fFILENAME, --file=FILENAME")
    #   * the user-supplied help string
    #   eg. ("turn on expert mode", "read data from FILENAME")
    #
    # If possible, we write both of these on the same line:
    #   -x    turn on expert mode
    #
    # But if the opt string list is too long, we put the help
    # string on a second line, indented to the same column it would
    # start in if it fit on the first line.
    #   -fFILENAME, --file=FILENAME
    #       read data from FILENAME
    result = []
    opts = self.option_strings[option]
    opt_width = self.help_position - self.current_indent - 2
    if len(opts) > opt_width:
      opts = "%*s%s\n" % (self.current_indent, "", opts)
      indent_first = self.help_position
    else: # start help on same line as opts
      opts = "%*s%-*s  " % (self.current_indent, "", opt_width, opts)
      indent_first = 0
    result.append(opts)
    if option.help:
      help_text = self.expand_default(option)
# Everything is the same up through here
      help_lines = []
      for para in help_text.split("\n"):
        help_lines.extend(textwrap.wrap(para, self.help_width))
# Everything is the same after here
      result.append("%*s%s\n" % (
        indent_first, "", help_lines[0]))
      result.extend(["%*s%s\n" % (self.help_position, "", line)
        for line in help_lines[1:]])
    elif opts[-1] != "\n":
      result.append("\n")
    return "".join(result)


def main():
  p = optparse.OptionParser( 
        formatter=IndentedHelpFormatterWithNL(),
        version=VERSION, 
        description=DESCRIPTION,
      )
  p.add_option('--node-id',   default="1234")
  p.add_option('--station1',  default="138")
  p.add_option('--station2',  default="")
  p.add_option('--station3',  default="164")
  p.add_option('--station4',  default="")
  p.add_option('--component1', default="125")
  p.add_option('--component2', default="114")
  p.add_option('--from-date', default="01.06.2013")
  p.add_option('--to-date', default="01.06.2013")
  opts, args = p.parse_args()

  # parse dates
  from_date = datetime.strptime(opts.from_date, '%d.%m.%Y')
  to_date   = datetime.strptime(opts.to_date, '%d.%m.%Y')

  node_id = opts.node_id
  url = 'http://app.luis.steiermark.at/luft2/export.php'
  values = {
      'station1'    : opts.station1,
      'station2'    : opts.station2,
      'komponente1' : opts.component1,
      'station3'    : opts.station3,
      'station4'    : opts.station4,
      'komponente2' : opts.component2,
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
  xls_file, headers = urllib.urlretrieve(url, '/run/shm/ogd-graz.xls', None, data)


  data = {}

  for sheet_name, values in parse_xls(xls_file, 'cp1251'):
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

  os.remove(xls_file)
  print json.dumps({ node_id: data })


if __name__ == '__main__':
  main()

#eof
