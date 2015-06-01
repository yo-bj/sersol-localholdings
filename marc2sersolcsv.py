#!/usr/bin/env python
 
import csv
from pymarc import MARCReader
from os import listdir
from re import search
 
# change this line to match your folder structure
SRC_DIR = 'C:\your\path\to\MARC\file\folder'
 
# get a list of all .out files in source directory - I'm leaving this as .out for our local workflow, but you can do .mrc as well
file_list = filter(lambda x: search('.out', x), listdir(SRC_DIR))

# note 'wb' and 'rb' below are due to this script running in Windows. If you are running this script in a Linux environment you can probably do away with the 'b' part 
csv_out = csv.writer(open('marc_records.csv', 'wb'), delimiter = '*', quotechar = '"', quoting = csv.QUOTE_MINIMAL)
 
for item in file_list:
  fd = file(SRC_DIR + '/' + item, 'rb')
  reader = MARCReader(fd)
  for record in reader:
    
    # moveon reset to 0. If the record meets any of the conditions below, then moveon is set to 1, which the script will move on to the next record
    moveon = 0
    
    # check to see if record is not suppressed
    bcode3 = ''
    bcode3 = record['998']['e']
    
    # check to see if record is not suppressed
    catdate = ''
    catdate = record['907']['c']

    # check to see if location code is not part of curriculum, reserve, order, crl, and other fun things...
    # loops for multiple location codes
    locationdata = locations = ''
    for locations in record.get_fields('998'):
      locationdata = locations.get_subfields('a')[0]
      if locationdata in ('order', 'currs', 'crl', 'ersrv', 'curre', 'stone'):
        moveon = 1
        break
      if locationdata == 'www':
        recordURL = ''
        recordURL = record['856']['u']
        # substring to find 'http://jy3ke7sv3s.search.serialssolutions.com', conditional to test if True
        if recordURL.startswith('http://jy3ke7sv3s.search.serialssolutions.com'):
          moveon = 1
          break

    # time to do ALL the checks - determine if record should be included in spreadsheet
    if bcode3 != '-' or catdate == '' or  moveon == 1:
      continue # moves on to the next record
        
    # reset ALL the variables... ok, just the ones used for data entry into the file.
    bibLvl = type = isbn = issn = title = author = date = recordNumber = publisher = edition = ''
 
    # determine if record should be treated as a "monograph" [m] or "serial" [s or i] in terms of what to pull from the record using biblvl field
    bibLvl = record['998']['c']
 
    # ISBN/ISSN - we're grabbing the first entry of the first 02X field; hence not trying to sort through all the 02X fields in one record (if multiple fields exist) 
 
    if bibLvl in ('m', 'a'):
      type = 'Book'
      if record['020'] is not None:
       isbn = record.isbn()
      if record['250'] is not None:
       edition = record['250']['a']
      if record['100'] is not None:
       author = record.author()
      if record['250'] is not None:
       edition = record['250']['a']
    elif bibLvl in ('s', 'i', 'b'):
      type = 'Journal'
      if record['022'] is not None:
       issn = record['022']['a']
    
    # author
    if record['100'] is not None:
      author = record.author()
    elif record['110'] is not None:
      author = record['110']['ab']
    elif record['700'] is not None:
      author = record['700']['ab']
    elif record['710'] is not None:
      author = record['710']['ab']
    
    # title: 229 for serials, 245 for gov docs/others
    if bibLvl == 's':
      if record['229'] is not None:
        title = record['229']['a'] 
      else:
        title = record.title()
    else:
      if record['245'] is not None:
        title = record.title()
      # nonetype object attribute error forced me to comment the strip below. Gah.
        # title = title.rstrip('/')
        
    # publisher - since records in our db have pub info in 260 OR 264, I'm pulling this info out manually instead of using record.publisher()
    if record['260'] is not None:
      publisher = record['260']['b']
      date = record.pubyear()
    elif record['264'] is not None:
      publisher = record['264']['b']
      # date = record['264']['c']
    
    # record number
    if record['907'] is not None:
      recordNumber = "https://cat.lib.grinnell.edu/record=" + record['907']['a'].replace('.', '')[:-1]

    # time to clean up punctuation and initial articles!
    if title is not None:
      title = title.rstrip('/.')
      nonFiling = record['245'].indicators[1]
      nonFiling = int(nonFiling)
      title = title[nonFiling:]
    
    if author is not None:
      author = author.rstrip(',.')

    if publisher is not None:
      publisher = publisher.rstrip(';,.')      

    if date is not None:
      date = date.rstrip(',.')  

    if edition is not None:
      edition = edition.rstrip(',.')          

    # order for spreadsheet, what fields correspond to what variables, and what fields are left blank because I'm lazy: 
    # title > title
    # type > determined by bibLvl
    # URL > recordNumber
    # publisher > publisher
    # pubdate > date
    # public note > ''
    # display pubnote > '' 
    # location note > ''
    # display locnote > ''
    # ISSN > isn
    # coverage begin > ''
    # coverage end > ''
    # ISBN > isn
    # author > author
    # editor > '' [yes, I know that we are putting editors in the author field. Deal with it or submit a PR.]
    # edition > edition
    # language ID > ''
    # alphabetization > ''
    
    #csv_out.writerow([isn, title, author, date, recordNumber, publisher, edition, locationdata])
    csv_out.writerow([title, type, recordNumber, publisher, date, '', '', '', '', issn, '',  '', isbn, author, '', edition, '', ''])
  fd.close()

