'''
Author: @TenzingMax
Description: This program generates a movie list excel sheet with WebScrapping.
Date: 7/19/2022

'''

import requests
from bs4 import BeautifulSoup
import xlwt;

# Note: repls can't handle displaying xls or xlsx file
#create excel sheet with 2 colums
workbook = xlwt.Workbook()
table = workbook.add_sheet('datap',cell_overwrite_ok=True)
table.write(0, 0, 'Number')
table.write(0, 1, 'Movie Name')

# rottentomatoes url
url = "https://www.rottentomatoes.com/browse/movies_in_theaters/genres:sci_fi~sort:popular?page=1"
# This header is important for letting the server know about the client
headers = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE'
}
#send get request
resp = requests.get(url, headers = headers)
# parse the response to bs4
soup = BeautifulSoup(resp.content, 'lxml')
# use the class selector to get movie titles
movies = soup.find_all(class_="p--small")

# write the movie name line by line
num, line = 0, 1
for anchor in movies:
  if num == 0:
    num = 1
    continue
  if num == 10:
    break
  num += 1
  table.write(line, 0, num)
  table.write(line, 1, anchor.string.strip())
  line += 1
  
#finally save the file
workbook.save('movies.xls')
