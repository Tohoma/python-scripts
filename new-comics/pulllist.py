from urllib.request import urlopen
from bs4 import BeautifulSoup

html = urlopen("http://comics.gocollect.com/new/this/week/dc")
bsObj = BeautifulSoup(html.read(),"html.parser")
comicList = bsObj.findAll("li",{"class":"comic"})
for comic in comicList:
	print(comic.find("strong").get_text())
