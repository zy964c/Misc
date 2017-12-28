import requests
from bs4 import BeautifulSoup
import sys


def rem_tags(urls):
    """
    Remove style tags
    """
    for i in urls:
        req = requests.get(i)
        html = req.text
        s = BeautifulSoup(html, 'html.parser')
        for j in s.find_all():
            if 'style' in j.attrs:
                del j.attrs['style']
        print str(s)
        #with open("output.html", "w") as file:
            #file.write(str(s))
        

if __name__ == "__main__":
    #test = ['http://localhost:8000/test.html']
    #rem_tags(test)
    urls = sys.argv[1:]
    rem_tags(urls)

