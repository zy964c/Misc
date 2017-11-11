from __future__ import unicode_literals
import json
import re
import sys
import requests
import youtube_dl
import concurrent.futures
from bs4 import BeautifulSoup
import time

# video Slatkin = 'https://www.safaribooksonline.com/library/view/effective-python/9780134175249/'
# video Grinberg = 'https://www.safaribooksonline.com/library/view/an-introduction-to/9781491912386/'
# video Sedgwick = 'https://www.safaribooksonline.com/library/view/algorithms-24-part-lecture/9780134384528/'

def parse_course_page():

    """
    safari online video extracter
    usage example: python3.6 safari_dl.py username password
    https://www.safaribooksonline.com/library/view/algorithms-24-part-lecture/9780134384528/
    5500 ~/safari_online/algorithms_sedgewick/ 10
    """
    
    resp = requests.get(video)
    main_site = 'https://www.safaribooksonline.com'

    c = resp.content
    soup = BeautifulSoup(c, "html.parser")
    found = soup.find_all(href = re.compile("\/library\/view\/.*html.*"))
    result = []
    for i in found:
        i1 = str(i)
        start = i1.find('href')
        end = i1.find('html')
        converted = main_site + i1[(start+6):(end+4)]
        if converted not in result:
            result.append(converted)
    return result
    

def download(url, count):

    ydl_opts = {'username': f"{username}", 'password': f"{password}",
                'format': f"[tbr<{qlty}]", 'outtmpl': f"{loctn}{count}-%(title)s.%(ext)s"}

    with youtube_dl.YoutubeDL(ydl_opts) as ydl:
        print(url)
        ydl.download([url])

if __name__ == "__main__":

    start = time.time()

    username = sys.argv[1]
    password = sys.argv[2]
    video = sys.argv[3]
    qlty = sys.argv[4]
    loctn = sys.argv[5]
    threads = sys.argv[6]

    res = parse_course_page()
    print(res)
    
    # for displaying only video format info uncomment these:
    # ydl_opts = {'username': f"{username}", 'password': f"{password}", 'listformats': True}
    # ydl_opts = {'writeinfojson': True}

    # We can use a with statement to ensure threads are cleaned up promptly
    with concurrent.futures.ThreadPoolExecutor(max_workers=int(threads)) as executor:
        # Start the load operations and mark each future with its URL
        future_to_url = {executor.submit(download, url, count): url for count, url in enumerate(res)}

    end = time.time()
    print(end - start)
