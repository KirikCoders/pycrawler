import requests,bs4


def get_vtu_page(url, usn):
    form_data = {'usn': usn}
    req = requests.post(url, data=form_data)
    soup = bs4.BeautifulSoup(req.text, "html.parser")
    tags = soup.select('td')
    print(tags[1].getText()+"\n")
    n = len(tags)
    print(n)
    for i in range(4, n-6, 6):
        str = tags[i].getText()+" "+tags[i+1].getText()+" "+tags[i+2].getText()+" "+tags[i+3].getText()+" "+tags[i+4].getText()+" "+tags[i+5].getText()
        print(str+"\n")


url = "http://results.vtu.ac.in/cbcs_17/result_page.php"
start_seq = "1BI15CS"
for i in range(200):
    if len(str(i)) == 1:
        roll = "00"+str(i)
    elif len(str(i)) == 2:
        roll = "0"+str(i)
    else:
        roll = str(i)
    usn = start_seq + roll
    get_vtu_page(url, usn)
