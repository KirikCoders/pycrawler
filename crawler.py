import requests,bs4,os


def get_vtu_page(file,url, usn):
    form_data = {'usn': usn}
    req = requests.post(url, data=form_data)
    soup = bs4.BeautifulSoup(req.text, "html.parser")
    tags = soup.select('td')
    n = len(tags)
    file.write(tags[1].getText() + "\n")
    for i in range(4, n-6, 6):
        str = tags[i].getText()+" "+tags[i+1].getText()+" "+tags[i+2].getText()+" "+tags[i+3].getText()+" "+tags[i+4].getText()+" "+tags[i+5].getText()
        file.write(str+"\n\n\n")


url = "http://results.vtu.ac.in/cbcs_17/result_page.php"
start_seq = "1BI15CS"
pwd = os.getcwd()
path = os.path.join(pwd, 'tests')
if not os.path.exists(path):
    os.makedirs(path)
write_path = os.path.join(path, 'out.txt')
if os.path.exists(write_path):
    os.remove(write_path)
file1 = open(write_path, 'a')
print("working", end='', flush=True)
for i in range(200):
    print('.', end='', flush=True)
    if len(str(i)) == 1:
        roll = "00"+str(i)
    elif len(str(i)) == 2:
        roll = "0"+str(i)
    else:
        roll = str(i)
    usn = start_seq + roll
    get_vtu_page(file1, url, usn)
print("\ndone.")
file1.close()
