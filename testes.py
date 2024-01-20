from bs4 import BeautifulSoup
import requests
r = requests.get("https://go.equatorialenergia.com.br/valor-de-tarifas-e-servicos/#tarifas-grupo-a")
soup = BeautifulSoup(r.content,"html.parser")

data = []
table = soup.find('table',width="577")
table_body = table.find('tbody')

rows = table_body.find_all('tr')
for row in rows:
    cols = row.find_all('td')
    cols = [ele.text.strip() for ele in cols]
    data.append([ele for ele in cols if ele]) 

print(data)