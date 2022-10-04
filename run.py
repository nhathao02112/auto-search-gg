import random
from googlesearch import search
import pandas


excel_data_df = pandas.read_excel('input.xlsx')

key = excel_data_df['Name'].tolist()

print(key)


print("Đang tìm")

website = []
linkcrunchbase = []
linkedin = []
marketcap = []
facebook=[]
careerCenter=[]
twitter=[]

ngaunhien = random.randint(3,10)

for i in range(len(key)):
      for j in search(key[i]+"+website", tld="co.in", num=1, stop=11, pause=1):
        print(j) 
        website.append(j)

for e in range(len(key)):
      for f in search(key[e]+"+Twitter", tld="co.in", num=1, stop=1, pause=30):
        print(f) 

        linkcrunchbase.append(f)   

for m in range(len(key)):
      for n in search(key[m]+"+linkedin", tld="co.in", num=1, stop=1, pause=ngaunhien):
        print(n) 

        linkedin.append(n) 
        
for q in range(len(key)):
      for e in search(key[q]+"+facebook", tld="co.in", num=1, stop=1, pause=ngaunhien):
        print(e) 

        facebook.append(e)     

for x in range(len(key)):
      for y in search("career center of +"+key[x], tld="co.in", num=1, stop=1, pause=ngaunhien):
        print(y) 

        careerCenter.append(y)     

for h in range(len(key)):
      for t in search(key[h]+"+twitter", tld="co.in", num=1, stop=1, pause=ngaunhien):
        print(t) 

        twitter.append(t)    

for k in range(len(key)):
      for l in search(key[k]+"+finance+yahoo", tld="co.in", num=1, stop=1, pause=ngaunhien):
        print(l) 

        marketcap.append(l)   

data_series = {"Name": key, "Website" : website, "Crunchbase" : linkcrunchbase, "Linkedin": linkedin, "MarketCap": marketcap}

df_data = pandas.DataFrame(data_series)


df_data.to_excel('output.xlsx')  
print('Save to Excel file successfully.')






