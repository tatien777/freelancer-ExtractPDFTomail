import pandas as pd
import requests
from bs4 import BeautifulSoup as bs

url  = "https://www.jotform.com/edit/5329710663913264587?utm_source=emailfooter&utm_medium=email&utm_term=201877586777880&utm_content=edit_submissions&utm_campaign=notification_email_footer_submission_links&email_type=notification"
response = requests.get(url)

soup = bs(response.content, 'html.parser')
rev_div = soup.findAll("div",{"id":"cid_19"})

print(soup)
