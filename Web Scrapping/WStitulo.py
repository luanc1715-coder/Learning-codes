import requests
from bs4 import BeautifulSoup

url = "https://statlocker.gg"

response = requests.get(url, timeout=10)
response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

titulo = soup.title.string.strip()

print("Título da página:", titulo)