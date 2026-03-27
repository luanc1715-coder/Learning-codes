import requests
from bs4 import BeautifulSoup

url = "https://www.youtube.com/watch?v=B3E0of_EifQ"

response = requests.get(url, timeout=10)
response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

links = soup.find_all("a")

for i, link in enumerate(links, start=1):
    texto = link.get_text(strip=True)
    href = link.get("href")
    print(f"{i}. Texto: {texto} | Link: {href}")