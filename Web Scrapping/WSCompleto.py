from pathlib import Path
import requests
import pandas as pd
from bs4 import BeautifulSoup


def extrair_dados(url: str) -> list[dict]:
    response = requests.get(url, timeout=10)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    titulo_pagina = soup.title.string.strip() if soup.title and soup.title.string else "Sem título"

    h1 = soup.find("h1")
    heading_principal = h1.get_text(strip=True) if h1 else "Sem H1"

    resultados = []
    links = soup.find_all("a")

    for link in links:
        texto_link = link.get_text(strip=True)
        href = link.get("href")

        if href:
            resultados.append({
                "pagina": url,
                "titulo_pagina": titulo_pagina,
                "heading_principal": heading_principal,
                "texto_link": texto_link,
                "url_link": href
            })

    return resultados


def salvar_excel(dados: list[dict], caminho_saida: Path) -> None:
    df = pd.DataFrame(dados)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(caminho_saida, index=False)


def main():
    urls = [
        "https://www.youtube.com/watch?v=B3E0of_EifQ",
        "https://www.iana.org/domains/reserved"
    ]

    output_file = Path("output/resultado_scraping_multiplas_paginas.xlsx")
    todos_os_dados = []

    for url in urls:
        try:
            print(f"Processando: {url}")
            dados = extrair_dados(url)
            todos_os_dados.extend(dados)
            print(f"Links encontrados: {len(dados)}")

        except requests.RequestException as e:
            print(f"Erro na requisição para {url}: {e}")

        except Exception as e:
            print(f"Erro inesperado em {url}: {e}")

    if not todos_os_dados:
        print("Nenhum dado foi coletado.")
        return

    salvar_excel(todos_os_dados, output_file)
    print(f"\nArquivo gerado com sucesso: {output_file}")


if __name__ == "__main__":
    main()