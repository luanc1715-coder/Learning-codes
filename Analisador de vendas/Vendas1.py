import os
import sys
import subprocess
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def abrir_arquivo(caminho: Path) -> None:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    if sys.platform.startswith("win"):
        os.startfile(caminho)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(caminho)], check=True)
    else:
        subprocess.run(["xdg-open", str(caminho)], check=True)


def processar_planilhas(input_dir: Path, output_dir: Path) -> tuple[Path, int]:
    output_dir.mkdir(exist_ok=True)

    arquivos_excel = list(input_dir.glob("*.xlsx"))

    if not arquivos_excel:
        raise FileNotFoundError("Nenhuma planilha .xlsx encontrada na pasta selecionada.")

    lista_dfs = []

    for arquivo in arquivos_excel:
        df = pd.read_excel(arquivo)
        df["Arquivo_Origem"] = arquivo.name
        lista_dfs.append(df)

    df = pd.concat(lista_dfs, ignore_index=True)

    colunas_necessarias = ["Produto", "Quantidade", "Valor_Unitario"]
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente: {coluna}")

    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce")
    df["Valor_Unitario"] = pd.to_numeric(df["Valor_Unitario"], errors="coerce")

    df = df.dropna(subset=["Produto", "Quantidade", "Valor_Unitario"]).copy()

    if df.empty:
        raise ValueError("Após a limpeza dos dados, nenhuma linha válida restou.")

    df["Total"] = df["Quantidade"] * df["Valor_Unitario"]

    total_vendido = df["Total"].sum()
    total_unidades = df["Quantidade"].sum()
    media_por_unidade = total_vendido / total_unidades if total_unidades != 0 else 0

    resumo_produtos = df.groupby("Produto").agg({
        "Quantidade": "sum",
        "Total": "sum"
    })

    ranking_faturamento = resumo_produtos.sort_values(by="Total", ascending=False)
    ranking_quantidade = resumo_produtos.sort_values(by="Quantidade", ascending=False)

    resumo_geral = pd.DataFrame({
        "Métrica": [
            "Faturamento total",
            "Total de unidades vendidas",
            "Média por unidade",
            "Quantidade de arquivos processados",
            "Quantidade de linhas válidas"
        ],
        "Valor": [
            total_vendido,
            total_unidades,
            media_por_unidade,
            len(arquivos_excel),
            len(df)
        ]
    })

    vendas_por_arquivo = df.groupby("Arquivo_Origem").agg({
        "Quantidade": "sum",
        "Total": "sum"
    }).sort_values(by="Total", ascending=False)

    arquivo_saida = output_dir / "relatorio_vendas_consolidado.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="xlsxwriter") as writer:
        resumo_geral.to_excel(writer, sheet_name="Resumo", index=False)
        ranking_faturamento.to_excel(writer, sheet_name="Ranking_Faturamento")
        ranking_quantidade.to_excel(writer, sheet_name="Ranking_Quantidade")
        vendas_por_arquivo.to_excel(writer, sheet_name="Por_Arquivo")
        df.to_excel(writer, sheet_name="Base_Consolidada", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Ranking_Faturamento"]

        chart = workbook.add_chart({"type": "column"})
        chart.add_series({
            "name": "Faturamento por Produto",
            "categories": ["Ranking_Faturamento", 1, 0, len(ranking_faturamento), 0],
            "values": ["Ranking_Faturamento", 1, 2, len(ranking_faturamento), 2],
        })
        chart.set_title({"name": "Faturamento por Produto"})
        chart.set_x_axis({"name": "Produto"})
        chart.set_y_axis({"name": "Valor Vendido"})
        worksheet.insert_chart("E2", chart)

    return arquivo_saida, len(arquivos_excel)


class SalesAnalyzerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Sales Analyzer")
        self.root.geometry("700x380")
        self.root.resizable(False, False)

        self.base_dir = Path(__file__).parent
        self.input_dir = self.base_dir / "input"
        self.output_dir = self.base_dir / "output"

        self.input_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)

        self.caminho_var = tk.StringVar(value=str(self.input_dir))
        self.status_var = tk.StringVar(value="Pronto para processar planilhas.")
        self.arquivos_var = tk.StringVar(value="Arquivos encontrados: 0")
        self.ultimo_relatorio: Path | None = None

        self.criar_interface()
        self.atualizar_contagem_arquivos()

    def criar_interface(self):
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        titulo = tk.Label(
            frame,
            text="Analisador de Vendas",
            font=("Arial", 16, "bold")
        )
        titulo.pack(pady=(0, 15))

        label_pasta = tk.Label(frame, text="Pasta com planilhas .xlsx:")
        label_pasta.pack(anchor="w")

        entrada_frame = tk.Frame(frame)
        entrada_frame.pack(fill="x", pady=5)

        entrada = tk.Entry(
            entrada_frame,
            textvariable=self.caminho_var,
            width=70
        )
        entrada.pack(side="left", fill="x", expand=True)

        botao_procurar = tk.Button(
            entrada_frame,
            text="Procurar",
            command=self.selecionar_pasta
        )
        botao_procurar.pack(side="left", padx=(10, 0))

        info_arquivos = tk.Label(
            frame,
            textvariable=self.arquivos_var,
            anchor="w",
            font=("Arial", 10, "italic")
        )
        info_arquivos.pack(anchor="w", pady=(8, 0))

        botoes_frame = tk.Frame(frame)
        botoes_frame.pack(pady=20)

        botao_gerar = tk.Button(
            botoes_frame,
            text="Gerar Relatório",
            width=18,
            height=2,
            command=self.gerar_relatorio
        )
        botao_gerar.pack(side="left", padx=5)

        botao_abrir = tk.Button(
            botoes_frame,
            text="Abrir Relatório",
            width=18,
            height=2,
            command=self.abrir_relatorio
        )
        botao_abrir.pack(side="left", padx=5)

        botao_atualizar = tk.Button(
            botoes_frame,
            text="Atualizar Contagem",
            width=18,
            height=2,
            command=self.atualizar_contagem_arquivos
        )
        botao_atualizar.pack(side="left", padx=5)

        status_titulo = tk.Label(frame, text="Status:")
        status_titulo.pack(anchor="w")

        status_label = tk.Label(
            frame,
            textvariable=self.status_var,
            justify="left",
            wraplength=640,
            bg="#f0f0f0",
            anchor="w",
            relief="sunken",
            padx=10,
            pady=10
        )
        status_label.pack(fill="x", pady=(5, 0))

    def contar_arquivos_excel(self, pasta: Path) -> int:
        if not pasta.exists() or not pasta.is_dir():
            return 0
        return len(list(pasta.glob("*.xlsx")))

    def atualizar_contagem_arquivos(self):
        pasta = Path(self.caminho_var.get())
        quantidade = self.contar_arquivos_excel(pasta)
        self.arquivos_var.set(f"Arquivos encontrados: {quantidade}")

        if quantidade > 0:
            self.status_var.set(f"Foram encontrados {quantidade} arquivo(s) .xlsx na pasta selecionada.")
        else:
            self.status_var.set("Nenhum arquivo .xlsx encontrado na pasta selecionada.")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(
            title="Selecione a pasta com as planilhas"
        )
        if pasta:
            self.caminho_var.set(pasta)
            self.atualizar_contagem_arquivos()

    def gerar_relatorio(self):
        try:
            input_dir = Path(self.caminho_var.get())

            if not input_dir.exists():
                raise FileNotFoundError("A pasta selecionada não existe.")

            arquivo_saida, quantidade_arquivos = processar_planilhas(input_dir, self.output_dir)
            self.ultimo_relatorio = arquivo_saida

            mensagem = (
                f"Relatório gerado com sucesso.\n"
                f"Arquivos processados: {quantidade_arquivos}\n"
                f"Arquivo salvo em:\n{arquivo_saida}"
            )
            self.status_var.set(mensagem)
            self.arquivos_var.set(f"Arquivos encontrados: {quantidade_arquivos}")
            messagebox.showinfo("Sucesso", mensagem)

        except Exception as e:
            self.status_var.set(f"Erro: {e}")
            messagebox.showerror("Erro", str(e))

    def abrir_relatorio(self):
        try:
            if self.ultimo_relatorio is None:
                arquivo_padrao = self.output_dir / "relatorio_vendas_consolidado.xlsx"
                if arquivo_padrao.exists():
                    self.ultimo_relatorio = arquivo_padrao
                else:
                    raise FileNotFoundError("Nenhum relatório foi gerado ainda.")

            abrir_arquivo(self.ultimo_relatorio)
            self.status_var.set(f"Relatório aberto: {self.ultimo_relatorio}")

        except Exception as e:
            self.status_var.set(f"Erro ao abrir relatório: {e}")
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        print("A biblioteca xlsxwriter não está instalada.")
        print("Instale com: python -m pip install xlsxwriter")
        sys.exit()

    root = tk.Tk()
    app = SalesAnalyzerApp(root)
    root.mainloop()