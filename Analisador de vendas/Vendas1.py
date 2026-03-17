import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def obter_pasta_base() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def obter_caminho_recurso(nome_arquivo: str) -> Path:
    """
    Retorna o caminho correto de um recurso tanto no modo .py
    quanto no executável gerado pelo PyInstaller.
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / nome_arquivo
    return obter_pasta_base() / nome_arquivo


def abrir_arquivo(caminho: Path) -> None:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    if sys.platform.startswith("win"):
        os.startfile(caminho)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(caminho)], check=True)
    else:
        subprocess.run(["xdg-open", str(caminho)], check=True)


def processar_planilhas(
    input_dir: Path,
    output_dir: Path,
    atualizar_progresso=None
) -> tuple[Path, int]:
    output_dir.mkdir(parents=True, exist_ok=True)

    arquivos_excel = sorted(input_dir.glob("*.xlsx"))

    if not arquivos_excel:
        raise FileNotFoundError(
            "Nenhum arquivo .xlsx foi encontrado na pasta selecionada."
        )

    lista_dfs = []
    total_arquivos = len(arquivos_excel)

    for indice, arquivo in enumerate(arquivos_excel, start=1):
        df = pd.read_excel(arquivo)
        df["Arquivo_Origem"] = arquivo.name
        lista_dfs.append(df)

        if atualizar_progresso:
            progresso = int((indice / total_arquivos) * 65)
            atualizar_progresso(
                progresso,
                f"Lendo arquivo {indice} de {total_arquivos}: {arquivo.name}"
            )

    df = pd.concat(lista_dfs, ignore_index=True)

    colunas_necessarias = ["Produto", "Quantidade", "Valor_Unitario"]
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            raise ValueError(
                f"A coluna obrigatória '{coluna}' não foi encontrada em uma ou mais planilhas."
            )

    if atualizar_progresso:
        atualizar_progresso(72, "Limpando e validando os dados...")

    df["Produto"] = df["Produto"].astype(str).str.strip()
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce")
    df["Valor_Unitario"] = pd.to_numeric(df["Valor_Unitario"], errors="coerce")

    df = df.dropna(subset=["Produto", "Quantidade", "Valor_Unitario"]).copy()
    df = df[df["Produto"] != ""]
    df = df[df["Quantidade"] >= 0]
    df = df[df["Valor_Unitario"] >= 0]

    if df.empty:
        raise ValueError(
            "Após a limpeza dos dados, nenhuma linha válida restou para análise."
        )

    df["Total"] = df["Quantidade"] * df["Valor_Unitario"]

    if atualizar_progresso:
        atualizar_progresso(80, "Calculando métricas e rankings...")

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

    if atualizar_progresso:
        atualizar_progresso(88, "Gerando relatório Excel...")

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    arquivo_saida = output_dir / f"relatorio_vendas_{timestamp}.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="xlsxwriter") as writer:
        resumo_geral.to_excel(writer, sheet_name="Resumo", index=False)
        ranking_faturamento.to_excel(writer, sheet_name="Ranking_Faturamento")
        ranking_quantidade.to_excel(writer, sheet_name="Ranking_Quantidade")
        vendas_por_arquivo.to_excel(writer, sheet_name="Por_Arquivo")
        df.to_excel(writer, sheet_name="Base_Consolidada", index=False)

        workbook = writer.book

        formato_titulo = workbook.add_format({"bold": True, "font_size": 12})
        formato_moeda = workbook.add_format({"num_format": "R$ #,##0.00"})
        formato_numero = workbook.add_format({"num_format": "0.00"})

        ws_resumo = writer.sheets["Resumo"]
        ws_rank_fat = writer.sheets["Ranking_Faturamento"]
        ws_rank_qtd = writer.sheets["Ranking_Quantidade"]
        ws_arquivo = writer.sheets["Por_Arquivo"]
        ws_base = writer.sheets["Base_Consolidada"]

        ws_resumo.set_column("A:A", 35)
        ws_resumo.set_column("B:B", 20, formato_numero)
        ws_rank_fat.set_column("A:A", 25)
        ws_rank_fat.set_column("B:B", 15)
        ws_rank_fat.set_column("C:C", 18, formato_moeda)

        ws_rank_qtd.set_column("A:A", 25)
        ws_rank_qtd.set_column("B:B", 15)
        ws_rank_qtd.set_column("C:C", 18, formato_moeda)

        ws_arquivo.set_column("A:A", 30)
        ws_arquivo.set_column("B:B", 15)
        ws_arquivo.set_column("C:C", 18, formato_moeda)

        ws_base.set_column("A:A", 25)
        ws_base.set_column("B:B", 15)
        ws_base.set_column("C:C", 18, formato_moeda)
        ws_base.set_column("D:D", 30)
        ws_base.set_column("E:E", 18, formato_moeda)

        ws_resumo.write("D2", "Resumo do Relatório", formato_titulo)
        ws_resumo.write("D3", f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

        chart = workbook.add_chart({"type": "column"})
        chart.add_series({
            "name": "Faturamento por Produto",
            "categories": ["Ranking_Faturamento", 1, 0, len(ranking_faturamento), 0],
            "values": ["Ranking_Faturamento", 1, 2, len(ranking_faturamento), 2],
        })
        chart.set_title({"name": "Faturamento por Produto"})
        chart.set_x_axis({"name": "Produto"})
        chart.set_y_axis({"name": "Valor Vendido"})
        ws_rank_fat.insert_chart("E2", chart)

    if atualizar_progresso:
        atualizar_progresso(100, "Relatório finalizado com sucesso.")

    return arquivo_saida, len(arquivos_excel)


class SalesAnalyzerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Analisador de Planilhas de Vendas")
        self.root.geometry("780x500")
        self.root.resizable(False, False)

        self.base_dir = obter_pasta_base()
        self.output_dir = self.base_dir / "Relatório"
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.caminho_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(
            value="Selecione a pasta que contém os arquivos Excel no formato .xlsx."
        )
        self.arquivos_var = tk.StringVar(value="Arquivos .xlsx encontrados: 0")
        self.ultimo_relatorio: Path | None = None

        self.definir_icone()
        self.criar_interface()

    def definir_icone(self):
        try:
            icone_path = obter_caminho_recurso("icon.ico")
            if icone_path.exists():
                self.root.iconbitmap(str(icone_path))
        except Exception:
            pass

    def criar_interface(self):
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        titulo = tk.Label(
            frame,
            text="Analisador de Planilhas de Vendas",
            font=("Arial", 16, "bold")
        )
        titulo.pack(pady=(0, 10))

        instrucoes = tk.Label(
            frame,
            text=(
                "Selecione a pasta onde estão os arquivos Excel no formato .xlsx.\n"
                "O programa irá analisar todos os arquivos .xlsx encontrados nessa pasta e salvar\n"
                "o relatório final automaticamente em uma pasta chamada 'Relatório', ao lado do executável."
            ),
            justify="left",
            wraplength=730,
            anchor="w"
        )
        instrucoes.pack(fill="x", pady=(0, 15))

        label_pasta = tk.Label(frame, text="Pasta com os arquivos .xlsx:")
        label_pasta.pack(anchor="w")

        entrada_frame = tk.Frame(frame)
        entrada_frame.pack(fill="x", pady=5)

        entrada = tk.Entry(
            entrada_frame,
            textvariable=self.caminho_var,
            width=82
        )
        entrada.pack(side="left", fill="x", expand=True)

        botao_procurar = tk.Button(
            entrada_frame,
            text="Selecionar Pasta",
            command=self.selecionar_pasta
        )
        botao_procurar.pack(side="left", padx=(10, 0))

        info_arquivos = tk.Label(
            frame,
            textvariable=self.arquivos_var,
            anchor="w",
            font=("Arial", 10, "italic")
        )
        info_arquivos.pack(anchor="w", pady=(10, 0))

        botoes_frame = tk.Frame(frame)
        botoes_frame.pack(pady=20)

        self.botao_gerar = tk.Button(
            botoes_frame,
            text="Gerar Relatório",
            width=20,
            height=2,
            command=self.gerar_relatorio
        )
        self.botao_gerar.pack(side="left", padx=5)

        botao_abrir = tk.Button(
            botoes_frame,
            text="Abrir Relatório",
            width=20,
            height=2,
            command=self.abrir_relatorio
        )
        botao_abrir.pack(side="left", padx=5)

        botao_atualizar = tk.Button(
            botoes_frame,
            text="Atualizar Contagem",
            width=20,
            height=2,
            command=self.atualizar_contagem_arquivos
        )
        botao_atualizar.pack(side="left", padx=5)

        progresso_titulo = tk.Label(frame, text="Progresso:")
        progresso_titulo.pack(anchor="w", pady=(5, 0))

        self.barra_progresso = ttk.Progressbar(
            frame,
            orient="horizontal",
            length=730,
            mode="determinate",
            maximum=100
        )
        self.barra_progresso.pack(fill="x", pady=(5, 10))

        status_titulo = tk.Label(frame, text="Status:")
        status_titulo.pack(anchor="w")

        status_label = tk.Label(
            frame,
            textvariable=self.status_var,
            justify="left",
            wraplength=730,
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
        caminho_texto = self.caminho_var.get().strip()

        if not caminho_texto:
            self.arquivos_var.set("Arquivos .xlsx encontrados: 0")
            self.status_var.set(
                "Selecione a pasta que contém os arquivos Excel no formato .xlsx."
            )
            self.atualizar_barra(0)
            return

        pasta = Path(caminho_texto)
        quantidade = self.contar_arquivos_excel(pasta)
        self.arquivos_var.set(f"Arquivos .xlsx encontrados: {quantidade}")

        if quantidade > 0:
            self.status_var.set(
                f"Foram encontrados {quantidade} arquivo(s) .xlsx na pasta selecionada."
            )
        else:
            self.status_var.set(
                "Nenhum arquivo .xlsx foi encontrado na pasta selecionada."
            )

        self.atualizar_barra(0)

    def atualizar_barra(self, valor: int, mensagem: str | None = None):
        self.barra_progresso["value"] = valor
        if mensagem:
            self.status_var.set(mensagem)
        self.root.update_idletasks()

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(
            title="Selecione a pasta com os arquivos Excel (.xlsx)"
        )
        if pasta:
            self.caminho_var.set(pasta)
            self.atualizar_contagem_arquivos()

    def gerar_relatorio(self):
        try:
            caminho_texto = self.caminho_var.get().strip()

            if not caminho_texto:
                raise ValueError(
                    "Nenhuma pasta foi selecionada. Selecione a pasta onde estão os arquivos .xlsx."
                )

            input_dir = Path(caminho_texto)

            if not input_dir.exists():
                raise FileNotFoundError("A pasta selecionada não existe.")

            if not input_dir.is_dir():
                raise ValueError("O caminho selecionado não é uma pasta válida.")

            self.output_dir.mkdir(parents=True, exist_ok=True)

            self.botao_gerar.config(state="disabled")
            self.atualizar_barra(5, "Iniciando processamento dos arquivos...")

            arquivo_saida, quantidade_arquivos = processar_planilhas(
                input_dir,
                self.output_dir,
                atualizar_progresso=self.atualizar_barra
            )
            self.ultimo_relatorio = arquivo_saida

            mensagem = (
                f"Relatório gerado com sucesso.\n\n"
                f"Arquivos processados: {quantidade_arquivos}\n"
                f"Pasta de saída: {self.output_dir}\n"
                f"Arquivo gerado: {arquivo_saida.name}"
            )
            self.status_var.set(mensagem)
            self.arquivos_var.set(f"Arquivos .xlsx encontrados: {quantidade_arquivos}")
            messagebox.showinfo("Sucesso", mensagem)

        except Exception as e:
            self.atualizar_barra(0)
            self.status_var.set(f"Erro: {e}")
            messagebox.showerror("Erro", str(e))

        finally:
            self.botao_gerar.config(state="normal")

    def abrir_relatorio(self):
        try:
            if self.ultimo_relatorio is None:
                relatorios = sorted(self.output_dir.glob("relatorio_vendas_*.xlsx"))
                if relatorios:
                    self.ultimo_relatorio = relatorios[-1]
                else:
                    raise FileNotFoundError(
                        "Nenhum relatório foi gerado ainda. Gere o relatório primeiro."
                    )

            abrir_arquivo(self.ultimo_relatorio)
            self.status_var.set(f"Relatório aberto com sucesso: {self.ultimo_relatorio}")

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