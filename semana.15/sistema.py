"""
Sistema de Exploração dos Microdados PDAD 2024 - Interface Gráfica
Recorte D — Infraestrutura e condições dos domicílios

Pergunta central: como varia o acesso a infraestrutura (água, esgoto,
internet, energia) entre as Regiões Administrativas do DF, e qual a
relação entre o tamanho do domicílio e o tipo de imóvel?

Como executar:
    python sistema.py

Dependências:
    pip install pandas matplotlib openpyxl
"""

import os
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO
# ---------------------------------------------------------------------------
# Os caminhos são montados a partir da pasta onde este arquivo sistema.py
# está salvo (e não da pasta "atual" do terminal), para que o programa
# funcione independentemente de onde o comando python é executado.
PASTA_DO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
CAMINHO_DOMICILIOS = os.path.join(PASTA_DO_SCRIPT, "PDAD_2024-Domicilios.xlsx")
CAMINHO_DICIONARIO = os.path.join(PASTA_DO_SCRIPT, "Dicionario_de_variaveis_PDAD_2024.xlsx")

# Códigos sentinela usados na pesquisa PDAD: 99999 = "não se aplica" e
# 88888 = "não declarado". Precisam ser removidos antes de qualquer
# cálculo estatístico, senão médias e percentuais ficam distorcidos.
SENTINELAS = [99999, 88888]

# Nomes de variáveis já conhecidos, informados no enunciado do trabalho.
COLUNA_LOCALIDADE = "localidade"
COLUNA_N_PESSOAS = "A01npessoas"

# Palavras-chave usadas para localizar automaticamente, no dicionário de
# variáveis, as colunas dos blocos B (características do domicílio) e D
# (infraestrutura) que representam água, esgoto, energia, internet e tipo
# de imóvel. Caso a detecção automática abaixo não encontre alguma coluna
# (o layout exato do dicionário pode variar), preencha manualmente o
# dicionário COLUNAS_MANUAIS com o código exato da coluna, encontrado
# abrindo o arquivo Dicionario_de_variaveis_PDAD_2024.xlsx.
PALAVRAS_CHAVE = {
    "agua": ["abastecimento de agua", "forma de abastecimento de agua", "agua utilizada"],
    "esgoto": ["esgotamento sanitario", "escoadouro", "destino do esgoto"],
    "energia": ["energia eletrica", "iluminacao eletrica"],
    "internet": ["acesso a internet", "possui internet", "internet no domicilio"],
    "tipo_imovel": ["tipo de imovel", "tipo do domicilio", "especie do domicilio", "tipo de domicilio"],
}

# Códigos confirmados no Dicionario_de_variaveis_PDAD_2024.xlsx real:
# B13 = acesso à rede de abastecimento de água (CAESB/Saneago)
# B14 = acesso à rede de coleta de esgoto (CAESB/Saneago)
# B15 = acesso à rede de energia elétrica (Neoenergia/Enel)
# C05 = domicílio possuía acesso à internet no último mês
# B02 = tipo do domicílio (Casa/Apartamento/Cômodo)
COLUNAS_MANUAIS = {
    "agua": "B13",
    "esgoto": "B14",
    "energia": "B15",
    "internet": "C05",
    "tipo_imovel": "B02",
}


# ---------------------------------------------------------------------------
# FUNÇÕES DE CARGA E PREPARO DOS DADOS
# ---------------------------------------------------------------------------
def normalizar_texto(texto):
    """Remove acentos e caixa alta para facilitar a comparação de palavras-chave."""
    if not isinstance(texto, str):
        return ""
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return sem_acento.lower().strip()


def carregar_dicionario(caminho):
    """Lê o dicionário de variáveis e devolve um DataFrame com colunas normalizadas."""
    try:
        dic = pd.read_excel(caminho)
    except FileNotFoundError:
        return None

    # O dicionário pode ter nomes de coluna variados; procuramos por colunas
    # que pareçam conter o código da variável e a descrição correspondente.
    col_codigo, col_descricao = None, None
    for coluna in dic.columns:
        nome = normalizar_texto(str(coluna))
        if col_codigo is None and ("variavel" in nome or "codigo" in nome or "campo" in nome):
            col_codigo = coluna
        if col_descricao is None and ("descri" in nome or "pergunta" in nome or "rotulo" in nome):
            col_descricao = coluna

    if col_codigo is None or col_descricao is None:
        return None

    dic = dic[[col_codigo, col_descricao]].dropna()
    dic.columns = ["codigo", "descricao"]
    dic["descricao_norm"] = dic["descricao"].apply(normalizar_texto)
    return dic


def carregar_categorias(caminho):
    """Lê do dicionário o significado de cada código de categoria (ex.: 1 = Sim, 2 = Não)."""
    try:
        dic = pd.read_excel(caminho)
    except FileNotFoundError:
        return {}

    col_codigo, col_valor, col_desc_valor = None, None, None
    for coluna in dic.columns:
        nome = normalizar_texto(str(coluna))
        if col_codigo is None and ("variavel" in nome or "codigo" in nome or "coluna" in nome):
            col_codigo = coluna
        if col_valor is None and nome == "valor":
            col_valor = coluna
        if col_desc_valor is None and "descri" in nome and "valor" in nome:
            col_desc_valor = coluna

    if col_codigo is None or col_valor is None or col_desc_valor is None:
        return {}

    # O nome da variável só aparece na primeira linha de cada bloco (células
    # mescladas no Excel); preenchemos as linhas seguintes com o mesmo código.
    dic[col_codigo] = dic[col_codigo].ffill()

    categorias = {}
    for _, linha in dic.iterrows():
        valor = linha[col_valor]
        descricao = linha[col_desc_valor]
        if pd.isna(valor) or pd.isna(descricao) or str(valor).strip() == "-":
            continue
        codigo = str(linha[col_codigo]).strip()
        categorias.setdefault(codigo, {})
        # Guardamos a mesma descrição sob a forma original, inteira e texto,
        # para casar com o valor independentemente do tipo lido do Excel.
        categorias[codigo][valor] = descricao
        try:
            categorias[codigo][int(valor)] = descricao
        except (ValueError, TypeError):
            pass
        categorias[codigo][str(valor)] = descricao

    return categorias


def rotular_categoria(categorias, coluna, valor):
    """Traduz um código de categoria (ex.: 1) para um texto legível (ex.: '1 - Sim')."""
    descricao = categorias.get(coluna, {}).get(valor)
    if descricao is None:
        return str(valor)
    return f"{valor} - {descricao}"


def detectar_coluna(dic, palavras_chave, colunas_disponiveis):
    """Procura no dicionário a coluna cuja descrição contém alguma palavra-chave."""
    if dic is None:
        return None
    for palavra in palavras_chave:
        alvo = normalizar_texto(palavra)
        encontrados = dic[dic["descricao_norm"].str.contains(alvo, na=False)]
        for codigo in encontrados["codigo"]:
            codigo = str(codigo).strip()
            if codigo in colunas_disponiveis:
                return codigo
    return None


def montar_mapa_de_colunas(df_domicilios):
    """Combina detecção automática (via dicionário) e overrides manuais em um único mapa."""
    dic = carregar_dicionario(CAMINHO_DICIONARIO)
    colunas_disponiveis = set(df_domicilios.columns.astype(str))

    mapa = {"localidade": COLUNA_LOCALIDADE, "n_pessoas": COLUNA_N_PESSOAS}
    for chave, palavras in PALAVRAS_CHAVE.items():
        if chave in COLUNAS_MANUAIS:
            mapa[chave] = COLUNAS_MANUAIS[chave]
        else:
            mapa[chave] = detectar_coluna(dic, palavras, colunas_disponiveis)
    return mapa


def carregar_domicilios():
    """Lê a tabela de domicílios do Excel e remove linhas com valores sentinela nas colunas usadas."""
    df = pd.read_excel(CAMINHO_DOMICILIOS)
    mapa = montar_mapa_de_colunas(df)
    categorias = carregar_categorias(CAMINHO_DICIONARIO)

    # Filtragem dos valores sentinela (99999 = não se aplica, 88888 = não
    # declarado) apenas nas colunas que o sistema efetivamente utiliza.
    for chave, coluna in mapa.items():
        if coluna is not None and coluna in df.columns:
            df = df[~df[coluna].isin(SENTINELAS)]

    return df, mapa, categorias


# ---------------------------------------------------------------------------
# FUNÇÕES DE ANÁLISE
# ---------------------------------------------------------------------------
def calcular_estatisticas_infra(df, coluna_indicador, coluna_categoria_valor):
    """Calcula total de domicílios, percentual da categoria escolhida e contagem de categorias."""
    total = len(df)
    if total == 0 or coluna_indicador not in df.columns:
        return {"total": 0, "percentual": 0.0, "n_categoria": 0}

    n_categoria = (df[coluna_indicador] == coluna_categoria_valor).sum()
    percentual = (n_categoria / total) * 100 if total else 0.0
    return {"total": total, "percentual": percentual, "n_categoria": n_categoria}


def calcular_estatisticas_tamanho(df, coluna_n_pessoas):
    """Calcula média, mediana e moda do número de pessoas por domicílio no filtro atual."""
    serie = df[coluna_n_pessoas].dropna()
    if serie.empty:
        return {"media": 0.0, "mediana": 0.0, "moda": 0, "total": 0}

    moda_serie = serie.mode()
    moda = moda_serie.iloc[0] if not moda_serie.empty else 0
    return {
        "media": serie.mean(),
        "mediana": serie.median(),
        "moda": moda,
        "total": len(serie),
    }


def exportar_dados(df, caminho):
    """Exporta o DataFrame filtrado para .csv ou .txt, conforme a extensão escolhida."""
    if caminho.lower().endswith(".txt"):
        with open(caminho, "w", encoding="utf-8") as arquivo:
            arquivo.write(df.to_string(index=False))
    else:
        df.to_csv(caminho, index=False, encoding="utf-8-sig")


# ---------------------------------------------------------------------------
# INTERFACE GRÁFICA
# ---------------------------------------------------------------------------
class SistemaInfraestrutura(tk.Tk):
    """Janela principal do sistema, com abas para infraestrutura e tamanho dos domicílios."""

    def __init__(self, df, mapa, categorias):
        super().__init__()
        self.df = df
        self.mapa = mapa
        self.categorias = categorias

        self.title("PDAD 2024 - Infraestrutura e Condições dos Domicílios (Recorte D)")
        self.geometry("980x700")

        self._construir_cabecalho()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.aba_infra = ttk.Frame(notebook)
        self.aba_tamanho = ttk.Frame(notebook)
        notebook.add(self.aba_infra, text="Infraestrutura por RA")
        notebook.add(self.aba_tamanho, text="Tamanho x Tipo de Imóvel")

        self._construir_aba_infraestrutura()
        self._construir_aba_tamanho()

    # -- Cabeçalho ----------------------------------------------------------
    def _construir_cabecalho(self):
        """Monta o título, a descrição do recorte e a contagem de registros carregados."""
        frame = tk.Frame(self)
        frame.pack(fill="x", padx=10, pady=(10, 0))

        tk.Label(
            frame,
            text="Infraestrutura e Condições dos Domicílios do DF",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")

        tk.Label(
            frame,
            text="Explore o acesso a água, esgoto, energia e internet por Região "
                 "Administrativa, e a relação entre tamanho do domicílio e tipo de imóvel.",
            wraplength=900,
            justify="left",
        ).pack(anchor="w")

        tk.Label(
            frame,
            text=f"{len(self.df):,}".replace(",", ".") + " domicílios carregados",
            font=("Segoe UI", 9, "italic"),
            fg="gray30",
        ).pack(anchor="w", pady=(2, 6))

    # -- Aba 1: Infraestrutura por RA ---------------------------------------
    def _construir_aba_infraestrutura(self):
        """Cria os widgets de filtro, gráfico e estatísticas da aba de infraestrutura."""
        aba = self.aba_infra
        indicadores_disponiveis = {
            chave: self.mapa.get(chave)
            for chave in ["agua", "esgoto", "energia", "internet"]
            if self.mapa.get(chave) is not None
        }

        controles = tk.Frame(aba)
        controles.pack(fill="x", pady=6)

        tk.Label(controles, text="Indicador:").grid(row=0, column=0, padx=4, sticky="w")
        self.combo_indicador = ttk.Combobox(
            controles, values=list(indicadores_disponiveis.keys()), state="readonly", width=15
        )
        self.combo_indicador.grid(row=0, column=1, padx=4)
        self.combo_indicador.bind("<<ComboboxSelected>>", lambda e: self._atualizar_categorias())

        tk.Label(controles, text="Categoria:").grid(row=0, column=2, padx=4, sticky="w")
        self.combo_categoria = ttk.Combobox(controles, state="readonly", width=15)
        self.combo_categoria.grid(row=0, column=3, padx=4)

        tk.Label(controles, text="RA:").grid(row=0, column=4, padx=4, sticky="w")
        ras = ["Todas as RAs"] + sorted(self.df[self.mapa["localidade"]].dropna().unique().astype(str))
        self.combo_ra = ttk.Combobox(controles, values=ras, state="readonly", width=22)
        self.combo_ra.current(0)
        self.combo_ra.grid(row=0, column=5, padx=4)

        tk.Button(controles, text="Atualizar", command=self._atualizar_infraestrutura).grid(
            row=0, column=6, padx=8
        )
        tk.Button(controles, text="Exportar filtro", command=self._exportar_infraestrutura).grid(
            row=0, column=7, padx=4
        )

        self.figura_infra = plt.Figure(figsize=(7.5, 4.2), dpi=100)
        self.eixo_infra = self.figura_infra.add_subplot(111)
        self.canvas_infra = FigureCanvasTkAgg(self.figura_infra, master=aba)
        self.canvas_infra.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=6)

        self.label_stats_infra = tk.Label(aba, text="", justify="left", anchor="w")
        self.label_stats_infra.pack(fill="x", padx=6, pady=(0, 6))

        if indicadores_disponiveis:
            self.combo_indicador.current(0)
            self._atualizar_categorias()
        else:
            messagebox.showwarning(
                "Colunas não encontradas",
                "Não foi possível identificar automaticamente as colunas de água, esgoto, "
                "energia ou internet. Preencha o dicionário COLUNAS_MANUAIS no início do "
                "arquivo sistema.py com os códigos corretos, consultando o "
                "Dicionario_de_variaveis_PDAD_2024.xlsx.",
            )

    def _atualizar_categorias(self):
        """Repopula a lista de categorias disponíveis (com rótulos legíveis) conforme o indicador escolhido."""
        chave = self.combo_indicador.get()
        coluna = self.mapa.get(chave)
        if coluna is None or coluna not in self.df.columns:
            return
        valores = sorted(self.df[coluna].dropna().unique().tolist(), key=str)

        # Mapa rótulo -> valor original, para conseguirmos voltar ao código
        # bruto quando o usuário escolher um item já traduzido no combobox.
        self.categoria_valores = {}
        rotulos = []
        for valor in valores:
            rotulo = rotular_categoria(self.categorias, coluna, valor)
            self.categoria_valores[rotulo] = valor
            rotulos.append(rotulo)

        self.combo_categoria["values"] = rotulos
        if rotulos:
            self.combo_categoria.current(0)
        self._atualizar_infraestrutura()

    def _dados_filtrados_por_ra(self):
        """Aplica o filtro de RA escolhido pelo usuário e devolve o DataFrame resultante."""
        ra = self.combo_ra.get()
        if ra and ra != "Todas as RAs":
            return self.df[self.df[self.mapa["localidade"]].astype(str) == ra]
        return self.df

    def _atualizar_infraestrutura(self):
        """Recalcula o gráfico e as estatísticas da aba de infraestrutura com o filtro atual."""
        chave = self.combo_indicador.get()
        coluna = self.mapa.get(chave)
        rotulo_categoria_escolhida = self.combo_categoria.get()
        if coluna is None or not rotulo_categoria_escolhida:
            return

        # O combobox mostra o rótulo traduzido (ex.: "1 - Sim"); recuperamos
        # o valor original correspondente para comparar com os dados.
        categoria_valor = self.categoria_valores.get(rotulo_categoria_escolhida, rotulo_categoria_escolhida)

        df_ra = self._dados_filtrados_por_ra()
        ra_selecionada = self.combo_ra.get()

        self.eixo_infra.clear()

        if ra_selecionada == "Todas as RAs":
            # Compara o percentual da categoria escolhida entre todas as RAs.
            grupo = self.df.groupby(self.mapa["localidade"])[coluna].apply(
                lambda serie: (serie == categoria_valor).mean() * 100
            ).sort_values(ascending=False)
            grupo.plot(kind="bar", ax=self.eixo_infra, color="#4C72B0")
            self.eixo_infra.set_ylabel("% de domicílios na categoria")
            self.eixo_infra.set_xlabel("Região Administrativa")
            self.eixo_infra.set_title(f"{chave} — {rotulo_categoria_escolhida} — por RA")
        else:
            # Mostra a distribuição de todas as categorias dentro da RA escolhida,
            # já traduzindo cada código para seu rótulo legível no eixo X.
            distrib = df_ra[coluna].value_counts(normalize=True).sort_index() * 100
            rotulos_x = [rotular_categoria(self.categorias, coluna, valor) for valor in distrib.index]
            self.eixo_infra.bar(rotulos_x, distrib.values, color="#55A868")
            self.eixo_infra.tick_params(axis="x", rotation=30)
            self.eixo_infra.set_ylabel("% de domicílios")
            self.eixo_infra.set_xlabel(f"Categorias de {chave}")
            self.eixo_infra.set_title(f"Distribuição de {chave} em {ra_selecionada}")

        self.figura_infra.tight_layout()
        self.canvas_infra.draw()

        stats = calcular_estatisticas_infra(df_ra, coluna, categoria_valor)
        self.label_stats_infra.config(
            text=(
                f"Domicílios no filtro atual: {stats['total']}   |   "
                f"Na categoria selecionada: {stats['n_categoria']} "
                f"({stats['percentual']:.1f}%)"
            )
        )
        self._df_filtrado_infra = df_ra

    def _exportar_infraestrutura(self):
        """Abre o diálogo de salvar arquivo e exporta os dados filtrados desta aba."""
        df_export = getattr(self, "_df_filtrado_infra", self.df)
        caminho = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Texto", "*.txt")],
            title="Exportar dados filtrados",
        )
        if caminho:
            exportar_dados(df_export, caminho)
            messagebox.showinfo("Exportação concluída", f"Dados salvos em:\n{caminho}")

    # -- Aba 2: Tamanho x Tipo de Imóvel ------------------------------------
    def _construir_aba_tamanho(self):
        """Cria os widgets de filtro, histograma e estatísticas da aba de tamanho dos domicílios."""
        aba = self.aba_tamanho
        coluna_tipo = self.mapa.get("tipo_imovel")

        controles = tk.Frame(aba)
        controles.pack(fill="x", pady=6)

        tk.Label(controles, text="Tipo de imóvel:").grid(row=0, column=0, padx=4, sticky="w")
        self.tipo_valores = {}
        if coluna_tipo is not None and coluna_tipo in self.df.columns:
            valores_tipo = sorted(self.df[coluna_tipo].dropna().unique().tolist(), key=str)
            rotulos_tipo = []
            for valor in valores_tipo:
                rotulo = rotular_categoria(self.categorias, coluna_tipo, valor)
                self.tipo_valores[rotulo] = valor
                rotulos_tipo.append(rotulo)
            tipos = ["Todos"] + rotulos_tipo
        else:
            tipos = ["Todos"]
        self.combo_tipo = ttk.Combobox(controles, values=tipos, state="readonly", width=20)
        self.combo_tipo.current(0)
        self.combo_tipo.grid(row=0, column=1, padx=4)

        tk.Button(controles, text="Atualizar", command=self._atualizar_tamanho).grid(
            row=0, column=2, padx=8
        )
        tk.Button(controles, text="Exportar filtro", command=self._exportar_tamanho).grid(
            row=0, column=3, padx=4
        )

        self.figura_tam = plt.Figure(figsize=(7.5, 4.2), dpi=100)
        self.eixo_tam = self.figura_tam.add_subplot(111)
        self.canvas_tam = FigureCanvasTkAgg(self.figura_tam, master=aba)
        self.canvas_tam.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=6)

        self.label_stats_tam = tk.Label(aba, text="", justify="left", anchor="w")
        self.label_stats_tam.pack(fill="x", padx=6, pady=(0, 6))

        if coluna_tipo is None:
            messagebox.showwarning(
                "Coluna não encontrada",
                "Não foi possível identificar automaticamente a coluna de tipo de imóvel. "
                "Preencha COLUNAS_MANUAIS['tipo_imovel'] no início do sistema.py.",
            )

        self._atualizar_tamanho()

    def _atualizar_tamanho(self):
        """Recalcula o histograma e as estatísticas de tamanho do domicílio com o filtro atual."""
        coluna_tipo = self.mapa.get("tipo_imovel")
        coluna_n = self.mapa["n_pessoas"]
        rotulo_tipo = self.combo_tipo.get()

        if coluna_tipo is not None and rotulo_tipo != "Todos":
            valor_tipo = self.tipo_valores.get(rotulo_tipo, rotulo_tipo)
            df_filtro = self.df[self.df[coluna_tipo] == valor_tipo]
        else:
            df_filtro = self.df

        self.eixo_tam.clear()
        self.eixo_tam.hist(
            df_filtro[coluna_n].dropna(), bins=range(1, 12), color="#C44E52", edgecolor="white"
        )
        self.eixo_tam.set_xlabel("Número de pessoas no domicílio")
        self.eixo_tam.set_ylabel("Quantidade de domicílios")
        self.eixo_tam.set_title(f"Tamanho do domicílio — {rotulo_tipo}")
        self.figura_tam.tight_layout()
        self.canvas_tam.draw()

        stats = calcular_estatisticas_tamanho(df_filtro, coluna_n)
        self.label_stats_tam.config(
            text=(
                f"Domicílios no filtro: {stats['total']}   |   "
                f"Média de pessoas: {stats['media']:.2f}   |   "
                f"Mediana: {stats['mediana']:.1f}   |   "
                f"Moda: {stats['moda']}"
            )
        )
        self._df_filtrado_tam = df_filtro

    def _exportar_tamanho(self):
        """Abre o diálogo de salvar arquivo e exporta os dados filtrados desta aba."""
        df_export = getattr(self, "_df_filtrado_tam", self.df)
        caminho = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Texto", "*.txt")],
            title="Exportar dados filtrados",
        )
        if caminho:
            exportar_dados(df_export, caminho)
            messagebox.showinfo("Exportação concluída", f"Dados salvos em:\n{caminho}")


# ---------------------------------------------------------------------------
# PONTO DE ENTRADA
# ---------------------------------------------------------------------------
def main():
    """Carrega os dados e inicia a interface gráfica do sistema."""
    try:
        df, mapa, categorias = carregar_domicilios()
    except FileNotFoundError as erro:
        janela_erro = tk.Tk()
        janela_erro.withdraw()
        messagebox.showerror(
            "Arquivo não encontrado",
            f"Não foi possível abrir o arquivo de dados:\n{erro}\n\n"
            "Verifique se PDAD_2024-Domicilios.xlsx e "
            "Dicionario_de_variaveis_PDAD_2024.xlsx estão na mesma pasta do sistema.py.",
        )
        return

    app = SistemaInfraestrutura(df, mapa, categorias)
    app.mainloop()


if __name__ == "__main__":
    main()
