# -*- coding: utf-8 -*-
"""BI de Falta de Bipagem - acompanhamento diário e mensal.

Como rodar:  streamlit run app.py
Cada linha (ou bloco de linhas) tem um comentário explicando o que faz.
"""

# ======================================================================
# 1. IMPORTAÇÕES (bibliotecas que o programa usa)
# ======================================================================
import io                                    # io: cria "arquivos na memória" (BytesIO) para ler uploads e gerar downloads
from datetime import timedelta               # timedelta: representa "N dias" para somar/subtrair de datas

import numpy as np                           # numpy: usado aqui só para tratar valores infinitos (np.inf)
import pandas as pd                          # pandas: biblioteca de tabelas (DataFrame) - o "Excel do Python"
import plotly.express as px                  # plotly.express: cria gráficos interativos com poucas linhas
import streamlit as st                       # streamlit: transforma este script em um site/BI

# ======================================================================
# 2. CONFIGURAÇÃO DA PÁGINA E CONSTANTES
# ======================================================================
st.set_page_config(                          # define título, ícone e largura da página (precisa ser o 1º comando st)
    page_title="BI - Falta de Bipagem",      # texto que aparece na aba do navegador
    page_icon="📦",                          # ícone da aba
    layout="wide",                           # usa a tela toda (ideal para dashboards)
)
px.defaults.template = "plotly_white"        # todos os gráficos com fundo branco e visual limpo

# Nomes das colunas da aba "Dados" guardados em constantes: se o nome mudar na planilha, você altera só aqui.
COL_REMESSA = "Número da Remessa"            # identificador da remessa
COL_PEDIDOS = "Pedidos"                      # quantidade de pedidos dentro da remessa
COL_DESCARGA = "Horário do bipe de descarga do veículo recebido"  # data/hora em que o veículo foi descarregado
COL_CLIENTE = "Nome do cliente"              # cliente dono da remessa
COL_OPERADOR = "Operador da descarga do veículo recebido"         # quem descarregou
COL_VIAGEM = "ID Viagem do veículo recebido"  # viagem do veículo
COL_ORIGEM = "Última parada"                 # de onde o veículo veio (última base)
COL_SEGMENTO = "3 Segmentos"                 # segmento/código de rota da remessa
COL_SC = "Nome SC"                           # centro de serviço (SC)
COL_REGIONAL = "Nome da regional"            # regional do SC
COL_PROBLEMA = "Tipo pacote problemático"    # nome do problema (ex.: avaria) - veio em chinês na planilha
COL_PROX = "Horário bipe próxima parada"     # bipe na próxima parada - veio em chinês na planilha

# Colunas da aba "Resumo"
R_TOTAL = "Qtd pedidos a bipar"                                   # total de pedidos que deveriam ser bipados
R_REC = "Qtd pedidos não bipados no recebimento"                  # pedidos sem bipe no recebimento
R_EXP = "Qtd pedidos não bipados na expedição"                    # pedidos sem bipe na expedição

# Dicionário para traduzir os cabeçalhos em chinês da planilha original para português
RENOMEAR = {
    "装车发件业务": "Negócio de carregamento",   # 'negócio de carregamento/expedição' (coluna vazia no arquivo)
    "问题件扫描编码": "Código pacote problemático",  # 'código do bipe de pacote problemático'
    "问题件扫描名称": COL_PROBLEMA,              # 'nome do bipe de pacote problemático' (ex.: Avaria)
    "下一站扫描时间": COL_PROX,                  # 'horário do bipe na próxima estação'
}

# Colunas de texto que usamos para filtrar/agrupar; se faltarem no arquivo, criamos com "(não informado)"
DIMENSOES = [COL_CLIENTE, COL_OPERADOR, COL_VIAGEM, COL_ORIGEM, COL_SEGMENTO, COL_SC, COL_REGIONAL]

DIAS_SEMANA = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex", 5: "Sáb", 6: "Dom"}  # 0 = segunda no pandas


# ======================================================================
# 3. FUNÇÕES AUXILIARES
# ======================================================================
def fmt(n):                                  # formata número no padrão brasileiro (milhar com ponto)
    return f"{n:,.0f}".replace(",", ".")     # 12345 -> "12,345" -> "12.345"


def pct(x):                                  # formata proporção como porcentagem brasileira
    return f"{x:.2%}".replace(".", ",")      # 0.0079 -> "0.79%" -> "0,79%"


@st.cache_data(show_spinner="Lendo arquivo...")  # cache: se o mesmo arquivo for reenviado, não lê de novo
def ler_arquivo(conteudo: bytes, nome: str) -> dict:  # devolve {nome_da_aba: tabela}
    ext = nome.lower().rsplit(".", 1)[-1]    # pega a extensão: "Relatorio.XLSX" -> "xlsx"
    buffer = io.BytesIO(conteudo)            # transforma os bytes do upload em um arquivo na memória
    if ext in ("csv", "tsv", "txt"):         # arquivos de texto separados por delimitador
        for encoding in ("utf-8-sig", "latin-1"):  # tenta UTF-8 (com BOM do Excel) e depois Latin-1
            try:                             # tenta ler com essa codificação
                buffer.seek(0)               # volta ao início do arquivo antes de cada tentativa
                tabela = pd.read_csv(buffer, sep=None, engine="python", encoding=encoding)  # sep=None detecta ; , ou tab
                return {"CSV": tabela}       # CSV só tem uma "aba", então devolvemos com o nome "CSV"
            except UnicodeDecodeError:       # se a codificação estiver errada...
                continue                     # ...tenta a próxima
        raise ValueError("Não consegui decodificar o arquivo de texto.")  # nenhuma codificação funcionou
    engine = "odf" if ext == "ods" else None  # .ods (LibreOffice) precisa do motor 'odf'; os demais o pandas escolhe
    return pd.read_excel(buffer, sheet_name=None, engine=engine)  # sheet_name=None lê TODAS as abas de uma vez


def preparar_detalhe(tabela: pd.DataFrame, arquivo: str) -> pd.DataFrame:  # limpa e enriquece a aba de detalhe
    df = tabela.rename(columns=RENOMEAR).copy()  # traduz cabeçalhos chineses e trabalha numa cópia
    df[COL_DESCARGA] = pd.to_datetime(df[COL_DESCARGA], errors="coerce")  # texto -> data/hora; inválidos viram NaT
    df = df.dropna(subset=[COL_DESCARGA])    # sem data de descarga não dá para analisar por dia: descarta
    if COL_PROX not in df:                   # se o arquivo não tem a coluna de próxima parada...
        df[COL_PROX] = pd.NaT                # ...cria vazia para o resto do código não quebrar
    df[COL_PROX] = pd.to_datetime(df[COL_PROX], errors="coerce")  # garante tipo data/hora
    if COL_PROBLEMA not in df:               # mesma ideia para a coluna de pacote problemático
        df[COL_PROBLEMA] = np.nan            # cria vazia
    if COL_PEDIDOS not in df:                # se não houver coluna de pedidos...
        df[COL_PEDIDOS] = 1                  # ...assume 1 pedido por linha
    df[COL_PEDIDOS] = pd.to_numeric(df[COL_PEDIDOS], errors="coerce").fillna(0)  # número; vazio vira 0
    for coluna in DIMENSOES:                 # para cada coluna de texto usada em filtros/rankings
        if coluna not in df:                 # se não existir...
            df[coluna] = "(não informado)"   # ...cria com um valor padrão
        df[coluna] = df[coluna].fillna("(não informado)").astype(str)  # vazios viram texto padrão; tudo vira string
    df["Data"] = df[COL_DESCARGA].dt.normalize()   # só a data (00:00) - chave para o acompanhamento diário
    df["Mês"] = df[COL_DESCARGA].dt.to_period("M").astype(str)  # "2026-09" - chave para o acompanhamento mensal
    df["Hora"] = df[COL_DESCARGA].dt.hour    # hora cheia (0 a 23) - usada no mapa de calor
    df["Dia da semana"] = df[COL_DESCARGA].dt.dayofweek.map(DIAS_SEMANA)  # 0..6 -> Seg..Dom
    df["Arquivo"] = arquivo                  # guarda de qual arquivo veio a linha (útil ao juntar vários)
    return df                                # devolve a tabela pronta


def consolidar_resumo(lista: list) -> pd.DataFrame:  # junta os "Resumo" de vários arquivos
    junto = pd.concat(lista, ignore_index=True)      # empilha todas as tabelas de resumo
    por_sc = junto.groupby(["Código SC", "Nome SC"], as_index=False)[[R_TOTAL, R_REC, R_EXP]].sum()  # soma por SC
    por_sc["Taxa recebimento"] = por_sc[R_REC] / por_sc[R_TOTAL]      # % não bipado no recebimento
    por_sc["Taxa expedição"] = por_sc[R_EXP] / por_sc[R_TOTAL]        # % não bipado na expedição
    por_sc["Taxa geral"] = (por_sc[R_REC] + por_sc[R_EXP]) / (2 * por_sc[R_TOTAL])  # média das duas taxas
    return por_sc                                    # devolve o resumo consolidado


def grafico(fig):                            # atalho para mostrar qualquer gráfico plotly
    st.plotly_chart(fig, width="stretch")    # 'stretch' = ocupa toda a largura disponível


# ======================================================================
# 4. PROGRAMA PRINCIPAL (a interface)
# ======================================================================
def main():                                  # tudo que aparece na tela fica dentro desta função
    st.title("📦 BI de Falta de Bipagem")    # título da página
    st.caption("Acompanhamento diário e mensal de pedidos sem bipe no recebimento e na expedição.")  # subtítulo

    # ---------- 4.1 Upload ----------
    with st.sidebar:                         # tudo aqui dentro aparece na barra lateral
        st.header("1. Carregar planilhas")   # título da seção
        arquivos = st.file_uploader(         # botão/área de upload
            "CSV, XLS, XLSX, XLSM, ODS (pode enviar vários)",  # texto de ajuda
            type=["csv", "tsv", "txt", "xls", "xlsx", "xlsm", "ods"],  # extensões permitidas
            accept_multiple_files=True,      # permite enviar vários relatórios de uma vez (ex.: um por dia)
        )
        remover_dup = st.checkbox(           # opção para evitar contar duas vezes o mesmo registro
            "Remover linhas idênticas entre arquivos", value=True,
            help="Útil quando relatórios diários se sobrepõem.",
        )

    if not arquivos:                         # se ninguém enviou nada ainda...
        st.info("⬅️ Envie o relatório na barra lateral para começar.")  # ...mostra instrução
        st.stop()                            # ...e interrompe o script aqui

    # ---------- 4.2 Leitura e classificação das abas ----------
    detalhes, resumos = [], []               # listas onde vamos guardar as tabelas encontradas
    for arq in arquivos:                     # percorre cada arquivo enviado
        abas = ler_arquivo(arq.getvalue(), arq.name)  # lê o arquivo (todas as abas)
        for nome_aba, tabela in abas.items():         # percorre cada aba
            tabela = tabela.rename(columns=lambda c: str(c).strip())  # tira espaços extras dos cabeçalhos
            if COL_REMESSA in tabela.columns:         # tem "Número da Remessa"? então é a aba de detalhe
                detalhes.append(preparar_detalhe(tabela, arq.name))
            elif R_TOTAL in tabela.columns:           # tem "Qtd pedidos a bipar"? então é a aba Resumo
                resumos.append(tabela)

    if not detalhes:                         # nenhuma tabela de detalhe encontrada
        st.error(f"Não encontrei a coluna '{COL_REMESSA}' em nenhum arquivo/aba.")  # avisa o usuário
        st.stop()                            # interrompe

    dados = pd.concat(detalhes, ignore_index=True)    # junta todos os detalhes em uma tabela só
    if remover_dup:                          # se a opção estiver marcada...
        colunas_chave = [c for c in dados.columns if c != "Arquivo"]  # compara todas as colunas menos o nome do arquivo
        dados = dados.drop_duplicates(subset=colunas_chave)           # apaga linhas 100% iguais

    # ---------- 4.3 Filtros ----------
    with st.sidebar:                         # volta à barra lateral
        st.header("2. Filtros")              # título da seção
        dmin, dmax = dados["Data"].min().date(), dados["Data"].max().date()  # menor e maior data dos dados
        periodo = st.date_input(             # seletor de intervalo de datas
            "Período (data da descarga)", value=(dmin, dmax), min_value=dmin, max_value=dmax,
        )
        if len(periodo) != 2:                # enquanto o usuário escolheu só a 1ª data do intervalo...
            st.stop()                        # ...espera a segunda
        sel_sc = st.multiselect("SC", sorted(dados[COL_SC].unique()))                # filtro de SC
        sel_reg = st.multiselect("Regional", sorted(dados[COL_REGIONAL].unique()))   # filtro de regional
        sel_cli = st.multiselect("Cliente", sorted(dados[COL_CLIENTE].unique()))     # filtro de cliente
        sel_ori = st.multiselect("Origem (última parada)", sorted(dados[COL_ORIGEM].unique()))  # filtro de origem

    mascara = dados["Data"].between(pd.Timestamp(periodo[0]), pd.Timestamp(periodo[1]))  # True nas linhas dentro do período
    for coluna, escolhidos in [(COL_SC, sel_sc), (COL_REGIONAL, sel_reg), (COL_CLIENTE, sel_cli), (COL_ORIGEM, sel_ori)]:
        if escolhidos:                       # só filtra se o usuário escolheu algo naquele filtro
            mascara &= dados[coluna].isin(escolhidos)  # mantém apenas as linhas com os valores escolhidos
    df = dados[mascara]                      # tabela final filtrada, usada em todas as abas
    if df.empty:                             # se o filtro não deixou nenhuma linha...
        st.warning("Nenhum registro para os filtros escolhidos.")  # ...avisa
        st.stop()                            # ...e interrompe

    # ---------- 4.4 Indicadores (KPIs) ----------
    st.subheader("Indicadores do período")   # título da seção
    k = st.columns(6)                        # cria 6 colunas lado a lado
    k[0].metric("Registros sem bipe de expedição", fmt(len(df)))       # nº de linhas = registros da aba Dados
    k[1].metric("Remessas únicas", fmt(df[COL_REMESSA].nunique()))     # remessas distintas (uma remessa pode repetir)
    k[2].metric("Pedidos (soma)", fmt(df[COL_PEDIDOS].sum()))          # soma da coluna Pedidos
    k[3].metric("Viagens", fmt(df[COL_VIAGEM].nunique()))              # viagens distintas
    k[4].metric("Clientes", fmt(df[COL_CLIENTE].nunique()))            # clientes distintos
    k[5].metric("Com bipe na próxima parada", fmt(df[COL_PROX].notna().sum()))  # linhas que têm data em "próxima parada"

    resumo = None                            # começa sem resumo
    if resumos:                              # se algum arquivo trouxe a aba Resumo...
        resumo = consolidar_resumo(resumos)  # consolida
        if sel_sc:                           # respeita o filtro de SC, se houver
            resumo = resumo[resumo["Nome SC"].isin(sel_sc)]
        tot, rec, exp = resumo[R_TOTAL].sum(), resumo[R_REC].sum(), resumo[R_EXP].sum()  # totais
        if tot > 0:                          # evita divisão por zero
            r = st.columns(4)                # 4 colunas para as taxas
            r[0].metric("Pedidos a bipar (Resumo)", fmt(tot))                        # denominador das taxas
            r[1].metric("Taxa não bipados no recebimento", pct(rec / tot))           # rec / total
            r[2].metric("Taxa não bipados na expedição", pct(exp / tot))             # exp / total
            r[3].metric("Taxa geral de falta de bipagem", pct((rec + exp) / (2 * tot)))  # média das duas taxas
            st.caption("Taxas vêm da aba Resumo (período completo do relatório) e não mudam com o filtro de datas.")

    # ---------- 4.5 Abas ----------
    aba_dia, aba_mes, aba_ofe, aba_hor, aba_dad = st.tabs(  # cria 5 abas na página
        ["📅 Diário", "🗓️ Mensal", "🎯 Ofensores", "🕒 Horários", "📄 Dados"]
    )

    # ===== Aba Diário =====
    with aba_dia:                            # tudo aqui aparece na aba "Diário"
        diario = df.groupby("Data").agg(     # agrupa por dia
            Registros=(COL_REMESSA, "size"), Pedidos=(COL_PEDIDOS, "sum"),  # conta linhas e soma pedidos
        )
        calendario = pd.date_range(diario.index.min(), diario.index.max(), freq="D")  # todos os dias do período
        diario = diario.reindex(calendario, fill_value=0).rename_axis("Data").reset_index()  # dias sem registro = 0
        diario["Média móvel 7d"] = diario["Registros"].rolling(7, min_periods=1).mean().round(1)  # tendência de 7 dias
        variacao = diario["Registros"].pct_change().replace([np.inf, -np.inf], np.nan)  # variação vs dia anterior
        diario["Var. % vs dia anterior"] = (variacao * 100).round(1)  # em porcentagem

        fig = px.bar(diario, x="Data", y="Registros", title="Registros sem bipe de expedição por dia")  # barras por dia
        fig.add_scatter(x=diario["Data"], y=diario["Média móvel 7d"], mode="lines", name="Média móvel 7d")  # linha de tendência
        grafico(fig)                         # mostra o gráfico

        st.subheader("Detalhe de um dia")    # seção de acompanhamento diário
        dias = list(diario["Data"].dt.date.iloc[::-1])       # lista de dias, do mais recente para o mais antigo
        dia = st.selectbox("Escolha o dia", dias)            # caixa de seleção
        serie = diario.set_index(diario["Data"].dt.date)["Registros"]  # registros por dia, indexado por data
        atual = int(serie[dia])              # registros do dia escolhido
        anterior = serie.get(dia - timedelta(days=1))        # registros do dia anterior (None se não existir)
        d1, d2 = st.columns(2)               # duas colunas lado a lado
        d1.metric(                           # cartão com o valor do dia
            "Registros no dia", fmt(atual),
            delta=None if anterior is None else int(atual - anterior),  # diferença vs dia anterior
            delta_color="inverse",           # inverso: aumentar é ruim (vermelho), diminuir é bom (verde)
        )
        d2.metric("Pedidos no dia", fmt(diario.loc[diario["Data"].dt.date == dia, "Pedidos"].sum()))  # soma de pedidos
        do_dia = df[df["Data"].dt.date == dia]               # linhas somente do dia escolhido
        top_op = do_dia[COL_OPERADOR].value_counts().head(10).rename_axis("Operador").reset_index(name="Registros")
        fig = px.bar(top_op, x="Registros", y="Operador", orientation="h", title="Top 10 operadores no dia")
        fig.update_layout(yaxis={"categoryorder": "total ascending"})  # maior barra no topo
        grafico(fig)                         # mostra o gráfico

        st.subheader("Idade dos registros (em relação ao último dia do período)")  # backlog por antiguidade
        idade = (df["Data"].max() - df["Data"]).dt.days      # dias entre a descarga e o último dia do período
        faixas = pd.cut(                     # separa em faixas
            idade, bins=[-1, 0, 2, 6, 13, 10**6],
            labels=["Hoje (0d)", "1-2 dias", "3-6 dias", "7-13 dias", "14+ dias"],
        )
        idade_df = faixas.value_counts(sort=False).rename_axis("Faixa").reset_index(name="Registros")  # conta por faixa
        grafico(px.bar(idade_df, x="Faixa", y="Registros", text="Registros"))  # barras com o valor escrito

        st.dataframe(diario, width="stretch", hide_index=True)  # tabela diária completa

    # ===== Aba Mensal =====
    with aba_mes:                            # tudo aqui aparece na aba "Mensal"
        mensal = df.groupby("Mês").agg(      # agrupa por mês
            Registros=(COL_REMESSA, "size"), Pedidos=(COL_PEDIDOS, "sum"), Dias_com_registro=("Data", "nunique"),
        ).reset_index()
        mensal["Média por dia"] = (mensal["Registros"] / mensal["Dias_com_registro"]).round(1)  # compara meses parciais
        mensal["Var. % vs mês anterior"] = (mensal["Registros"].pct_change() * 100).round(1)   # variação mensal
        fig = px.bar(mensal, x="Mês", y="Registros", text="Registros", title="Registros por mês")  # barras por mês
        grafico(fig)                         # mostra o gráfico

        top_cli = df[COL_CLIENTE].value_counts().head(8).index      # 8 maiores clientes
        rotulo = df[COL_CLIENTE].where(df[COL_CLIENTE].isin(top_cli), "Outros")  # demais viram "Outros"
        por_cli = df.assign(Cliente=rotulo).groupby(["Mês", "Cliente"]).size().reset_index(name="Registros")
        grafico(px.bar(por_cli, x="Mês", y="Registros", color="Cliente", title="Mês x cliente"))  # barras empilhadas
        st.dataframe(mensal, width="stretch", hide_index=True)     # tabela mensal
        st.caption("Meses incompletos: compare pela 'Média por dia'.")

    # ===== Aba Ofensores =====
    with aba_ofe:                            # tudo aqui aparece na aba "Ofensores"
        dimensoes = {                        # nome amigável -> coluna real
            "Operador da descarga": COL_OPERADOR, "Cliente": COL_CLIENTE, "Origem (última parada)": COL_ORIGEM,
            "Viagem": COL_VIAGEM, "Segmento": COL_SEGMENTO, "SC": COL_SC,
        }
        c1, c2, c3 = st.columns(3)           # três controles lado a lado
        escolha = c1.selectbox("Analisar por", list(dimensoes))      # dimensão do ranking
        ordem = c2.radio("Ordem", ["Maiores ofensores", "Menores ofensores"], horizontal=True)  # ranking crescente/decrescente
        topo = c3.slider("Itens no gráfico", 5, 30, 10)              # quantos itens mostrar
        col = dimensoes[escolha]             # coluna real escolhida
        rank = df.groupby(col).agg(Registros=(COL_REMESSA, "size"), Pedidos=(COL_PEDIDOS, "sum")).reset_index()
        rank = rank.sort_values("Registros", ascending=False)        # maiores primeiro
        rank["% do total"] = (rank["Registros"] / rank["Registros"].sum() * 100).round(2)  # participação
        rank["% acumulado"] = rank["% do total"].cumsum().round(2)   # acumulado (curva de Pareto)
        vis = rank.tail(topo) if ordem.startswith("Menores") else rank.head(topo)  # recorta o topo ou o fim
        fig = px.bar(vis, x="Registros", y=col, orientation="h", text=vis["% do total"].astype(str) + "%")
        fig.update_layout(yaxis={"categoryorder": "total ascending"})  # maior barra no topo
        grafico(fig)                         # mostra o gráfico
        st.dataframe(rank, width="stretch", hide_index=True)         # ranking completo

    # ===== Aba Horários =====
    with aba_hor:                            # tudo aqui aparece na aba "Horários"
        mapa = df.groupby(["Dia da semana", "Hora"]).size().unstack(fill_value=0)  # linhas = dia, colunas = hora
        mapa = mapa.reindex(list(DIAS_SEMANA.values()), fill_value=0).reindex(columns=range(24), fill_value=0)  # completa
        fig = px.imshow(                     # mapa de calor
            mapa, aspect="auto", color_continuous_scale="Reds",
            labels={"x": "Hora da descarga", "y": "Dia da semana", "color": "Registros"},
        )
        grafico(fig)                         # mostra o gráfico

    # ===== Aba Dados =====
    with aba_dad:                            # tudo aqui aparece na aba "Dados"
        if resumo is not None:               # se existir resumo, mostra a tabela consolidada
            st.subheader("Resumo por SC")
            st.dataframe(resumo, width="stretch", hide_index=True)
        st.subheader("Detalhe filtrado")     # tabela de detalhe
        st.dataframe(df, width="stretch", hide_index=True)
        csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")  # CSV com ; e BOM, abre certo no Excel BR
        st.download_button("⬇️ Baixar CSV", csv, "falta_bipagem_filtrado.csv", "text/csv")  # botão de download
        buf = io.BytesIO()                   # arquivo na memória para o Excel
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:        # escritor de Excel
            df.to_excel(w, sheet_name="Dados", index=False)        # aba com o detalhe filtrado
            if resumo is not None:           # se houver resumo...
                resumo.to_excel(w, sheet_name="Resumo", index=False)  # ...grava a aba Resumo
        st.download_button(                  # botão de download do Excel
            "⬇️ Baixar Excel", buf.getvalue(), "falta_bipagem_filtrado.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":                   # o Streamlit executa o arquivo como script principal
    main()                                   # chama a interface
