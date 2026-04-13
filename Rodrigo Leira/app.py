import requests
from datetime import datetime
import pandas as pd
import numpy as np
from pathlib import Path
import base64
import io
from calendar import monthrange

import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objs as go

# ==================== CONFIGURAÇÕES ====================
import os

# Pipedrive
API_TOKEN     = os.environ.get("PIPEDRIVE_TOKEN", "")
FILTER_ID     = 74674
PIPEDRIVE_URL = (
    f"https://api.pipedrive.com/v1/deals"
    f"?filter_id={FILTER_ID}&status=won&sort=won_time DESC"
    f"&limit=500&start=0&api_token={API_TOKEN}"
)

# Azure / Graph API
TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "ea06a4f8-af74-49ad-ade9-90eedd9d720e")
CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "85f7de7e-7983-4c0f-92c2-376cfb34df68")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

# OneDrive — dono dos arquivos
ONEDRIVE_USER = "mariana_montagneri_boardacademy_com_br"

# Nomes dos arquivos no OneDrive
FILE_COLABORADORES = "Colaboradores comercial.xlsx"
FILE_METAS         = "metas_comercial.xlsx"
FOLDER_FOTOS       = "Fotos - Time Comercial"

coroas = ["🥇", "🥈", "🥉"]

# ==================== GRAPH API ====================

def get_graph_token():
    url  = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default"
    }
    r = requests.post(url, data=data, timeout=15)
    r.raise_for_status()
    return r.json()["access_token"]


def graph_get(path, token, stream=False):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(
        f"https://graph.microsoft.com/v1.0{path}",
        headers=headers, timeout=30, stream=stream
    )
    r.raise_for_status()
    return r


def baixar_excel(token, nome_arquivo):
    """Baixa um xlsx do OneDrive pessoal da Mariana e retorna bytes."""
    caminho = requests.utils.quote(nome_arquivo)
    url_path = f"/users/{ONEDRIVE_USER}/drive/root:/{caminho}:/content"
    r = graph_get(url_path, token, stream=True)
    return io.BytesIO(r.content)


def listar_fotos(token):
    """Retorna dict {nome_sem_ext: base64} de todas as fotos na pasta."""
    pasta = requests.utils.quote(FOLDER_FOTOS)
    url_path = f"/users/{ONEDRIVE_USER}/drive/root:/{pasta}:/children"
    r = graph_get(url_path, token)
    items = r.json().get("value", [])
    fotos = {}
    for item in items:
        nome = item.get("name", "")
        stem = Path(nome).stem
        download_url = item.get("@microsoft.graph.downloadUrl")
        if download_url and Path(nome).suffix.lower() in [".jpg", ".jpeg", ".png"]:
            conteudo = requests.get(download_url, timeout=15).content
            fotos[stem] = base64.b64encode(conteudo).decode()
    return fotos


# ==================== HELPERS ====================

def formatar_mil(valor):
    return f"{valor/1000:.0f} mil"


def buscar_foto_dict(nome, fotos_dict):
    return fotos_dict.get(nome)


# ==================== BUSCAR DADOS ====================

def buscar_dados_pipedrive():
    print("🔄 Buscando token Graph API...")
    token = get_graph_token()

    print("📸 Baixando fotos do OneDrive...")
    fotos_dict = listar_fotos(token)

    print("🔄 Buscando vendas do Pipedrive...")
    resp = requests.get(PIPEDRIVE_URL, timeout=30)
    resp.raise_for_status()
    vendas = resp.json().get("data", [])

    dados = []
    for v in vendas:
        valor = v.get("value", 0)
        if valor <= 0:
            continue
        nome             = v.get("user_id", {}).get("name", "—")
        valor_multi      = float(v.get("7e0e43c2734751f77be292a72527f638a850ad50") or 0)
        tag_aplicacao    = v.get("54fc9258843cdf7ea126b6c5aca9d4dc93a3a718", "")
        won_time         = v.get("won_time", "")
        try:
            dia = datetime.fromisoformat(won_time.replace("Z", "+00:00")).day if won_time else None
        except Exception:
            dia = None
        dados.append({
            "Nome":               nome,
            "Valor":              valor,
            "Valor_multiplicador": valor_multi,
            "Dia":                dia,
            "referido":           "indicacao-comercial" in str(tag_aplicacao).lower()
        })

    df = pd.DataFrame(dados)

    hoje          = datetime.now()
    qt_dias_mes   = monthrange(hoje.year, hoje.month)[1]
    dias          = list(range(1, qt_dias_mes + 1))

    acum_bruto, acum_multi    = [], []
    valor_bruto_por_dia       = []
    valor_multi_por_dia       = []
    volume_por_dia            = []
    soma_bruto = soma_multi   = 0
    volume_referidos_total    = int(df[df["referido"]].shape[0])

    for d in dias:
        vd              = df[df["Dia"] == d]
        vb              = vd["Valor"].sum()
        vm              = vd["Valor_multiplicador"].sum()
        soma_bruto     += vb
        soma_multi     += vm
        acum_bruto.append(soma_bruto)
        acum_multi.append(soma_multi)
        valor_bruto_por_dia.append(vb)
        valor_multi_por_dia.append(vm)
        volume_por_dia.append(len(vd))

    # --- Planilhas ---
    print("📋 Lendo planilha de colaboradores...")
    buf_col   = baixar_excel(token, FILE_COLABORADORES)
    df_equipes = pd.read_excel(buf_col, sheet_name=0, usecols="A,C", engine="openpyxl")
    df_equipes.columns = ["Nome", "Equipe"]
    df_equipes["Equipe"] = df_equipes["Equipe"].replace({"Pioneer + Discovery": "Falcon"})

    equipes_desejadas = ["Sniper", "Elite", "Pioneer + Discover", "Orion", "LATAM", "MGM", "Atlantis", "Legacy"]
    todas_equipes     = [e for e in df_equipes["Equipe"].dropna().unique() if e in equipes_desejadas]

    print("🎯 Lendo metas...")
    buf_metas   = baixar_excel(token, FILE_METAS)
    df_metas    = pd.read_excel(buf_metas, sheet_name=0, usecols="A,B,D,F", engine="openpyxl")
    df_metas.columns = ["Ano", "Mes", "Nome", "Meta"]
    df_metas_f  = df_metas[(df_metas["Ano"] == hoje.year) & (df_metas["Mes"] == hoje.month)]
    meta_board_total = df_metas_f["Meta"].sum()

    resumo = (
        df.groupby("Nome")
          .agg(Valor=("Valor","sum"), Valor_multiplicador=("Valor_multiplicador","sum"), Volume_ganhos=("Valor","count"))
          .reset_index()
    )
    resumo["Ticket_medio"] = resumo["Valor"] / resumo["Volume_ganhos"]
    resumo = resumo.merge(df_equipes, on="Nome", how="left")
    resumo = resumo.merge(df_metas_f[["Nome","Meta"]], on="Nome", how="left")
    resumo["Meta"] = resumo["Meta"].fillna(0)

    df_metas_eq   = df_metas_f.merge(df_equipes, on="Nome", how="left")
    metas_equipe  = df_metas_eq.groupby("Equipe")["Meta"].sum().to_dict()

    df_com_eq     = df.merge(df_equipes, on="Nome", how="left")
    vd_eq         = df_com_eq.groupby(["Equipe","Dia"]).agg(Valor=("Valor","sum")).reset_index()

    dias_completos       = list(range(1, 32))
    vendas_diarias_dict  = {}
    for equipe in todas_equipes:
        ve, acum = vd_eq[vd_eq["Equipe"] == equipe], 0
        vp = []
        for dia in dias_completos:
            acum += ve[ve["Dia"] == dia]["Valor"].sum()
            vp.append(acum)
        vendas_diarias_dict[equipe] = vp

    totais_equipe    = {eq: vendas_diarias_dict[eq][-1] for eq in todas_equipes}
    equipes_ordem    = [e for e, _ in sorted(totais_equipe.items(), key=lambda x: x[1], reverse=True)]
    valores_ordem    = [totais_equipe[e] for e in equipes_ordem]

    tabela_equipes = sorted([
        {
            "Equipe": eq,
            "Meta":   metas_equipe.get(eq, 0),
            "Valor_Realizado": totais_equipe.get(eq, 0),
            "Atingimento": (totais_equipe.get(eq,0) / metas_equipe.get(eq,1) * 100) if metas_equipe.get(eq,0) > 0 else 0
        }
        for eq in equipes_ordem
    ], key=lambda x: x["Atingimento"], reverse=True)

    valor_vendas_total      = resumo["Valor"].sum()
    valor_vendas_multi_total= resumo["Valor_multiplicador"].sum()
    volume_vendas_total     = resumo["Volume_ganhos"].sum()
    ticket_medio_geral      = valor_vendas_total / volume_vendas_total if volume_vendas_total > 0 else 0
    percentual_atingimento  = (valor_vendas_total / meta_board_total * 100) if meta_board_total > 0 else 0

    vendedor_destaque_row   = resumo.loc[resumo["Valor"].idxmax()]
    equipe_destaque_nome    = equipes_ordem[0]
    top3                    = resumo.sort_values("Valor", ascending=False).head(3)

    logo_base64             = fotos_dict.get("logo_board") or fotos_dict.get("logo_Board")
    foto_vendedor_base64    = buscar_foto_dict(vendedor_destaque_row["Nome"], fotos_dict)

    return {
        "valor_vendas_total":       valor_vendas_total,
        "valor_vendas_multi_total": valor_vendas_multi_total,
        "volume_vendas_total":      int(volume_vendas_total),
        "ticket_medio_geral":       ticket_medio_geral,
        "vendedor_destaque_nome":   vendedor_destaque_row["Nome"],
        "vendedor_destaque_valor":  vendedor_destaque_row["Valor"],
        "vendedor_destaque_valor_multi": vendedor_destaque_row["Valor_multiplicador"],
        "equipe_destaque_nome":     equipe_destaque_nome,
        "equipes_ordenadas":        equipes_ordem,
        "valores_ordenados":        valores_ordem,
        "foto_vendedor_base64":     foto_vendedor_base64,
        "logo_base64":              logo_base64,
        "acum_dias":                dias,
        "acum_bruto":               acum_bruto,
        "acum_multi":               acum_multi,
        "volume_por_dia":           volume_por_dia,
        "top3_nomes":               top3["Nome"].tolist(),
        "top3_valores":             top3["Valor"].tolist(),
        "top3_fotos":               [buscar_foto_dict(n, fotos_dict) for n in top3["Nome"].tolist()],
        "meta_board_total":         meta_board_total,
        "percentual_atingimento":   percentual_atingimento,
        "tabela_equipes":           tabela_equipes,
        "valor_bruto_por_dia":      valor_bruto_por_dia,
        "valor_multi_por_dia":      valor_multi_por_dia,
        "volume_referidos_total":   volume_referidos_total,
    }


# ==================== DADOS INICIAIS ====================
dados = buscar_dados_pipedrive()

valor_vendas            = dados["valor_vendas_total"]
valor_vendas_multi      = dados["valor_vendas_multi_total"]
volume_vendas           = dados["volume_vendas_total"]
ticket_medio            = dados["ticket_medio_geral"]
vendedor_destaque_nome  = dados["vendedor_destaque_nome"]
top3_nomes              = dados["top3_nomes"]
top3_valores            = dados["top3_valores"]
top3_fotos              = dados["top3_fotos"]
vendedor_destaque_valor = dados["vendedor_destaque_valor"]
vendedor_destaque_valor_multi = dados["vendedor_destaque_valor_multi"]
equipe_destaque_nome    = dados["equipe_destaque_nome"]
equipes_ordenadas       = dados["equipes_ordenadas"]
valores_ordenados       = dados["valores_ordenados"]
foto_vendedor_base64    = dados["foto_vendedor_base64"]
logo_base64             = dados["logo_base64"]
acum_dias               = dados["acum_dias"]
acum_bruto              = dados["acum_bruto"]
acum_multi              = dados["acum_multi"]
volume_por_dia          = dados["volume_por_dia"]
meta_board_total        = dados["meta_board_total"]
percentual_atingimento  = dados["percentual_atingimento"]
tabela_equipes          = dados["tabela_equipes"]
valor_bruto_por_dia     = dados["valor_bruto_por_dia"]
valor_multi_por_dia     = dados["valor_multi_por_dia"]
volume_referidos        = dados["volume_referidos_total"]

# Cores equipes
cores_padrao  = ["#3498db","#e74c3c","#2ecc71","#9b59b6","#f39c12","#1abc9c","#e67e22"]
cores_equipes = {
    eq: "#D4AF37" if eq == equipe_destaque_nome else cores_padrao[i % len(cores_padrao)]
    for i, eq in enumerate(equipes_ordenadas)
}

# ==================== ESTILOS ====================
external_stylesheets = [{
    "href": "data:text/css;charset=utf-8," + """
        body { zoom:1.5; -moz-transform:scale(1.3); -moz-transform-origin:0 0; overflow:hidden; }
        .modebar { background-color:transparent!important; box-shadow:none!important; display:none!important; }
        @keyframes pulseGold { 0%{box-shadow:0 0 6px rgba(212,175,55,.4)} 50%{box-shadow:0 0 18px rgba(212,175,55,.9)} 100%{box-shadow:0 0 6px rgba(212,175,55,.4)} }
        .card-pulse { animation:pulseGold 3s infinite ease-in-out; }
        @keyframes float { 0%{transform:translateY(0)} 50%{transform:translateY(-6px)} 100%{transform:translateY(0)} }
        .top1 { animation:float 4s ease-in-out infinite; }
    """,
    "rel": "stylesheet"
}]

colors = {
    "background": "#0a0a0a", "card": "#1a1a1a", "text": "#ffffff",
    "text_secondary": "#999999", "gold": "#D4AF37", "border": "#2a2a2a",
}

mes_atual = datetime.now().strftime("%B").capitalize()
ano_atual = datetime.now().strftime("%Y")
subtitulo = f"{mes_atual} - {ano_atual}"

# ==================== APP ====================
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
app.config["suppress_callback_exceptions"] = True
server = app.server  # necessário para o Render

app.layout = html.Div(
    style={"backgroundColor": colors["background"], "padding": "20px", "fontFamily": "Arial, sans-serif", "height": "100vh"},
    children=[
        # Intervalo de atualização (5 min)
        dcc.Interval(id="interval-refresh", interval=5*60*1000, n_intervals=0),

        # Header
        html.Div(style={"display":"flex","justifyContent":"space-between","alignItems":"center","marginBottom":"30px"}, children=[
            html.Div(children=[
                html.H1("Vendas do Mês", style={"color":colors["text"],"margin":"0","fontSize":"56px","fontWeight":"bold"}),
                html.Div(subtitulo, style={"color":colors["text_secondary"],"fontSize":"32px","marginTop":"5px"})
            ]),
            html.Div([
                html.Img(src=f"data:image/png;base64,{logo_base64}", style={"height":"70px"}) if logo_base64 else html.Div("Board Academy")
            ])
        ]),

        # Gráficos superiores
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"20px","marginBottom":"30px"}, children=[

            html.Div(style={"backgroundColor":colors["card"],"padding":"25px","borderRadius":"12px","border":f"1px solid {colors['border']}"}, children=[
                dcc.Graph(id="grafico-acumulado", config={"displayModeBar":False}, figure={
                    "data": [
                        go.Scatter(x=acum_dias, y=valor_bruto_por_dia, mode="lines+markers+text", name="Bruto",
                            line=dict(width=4,shape="spline"), marker=dict(size=6,color="#FFD700"),
                            text=[formatar_mil(v) if v>0 else None for v in valor_bruto_por_dia],
                            textposition="top center", textfont=dict(color="#FFD700",size=18)),
                        go.Scatter(x=acum_dias, y=valor_multi_por_dia, mode="lines+markers+text", name="Multiplicador",
                            line=dict(width=4,shape="spline"), marker=dict(size=6,color="#B8860B"),
                            text=[formatar_mil(v) if v>0 else None for v in valor_multi_por_dia],
                            textposition="top center", textfont=dict(color="#B8860B",size=18))
                    ],
                    "layout": go.Layout(
                        title=dict(text="Valor por Dia - Bruto x Multiplicador", font=dict(color=colors["text"],size=16)),
                        xaxis=dict(title="Dia",range=[1,31],tickmode="linear",tick0=1,dtick=1,color=colors["text_secondary"],gridcolor="#2a2a2a",tickfont=dict(size=18)),
                        yaxis=dict(title="Valor (R$)",color=colors["text_secondary"],gridcolor="#2a2a2a",range=[0,max(valor_multi_por_dia+[1])+50000]),
                        plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font={"color":colors["text"]},
                        margin={"l":100,"r":80,"t":70,"b":70}, height=700, showlegend=True,
                        legend=dict(font=dict(color=colors["text"]))
                    )
                })
            ]),

            html.Div(style={"backgroundColor":colors["card"],"padding":"25px","borderRadius":"12px","border":f"1px solid {colors['border']}"}, children=[
                dcc.Graph(id="grafico-volume-diario", config={"displayModeBar":False}, figure={
                    "data": [go.Bar(
                        x=acum_dias, y=volume_por_dia,
                        marker=dict(color="#FFD700",line=dict(width=2,color="#f5f25c")),
                        text=volume_por_dia, textposition="inside",
                        textfont=dict(color="black",size=22,family="Arial Black"),
                        hovertemplate="Dia %{x}: %{y} vendas<extra></extra>"
                    )],
                    "layout": go.Layout(
                        bargap=0.06,
                        title=dict(text="Volume de ganhos por dia", font=dict(color=colors["text"],size=16)),
                        xaxis=dict(title="Dia",range=[0.5,max(acum_dias)+0.5],tickmode="linear",tick0=1,dtick=1,color=colors["text_secondary"],gridcolor="#2a2a2a",tickfont=dict(size=18)),
                        yaxis=dict(title="Volume",range=[0,max(volume_por_dia+[1])+3],color=colors["text_secondary"],gridcolor="#2a2a2a"),
                        plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font={"color":colors["text"]},
                        margin={"l":100,"r":80,"t":70,"b":70}, height=700, showlegend=False
                    )
                })
            ])
        ]),

        # Seção inferior
        html.Div(style={"display":"grid","gridTemplateColumns":"2fr 1fr 1fr","gap":"18px"}, children=[

            # Ranking equipes
            html.Div(style={"backgroundColor":colors["card"],"padding":"18px","borderRadius":"12px","border":f"1px solid {colors['border']}"}, children=[
                dcc.Graph(id="grafico-vendas", config={"displayModeBar":False}, figure={
                    "data": [go.Bar(
                        y=equipes_ordenadas, x=valores_ordenados, width=0.7, orientation="h",
                        marker=dict(color=[colors["gold"] if eq==equipe_destaque_nome else cores_equipes.get(eq,"#666") for eq in equipes_ordenadas]),
                        text=[f"R$ {v:,.2f}" for v in valores_ordenados],
                        textposition="inside", insidetextanchor="start",
                        textfont=dict(color="#ffffff",size=30),
                        hovertemplate="%{y} Total: R$ %{x:,.2f}"
                    )],
                    "layout": go.Layout(
                        title=dict(text="Ranking de Vendas por Equipe", font=dict(color=colors["text"],size=45)),
                        xaxis=dict(title="Total de Vendas (R$)",color=colors["text_secondary"],gridcolor="#2a2a2a"),
                        yaxis=dict(title="",color=colors["text"],gridcolor="#2a2a2a",autorange="reversed",tickfont=dict(size=25)),
                        plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font={"color":colors["text"]},
                        margin={"l":180,"r":140,"t":70,"b":60}, height=720, showlegend=False
                    )
                })
            ]),

            # Tabela metas
            html.Div(style={"backgroundColor":colors["card"],"padding":"20px","borderRadius":"12px","border":f"1px solid {colors['border']}"}, children=[
                html.H3("Ranking Atingimento de Meta por Equipe", style={"color":colors["gold"],"marginBottom":"20px","fontSize":"22px","textAlign":"center"}),
                html.Table(style={"width":"100%","borderCollapse":"collapse"}, children=[
                    html.Thead(children=[html.Tr(style={"borderBottom":f"2px solid {colors['border']}"}, children=[
                        html.Th(t, style={"padding":"12px 8px","textAlign":"left" if t=="Equipe" else "right","color":colors["text"],"fontSize":"24px","fontWeight":"bold"})
                        for t in ["Equipe","Meta","Real.","%"]
                    ])]),
                    html.Tbody(children=[
                        html.Tr(style={"borderBottom":f"1px solid {colors['border']}"}, children=[
                            html.Td(eq["Equipe"], style={"padding":"15px 8px","color":colors["text"],"fontSize":"28px"}),
                            html.Td(f"R$ {eq['Meta']:,.0f}".replace(",","."), style={"padding":"15px 8px","textAlign":"right","color":colors["text"],"fontSize":"28px"}),
                            html.Td(f"R$ {eq['Valor_Realizado']:,.0f}".replace(",","."), style={"padding":"15px 8px","textAlign":"right","color":colors["text"],"fontSize":"28px"}),
                            html.Td(f"{eq['Atingimento']:.0f}%", style={
                                "padding":"10px 5px","textAlign":"right","fontSize":"28px","fontWeight":"bold",
                                "color": colors["gold"] if eq["Atingimento"]>=100 else "#a7ffcc" if eq["Atingimento"]>=70 else "#fff1a2" if eq["Atingimento"]>=20 else "#ff9589"
                            })
                        ]) for eq in tabela_equipes
                    ])
                ])
            ]),

            # Top 3
            html.Div(style={"backgroundColor":colors["card"],"padding":"25px","borderRadius":"12px","border":f"1px solid {colors['border']}","textAlign":"center"}, children=[
                html.H3("Top 3 Vendedores do Mês", style={"color":colors["text"],"marginBottom":"30px","fontSize":"46px"}),
                html.Div(style={"display":"flex","justifyContent":"space-around","alignItems":"center","gap":"20px"}, children=[
                    html.Div(children=[
                        html.Img(
                            src=f"data:image/jpeg;base64,{top3_fotos[i]}" if i < len(top3_fotos) and top3_fotos[i] else None,
                            className="top1",
                            style={"width":"240px","height":"240px","borderRadius":"50%","objectFit":"cover",
                                   "border":"4px solid #D4AF37" if i==0 else "4px solid #888"}
                        ),
                        html.Div(f"{['1º -','2º -','3º -'][i]} {top3_nomes[i]}", style={"fontSize":"28px","fontWeight":"bold","color":colors["gold"],"marginTop":"12px"}),
                        html.Div(f"R$ {top3_valores[i]:,.2f}", style={"fontSize":"26px","color":colors["text_secondary"],"marginTop":"5px"})
                    ], style={"textAlign":"center"})
                    for i in range(len(top3_nomes))
                ])
            ])
        ])
    ]
)


# ==================== CALLBACK ATUALIZAÇÃO ====================
@app.callback(
    Output("grafico-acumulado",    "figure"),
    Output("grafico-volume-diario","figure"),
    Output("grafico-vendas",       "figure"),
    Input("interval-refresh",      "n_intervals")
)
def atualizar(n):
    if n == 0:
        raise dash.exceptions.PreventUpdate
    d = buscar_dados_pipedrive()
    # reconstrói as figuras com dados novos — mesma lógica do layout inicial
    fig_acum = {
        "data": [
            go.Scatter(x=d["acum_dias"], y=d["valor_bruto_por_dia"], mode="lines+markers+text", name="Bruto",
                line=dict(width=4,shape="spline"), marker=dict(size=6,color="#FFD700"),
                text=[formatar_mil(v) if v>0 else None for v in d["valor_bruto_por_dia"]],
                textposition="top center", textfont=dict(color="#FFD700",size=18)),
            go.Scatter(x=d["acum_dias"], y=d["valor_multi_por_dia"], mode="lines+markers+text", name="Multiplicador",
                line=dict(width=4,shape="spline"), marker=dict(size=6,color="#B8860B"),
                text=[formatar_mil(v) if v>0 else None for v in d["valor_multi_por_dia"]],
                textposition="top center", textfont=dict(color="#B8860B",size=18))
        ],
        "layout": go.Layout(
            title=dict(text="Valor por Dia - Bruto x Multiplicador", font=dict(color=colors["text"],size=16)),
            xaxis=dict(title="Dia",range=[1,31],tickmode="linear",tick0=1,dtick=1,color=colors["text_secondary"],gridcolor="#2a2a2a",tickfont=dict(size=18)),
            yaxis=dict(title="Valor (R$)",color=colors["text_secondary"],gridcolor="#2a2a2a",range=[0,max(d["valor_multi_por_dia"]+[1])+50000]),
            plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font={"color":colors["text"]},
            margin={"l":100,"r":80,"t":70,"b":70}, height=700, showlegend=True,
            legend=dict(font=dict(color=colors["text"]))
        )
    }
    fig_vol = {
        "data": [go.Bar(x=d["acum_dias"], y=d["volume_por_dia"],
            marker=dict(color="#FFD700",line=dict(width=2,color="#f5f25c")),
            text=d["volume_por_dia"], textposition="inside",
            textfont=dict(color="black",size=22,family="Arial Black"))],
        "layout": go.Layout(
            bargap=0.06,
            title=dict(text="Volume de ganhos por dia", font=dict(color=colors["text"],size=16)),
            xaxis=dict(title="Dia",range=[0.5,max(d["acum_dias"])+0.5],tickmode="linear",tick0=1,dtick=1,color=colors["text_secondary"],gridcolor="#2a2a2a",tickfont=dict(size=18)),
            yaxis=dict(title="Volume",range=[0,max(d["volume_por_dia"]+[1])+3],color=colors["text_secondary"],gridcolor="#2a2a2a"),
            plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font={"color":colors["text"]},
            margin={"l":100,"r":80,"t":70,"b":70}, height=700, showlegend=False
        )
    }
    eq_ord = d["equipes_ordenadas"]
    val_ord = d["valores_ordenados"]
    eq_dest = d["equipe_destaque_nome"]
    cores_eq_new = {eq: "#D4AF37" if eq==eq_dest else cores_padrao[i%len(cores_padrao)] for i,eq in enumerate(eq_ord)}
    fig_eq = {
        "data": [go.Bar(
            y=eq_ord, x=val_ord, width=0.7, orientation="h",
            marker=dict(color=[cores_eq_new.get(eq,"#666") for eq in eq_ord]),
            text=[f"R$ {v:,.2f}" for v in val_ord],
            textposition="inside", insidetextanchor="start",
            textfont=dict(color="#ffffff",size=30))],
        "layout": go.Layout(
            title=dict(text="Ranking de Vendas por Equipe", font=dict(color=colors["text"],size=45)),
            xaxis=dict(title="Total de Vendas (R$)",color=colors["text_secondary"],gridcolor="#2a2a2a"),
            yaxis=dict(title="",color=colors["text"],gridcolor="#2a2a2a",autorange="reversed",tickfont=dict(size=25)),
            plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font={"color":colors["text"]},
            margin={"l":180,"r":140,"t":70,"b":60}, height=720, showlegend=False
        )
    }
    return fig_acum, fig_vol, fig_eq


if __name__ == "__main__":
    print("\n🚀 Dashboard iniciado!")
    print("📺 Acesse: http://localhost:8050")
    app.run(debug=True, host="0.0.0.0", port=8050)
