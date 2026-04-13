import os
import io
import base64
import requests
import pandas as pd
from pathlib import Path
from calendar import monthrange
from datetime import datetime

import streamlit as st
import plotly.graph_objs as go

# ==================== CONFIGURAÇÕES ====================
API_TOKEN     = os.environ.get("PIPEDRIVE_TOKEN", "ea06a4f8-af74-49ad-ade9-90eedd9d720e")
FILTER_ID     = 74674
PIPEDRIVE_URL = (
    f"https://api.pipedrive.com/v1/deals"
    f"?filter_id={FILTER_ID}&status=won&sort=won_time DESC"
    f"&limit=500&start=0&api_token={API_TOKEN}"
)

TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "ea06a4f8-af74-49ad-ade9-90eedd9d720e")
CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "85f7de7e-7983-4c0f-92c2-376cfb34df68")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "uIX8Q~OShLe~HULZnX1NRu-acFPx1glIBSX-raZj")

ONEDRIVE_USER = "mariana.montagneri@boardacademy.com.br"
FILE_COLAB    = "Colaboradores comercial.xlsx"
FILE_METAS    = "metas_comercial.xlsx"
FOLDER_FOTOS  = "Fotos - Time Comercial"

cores_padrao = ["#3498db","#e74c3c","#2ecc71","#9b59b6","#f39c12","#1abc9c","#e67e22"]
colors = {"gold":"#D4AF37","text":"#ffffff","text_secondary":"#999999"}

st.set_page_config(page_title="Vendas do Mês", layout="wide")
st.markdown("""
<style>
    .block-container { padding: 1rem 2rem; }
    header, footer { visibility: hidden; }
    [data-testid="stToolbar"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ==================== GRAPH API ====================

def get_graph_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":CLIENT_ID,
              "client_secret":CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"},
        timeout=15
    )
    r.raise_for_status()
    return r.json()["access_token"]

def graph_get(path, token):
    r = requests.get(f"https://graph.microsoft.com/v1.0{path}",
                     headers={"Authorization":f"Bearer {token}"}, timeout=30)
    if not r.ok:
        st.error(f"Graph API erro {r.status_code}: {r.text}")
        r.raise_for_status()
    return r

def baixar_excel(token, nome):
    caminho = requests.utils.quote(nome)
    r = graph_get(f"/users/{ONEDRIVE_USER}/drive/root:/{caminho}:/content", token)
    return io.BytesIO(r.content)

def listar_fotos(token):
    pasta = requests.utils.quote(FOLDER_FOTOS)
    r = graph_get(f"/users/{ONEDRIVE_USER}/drive/root:/{pasta}:/children", token)
    fotos = {}
    for item in r.json().get("value", []):
        nome = item.get("name","")
        dl   = item.get("@microsoft.graph.downloadUrl")
        if dl and Path(nome).suffix.lower() in [".jpg",".jpeg",".png"]:
            fotos[Path(nome).stem] = base64.b64encode(requests.get(dl,timeout=15).content).decode()
    return fotos

def formatar_mil(v):
    return f"{v/1000:.0f} mil"

# ==================== DADOS ====================

@st.cache_data(ttl=300)
def buscar_dados():
    token      = get_graph_token()
    fotos_dict = listar_fotos(token)

    resp  = requests.get(PIPEDRIVE_URL, timeout=30)
    resp.raise_for_status()
    vendas = resp.json().get("data", [])

    dados = []
    for v in vendas:
        valor = v.get("value", 0)
        if valor <= 0:
            continue
        won_time = v.get("won_time","")
        try:
            dia = datetime.fromisoformat(won_time.replace("Z","+00:00")).day if won_time else None
        except Exception:
            dia = None
        dados.append({
            "Nome":  v.get("user_id",{}).get("name","—"),
            "Valor": valor,
            "Valor_multiplicador": float(v.get("7e0e43c2734751f77be292a72527f638a850ad50") or 0),
            "Dia":   dia,
            "referido": "indicacao-comercial" in str(v.get("54fc9258843cdf7ea126b6c5aca9d4dc93a3a718","")).lower()
        })

    df   = pd.DataFrame(dados)
    hoje = datetime.now()
    dias = list(range(1, monthrange(hoje.year, hoje.month)[1]+1))

    vb_dia, vm_dia, vol_dia = [], [], []
    for d in dias:
        vd = df[df["Dia"]==d]
        vb_dia.append(vd["Valor"].sum())
        vm_dia.append(vd["Valor_multiplicador"].sum())
        vol_dia.append(len(vd))

    buf_col    = baixar_excel(token, FILE_COLAB)
    df_equipes = pd.read_excel(buf_col, sheet_name=0, usecols="A,C", engine="openpyxl")
    df_equipes.columns = ["Nome","Equipe"]
    df_equipes["Equipe"] = df_equipes["Equipe"].replace({"Pioneer + Discovery":"Falcon"})

    equipes_desejadas = ["Sniper","Elite","Pioneer + Discover","Orion","LATAM","MGM","Atlantis","Legacy"]
    todas_equipes     = [e for e in df_equipes["Equipe"].dropna().unique() if e in equipes_desejadas]

    buf_metas  = baixar_excel(token, FILE_METAS)
    df_metas   = pd.read_excel(buf_metas, sheet_name=0, usecols="A,B,D,F", engine="openpyxl")
    df_metas.columns = ["Ano","Mes","Nome","Meta"]
    df_metas_f = df_metas[(df_metas["Ano"]==hoje.year)&(df_metas["Mes"]==hoje.month)]
    meta_board = df_metas_f["Meta"].sum()

    resumo = (df.groupby("Nome")
                .agg(Valor=("Valor","sum"),
                     Valor_multiplicador=("Valor_multiplicador","sum"),
                     Volume_ganhos=("Valor","count"))
                .reset_index()
                .merge(df_equipes, on="Nome", how="left")
                .merge(df_metas_f[["Nome","Meta"]], on="Nome", how="left"))
    resumo["Meta"] = resumo["Meta"].fillna(0)

    metas_eq = df_metas_f.merge(df_equipes, on="Nome", how="left").groupby("Equipe")["Meta"].sum().to_dict()
    df_eq    = df.merge(df_equipes, on="Nome", how="left")
    vd_eq    = df_eq.groupby(["Equipe","Dia"]).agg(Valor=("Valor","sum")).reset_index()

    totais = {}
    for eq in todas_equipes:
        ve, acum = vd_eq[vd_eq["Equipe"]==eq], 0
        for d2 in range(1,32):
            acum += ve[ve["Dia"]==d2]["Valor"].sum()
        totais[eq] = acum

    eq_ord  = [e for e,_ in sorted(totais.items(), key=lambda x: x[1], reverse=True)]
    val_ord = [totais[e] for e in eq_ord]

    tabela = sorted([{
        "Equipe": eq,
        "Meta": metas_eq.get(eq,0),
        "Valor_Realizado": totais.get(eq,0),
        "Atingimento": (totais.get(eq,0)/metas_eq.get(eq,1)*100) if metas_eq.get(eq,0)>0 else 0
    } for eq in eq_ord], key=lambda x: x["Atingimento"], reverse=True)

    vt  = resumo["Valor"].sum()
    vmt = resumo["Volume_ganhos"].sum()
    top3 = resumo.sort_values("Valor", ascending=False).head(3)

    return {
        "dias": dias, "vb_dia": vb_dia, "vm_dia": vm_dia, "vol_dia": vol_dia,
        "eq_ord": eq_ord, "val_ord": val_ord,
        "equipe_destaque": eq_ord[0] if eq_ord else "",
        "tabela": tabela, "meta_board": meta_board,
        "valor_total": vt, "volume_total": int(vmt),
        "top3_nomes": top3["Nome"].tolist(),
        "top3_valores": top3["Valor"].tolist(),
        "top3_fotos": [fotos_dict.get(n) for n in top3["Nome"].tolist()],
        "logo": fotos_dict.get("logo_board") or fotos_dict.get("logo_Board"),
        "volume_referidos": int(df[df["referido"]].shape[0]),
    }

# ==================== LAYOUT ====================
d = buscar_dados()
mes_label = datetime.now().strftime("%B").capitalize() + " - " + datetime.now().strftime("%Y")

# Header
c1, c2 = st.columns([4,1])
with c1:
    st.markdown(f"<h1 style='color:#fff;font-size:52px;margin:0'>Vendas do Mês</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#999;font-size:28px;margin:0'>{mes_label}</p>", unsafe_allow_html=True)
with c2:
    if d["logo"]:
        logo_b64 = d["logo"]
        st.markdown(f"<div style='text-align:right'><img src='data:image/png;base64,{logo_b64}' style='height:70px'></div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Gráficos superiores
g1, g2 = st.columns(2)

with g1:
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=d["dias"], y=d["vb_dia"], mode="lines+markers+text", name="Bruto",
        line=dict(width=4,shape="spline"), marker=dict(size=6,color="#FFD700"),
        text=[formatar_mil(v) if v>0 else None for v in d["vb_dia"]],
        textposition="top center", textfont=dict(color="#FFD700",size=14)))
    fig.add_trace(go.Scatter(x=d["dias"], y=d["vm_dia"], mode="lines+markers+text", name="Multiplicador",
        line=dict(width=4,shape="spline"), marker=dict(size=6,color="#B8860B"),
        text=[formatar_mil(v) if v>0 else None for v in d["vm_dia"]],
        textposition="top center", textfont=dict(color="#B8860B",size=14)))
    fig.update_layout(
        title=dict(text="Valor por Dia - Bruto x Multiplicador", font=dict(color="#fff",size=16)),
        xaxis=dict(title="Dia",range=[1,31],tickmode="linear",tick0=1,dtick=1,color="#999",gridcolor="#2a2a2a"),
        yaxis=dict(title="Valor (R$)",color="#999",gridcolor="#2a2a2a",range=[0,max(d["vm_dia"]+[1])+50000]),
        plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font=dict(color="#fff"),
        margin=dict(l=80,r=60,t=60,b=60), height=480, legend=dict(font=dict(color="#fff"))
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

with g2:
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=d["dias"], y=d["vol_dia"],
        marker=dict(color="#FFD700",line=dict(width=2,color="#f5f25c")),
        text=d["vol_dia"], textposition="inside",
        textfont=dict(color="black",size=18,family="Arial Black")))
    fig2.update_layout(
        bargap=0.06,
        title=dict(text="Volume de ganhos por dia", font=dict(color="#fff",size=16)),
        xaxis=dict(title="Dia",range=[0.5,max(d["dias"])+0.5],tickmode="linear",tick0=1,dtick=1,color="#999",gridcolor="#2a2a2a"),
        yaxis=dict(title="Volume",range=[0,max(d["vol_dia"]+[1])+3],color="#999",gridcolor="#2a2a2a"),
        plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font=dict(color="#fff"),
        margin=dict(l=80,r=60,t=60,b=60), height=480
    )
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

# Seção inferior
b1, b2, b3 = st.columns([2,1,1])

with b1:
    cores_eq = {eq:"#D4AF37" if eq==d["equipe_destaque"] else cores_padrao[i%len(cores_padrao)] for i,eq in enumerate(d["eq_ord"])}
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        y=d["eq_ord"], x=d["val_ord"], width=0.7, orientation="h",
        marker=dict(color=[cores_eq.get(eq,"#666") for eq in d["eq_ord"]]),
        text=[f"R$ {v:,.2f}" for v in d["val_ord"]],
        textposition="inside", insidetextanchor="start",
        textfont=dict(color="#fff",size=20)))
    fig3.update_layout(
        title=dict(text="Ranking de Vendas por Equipe", font=dict(color="#fff",size=28)),
        xaxis=dict(color="#999",gridcolor="#2a2a2a"),
        yaxis=dict(color="#fff",gridcolor="#2a2a2a",autorange="reversed",tickfont=dict(size=18)),
        plot_bgcolor="#0a0a0a", paper_bgcolor="#1a1a1a", font=dict(color="#fff"),
        margin=dict(l=150,r=100,t=70,b=60), height=580
    )
    st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

with b2:
    st.markdown(f"<h3 style='color:{colors['gold']};text-align:center;font-size:16px'>Ranking Atingimento de Meta por Equipe</h3>", unsafe_allow_html=True)
    html_tab = "<table style='width:100%;border-collapse:collapse;color:#fff;font-size:13px'>"
    html_tab += "<thead><tr style='border-bottom:2px solid #2a2a2a'><th style='padding:6px;text-align:left'>Equipe</th><th style='padding:6px;text-align:right'>Meta</th><th style='padding:6px;text-align:right'>Real.</th><th style='padding:6px;text-align:right'>%</th></tr></thead><tbody>"
    for eq in d["tabela"]:
        p   = eq["Atingimento"]
        cor = "#D4AF37" if p>=100 else "#a7ffcc" if p>=70 else "#fff1a2" if p>=20 else "#ff9589"
        html_tab += f"<tr style='border-bottom:1px solid #2a2a2a'><td style='padding:8px'>{eq['Equipe']}</td><td style='padding:8px;text-align:right'>R$ {eq['Meta']:,.0f}</td><td style='padding:8px;text-align:right'>R$ {eq['Valor_Realizado']:,.0f}</td><td style='padding:8px;text-align:right;color:{cor};font-weight:bold'>{p:.0f}%</td></tr>"
    html_tab += "</tbody></table>"
    st.markdown(html_tab, unsafe_allow_html=True)

with b3:
    st.markdown("<h3 style='color:#fff;text-align:center;font-size:20px'>Top 3 Vendedores do Mês</h3>", unsafe_allow_html=True)
    medalhas = ["1º","2º","3º"]
    for i,(nome,valor,foto) in enumerate(zip(d["top3_nomes"],d["top3_valores"],d["top3_fotos"])):
        borda = "#D4AF37" if i==0 else "#aaa"
        img   = f"<img src='data:image/jpeg;base64,{foto}' style='width:90px;height:90px;border-radius:50%;object-fit:cover;border:3px solid {borda}'>" if foto else ""
        st.markdown(f"<div style='text-align:center;margin-bottom:14px'>{img}<div style='color:{colors['gold']};font-weight:bold;font-size:14px;margin-top:6px'>{medalhas[i]} - {nome}</div><div style='color:#999;font-size:12px'>R$ {valor:,.2f}</div></div>", unsafe_allow_html=True)

# Auto-refresh 5 min
st.markdown("<script>setTimeout(()=>location.reload(),300000)</script>", unsafe_allow_html=True)
