import os
import json
import pandas as pd
import geopandas as gpd
from dash import Dash, dcc, html, Input, Output
import plotly.graph_objects as go

# =============== 1) 路径设置（服务器/本地通用） ===============
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
LOCAL_GEOJSON = os.path.join(DATA_DIR, "NSW_LGA_2025.geojson")

# ✅ 直接读取本地文件
print("📂 Using local NSW_LGA_2025.geojson")
gdf = gpd.read_file(LOCAL_GEOJSON)
gdf.set_crs("EPSG:4326", inplace=True)


# =============== 2) 保留你原来的几何处理 ===============
gdf = gdf[gdf["STE_NAME21"] == "New South Wales"]
gdf = gdf.to_crs(3857)
gdf["geometry"] = gdf["geometry"].simplify(12, preserve_topology=True)
gdf["geometry"] = gdf["geometry"].buffer(0)
gdf = gdf.to_crs(4326)
geojson = json.loads(gdf.to_json())


# ================== 2) 名称标准化 ==================
STOP_PATTERNS = [
    r"\bCOUNCIL OF THE CITY OF\b",
    r"\bTHE CITY OF\b",
    r"\bCITY OF\b",
    r"\bCOUNCIL OF\b",
    r"\bMUNICIPAL COUNCIL\b",
    r"\bREGIONAL COUNCIL\b",
    r"\bSHIRE COUNCIL\b",
    r"\bCITY COUNCIL\b",
]
TRAILING_TYPES = [" COUNCIL"," SHIRE"," CITY"," REGIONAL"," MUNICIPALITY"]

def norm(name: str) -> str:
    if pd.isna(name):
        return ""
    s = str(name).upper()
    s = re.sub(r"\([^)]*\)", "", s)
    for pat in STOP_PATTERNS:
        s = re.sub(pat, "", s)
    for t in TRAILING_TYPES:
        s = s.replace(t, "")
    s = re.sub(r"[^A-Z ]", " ", s)
    s = re.sub(r"\b(OF|THE|AND|LOCAL|AREA|REGION)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

ALIASES = {
    "SYDNEY COUNCIL OF THE CITY OF": "SYDNEY",
    "COUNCIL OF THE CITY OF SYDNEY": "SYDNEY",
    "CITY OF SYDNEY": "SYDNEY",
    "THE HILLS": "HILLS SHIRE",
    "HILLS SHIRE COUNCIL": "HILLS SHIRE",
    "THE HILLS SHIRE COUNCIL": "HILLS SHIRE",
    "BAYSIDE NSW": "BAYSIDE",
    "CENTRAL COAST NSW": "CENTRAL COAST",
    "CAMPBELLTOWN NSW": "CAMPBELLTOWN",
    "LAKE MACQUARIE CITY": "LAKE MACQUARIE",
    "CITY OF LAKE MACQUARIE": "LAKE MACQUARIE",
    "NAMBUCCA SHIRE COUNCIL": "NAMBUCCA VALLEY",
    "NAMBUCCA": "NAMBUCCA VALLEY",
    "TWEED SHIRE COUNCIL": "TWEED",
    "RICHMOND VALLEY COUNCIL": "RICHMOND VALLEY",
    "CLARENCE VALLEY COUNCIL": "CLARENCE VALLEY",
    "MURRUMBIDGEE COUNCIL": "MURRUMBIDGEE",
    "WENTWORTH SHIRE": "WENTWORTH",
    "COONAMBLE COUNCIL": "COONAMBLE",
    "COOTAMUNDRA-GUNDAGAI REGIONAL COUNCIL": "COOTAMUNDRA-GUNDAGAI REGIONAL",
}
def apply_alias(s: str) -> str:
    return ALIASES.get(s, s)

# gdf 也做 norm，准备映射
gdf["LGA_norm"] = gdf["LGA_NAME25"].map(norm)
LGA_MAP = dict(zip(gdf["LGA_norm"], gdf["LGA_NAME25"]))

# ================== 3) 列出所有年度 Excel ==================
# ================== 3) 列出所有年度 Excel ==================
def extract_year(label):
    # 从文件名提取年份数字，例如 'Pound-Data-Report-2019-20.xlsx' → 2019
    m = re.search(r"(\d{4})-(\d{2})", label)
    if not m:
        return 0
    y1 = int(m.group(1))
    return y1  # 用前一年作排序基准

excel_files = sorted(
    [f for f in os.listdir(BASE_DIR)
     if f.startswith("Pound-Data-Report") and f.endswith(".xlsx")],
    key=extract_year
)

YEAR_OPTIONS = []
for f in excel_files:
    m = re.search(r"(\d{4}-\d{2})", f)
    label = m.group(1) if m else f
    YEAR_OPTIONS.append({"label": label, "value": f})



# ================== 4) 工具函数：读取某一年并做好 LGA 匹配 ==================
def load_year_df(filename: str) -> pd.DataFrame:
    path = os.path.join(BASE_DIR, filename)
    df = pd.read_excel(path, header=2)
    df.columns = df.columns.str.strip().str.replace("\n", " ")
    # 去 grand total
    df = df[~df["Council Name"].astype(str).str.contains(r"grand\s*total", case=False, na=False)]
    # 标准化
    df["Council_norm"] = df["Council Name"].map(norm).map(apply_alias)
    # 先直接匹配
    df["LGA_KEY"] = df["Council_norm"].map(LGA_MAP)
    # 模糊兜底
    missing = df["LGA_KEY"].isna()
    if missing.any():
        NSW_KEYS = list(gdf["LGA_norm"].unique())
        def fuzzy_match(x):
            m = get_close_matches(x, NSW_KEYS, n=1, cutoff=0.86)
            return LGA_MAP[m[0]] if m else None
        df.loc[missing, "LGA_KEY"] = df.loc[missing, "Council_norm"].map(fuzzy_match)
        # 特别处理 2019-20：这一年没有 Bayside，就补一行空的
    if "2019-20" in filename:
        has_bayside = df["Council_norm"].str.fullmatch("BAYSIDE").any()
        if not has_bayside:
            empty_row = {c: None for c in df.columns}
            empty_row["Council Name"] = "Bayside Council"
            empty_row["Council_norm"] = "BAYSIDE"
            empty_row["LGA_KEY"] = "Bayside (NSW)"  # 改成你 geojson 里的真名
            df = pd.concat([df, pd.DataFrame([empty_row])], ignore_index=True)

    return df


# ================== 5) Dash app ==================
app = Dash(__name__)
app.title = "NSW Council Animal Data"

app.layout = html.Div([
    html.H3(
        "NSW Council Animal Data – Multi Year Map",
        style={"textAlign": "center", "marginBottom": "6px"}
    ),

    # 顶部两个下拉
    html.Div([
        html.Div([
            html.Label("Year"),
            dcc.Dropdown(
                id="year-dropdown",
                options=YEAR_OPTIONS,
                value=excel_files[-1],
                clearable=False,
            )
        ], style={"width": "230px"}),

        html.Div([
            html.Label("Metric"),
            dcc.Dropdown(
                id="metric-dropdown",
                options=[],
                value=None,
                clearable=False,
            )
        ], style={"flex": "1", "minWidth": "420px", "marginLeft": "20px"}),
    ], style={
        "display": "flex",
        "flexWrap": "wrap",
        "gap": "16px",
        "margin": "10px 30px 4px 30px"
    }),

    # 主体：左地图 + 右折线
    html.Div([
        # 左：地图 70%
        html.Div([
            dcc.Graph(
                id="nsw-map",
                style={"height": "880px"}  # 地图更高
            )
        ], style={
            "flex": "0.7",
            "minWidth": "620px",        # 防止再被压瘪
        }),

        # 右：折线 30%
        html.Div([
            html.H5("Trend over years", style={"textAlign": "center"}),
            dcc.Graph(
                id="trend-line",
                style={"height": "880px"}
            )
        ], style={
            "flex": "0.3",
            "padding": "0 10px",
            "minWidth": "360px"
        })
    ], style={
        "display": "flex",
        "gap": "10px"
    }),

    html.Div(
        "Gray areas = No report",
        style={"textAlign": "center", "color": "gray", "marginTop": "6px"}
    )
])



# ================== 6) 回调1：选年份 → 填指标 ==================
@app.callback(
    Output("metric-dropdown", "options"),
    Output("metric-dropdown", "value"),
    Input("year-dropdown", "value"),
)
def update_metric_dropdown(excel_name):
    df = load_year_df(excel_name)
    cols = list(df.columns)

    # 按你原来的逻辑分狗猫
    cat_start_idx = next((i for i, c in enumerate(cols) if "(Cats)" in c), None)
    if cat_start_idx is not None:
        dog_section = cols[:cat_start_idx]
        cat_section = cols[cat_start_idx:]
        dog_cols = [c for c in dog_section if re.search(r"(dogs?)|euthanase", c, re.I)]
        cat_cols = [c for c in cat_section if re.search(r"(cats?)|euthanase", c, re.I)]
    else:
        dog_cols = [c for c in cols if re.search(r"(dogs?)|euthanase", c, re.I) and "cat" not in c.lower()]
        cat_cols = [c for c in cols if re.search(r"(cats?)|euthanase", c, re.I) and "dog" not in c.lower()]

    for drop_col in ["Council Name", "Council_norm", "LGA_KEY"]:
        dog_cols = [c for c in dog_cols if c != drop_col]
        cat_cols = [c for c in cat_cols if c != drop_col]

    metrics = dog_cols + cat_cols

    options = [
        {"label": ("🐶 " if m in dog_cols else "🐱 ") + m, "value": m}
        for m in metrics
    ]
    default_val = metrics[0] if metrics else None

    return options, default_val


# ================== 7) 回调2：画地图 ==================
@app.callback(
    Output("nsw-map", "figure"),
    Input("year-dropdown", "value"),
    Input("metric-dropdown", "value"),
)
def update_map(excel_name, metric_name):
    df = load_year_df(excel_name)

    if not metric_name:
        # 兜底选一个数值列
        possible = [c for c in df.columns if c not in ["Council Name","Council_norm","LGA_KEY"]]
        metric_name = possible[0]

    z_raw = pd.to_numeric(df[metric_name], errors="coerce")
    GRAY_SENTINEL = -0.001
    z_plot = z_raw.fillna(GRAY_SENTINEL)
    upper_limit = float(min(1000, z_raw.quantile(0.95) if z_raw.notna().any() else 0))

    hover_text = [
        f"{row['Council Name']}<br>{metric_name}: {val:.0f}" if pd.notna(val)
        else f"{row['Council Name']}<br>No report"
        for (_, row), val in zip(df.iterrows(), z_raw)
    ]

    fig = go.Figure()
    fig.add_choropleth(
        geojson=geojson,
        locations=df["LGA_KEY"],
        featureidkey="properties.LGA_NAME25",
        z=z_plot,
        zmin=GRAY_SENTINEL,
        zmax=max(upper_limit, 0),
        colorscale=[
            [0, "#f0f0f0"], [0.00001, "#e0ecf4"],
            [0.25, "#9ecae1"], [0.5, "#6baed6"],
            [0.75, "#3182bd"], [1.0, "#08519c"]
        ],
        marker_line_width=0.8,
        marker_line_color="white",
        hovertext=hover_text,
        hoverinfo="text",
        colorbar_title=metric_name
    )

    fig.update_layout(
    title=f"{metric_name} – {excel_name}",
    geo=dict(
        fitbounds="locations",
        visible=False,
        projection_type="mercator",
        center=dict(lat=-32.5, lon=147.0),  # NSW中心点
        lonaxis=dict(range=[140.5, 153.7]),  # 经度范围（手动放大）
        lataxis=dict(range=[-38.5, -27.8]),  # 纬度范围
    ),
    height=900,
    width=1100,  # ✅ 强制宽一点
    margin=dict(l=20, r=20, t=70, b=10),
    coloraxis_colorbar=dict(
        len=0.65,          # ✅ 缩短色条长度
        thickness=14,      # ✅ 变窄
        x=0.95,            # ✅ 靠右但不挤地图
        xanchor="right",
        outlinewidth=0
    )
)


    return fig


# ================== 8) 回调3：点击地图 → 右侧折线 ==================
# ======================
# 指标跨年度名称映射（老 ↔ 新）
# ======================
METRIC_ALIASES = {
    # ====== ① 进出场（Dogs）======
    # 老 → 新
    "Total Number of Dogs entering Facility during Year":
        "Total Incoming Animals (Dogs) - including Dogs in Council Facility on July 1",
    "Total Number of Dogs leaving Facility during Year":
        "Total Outgoing Animals (Dogs)",
    # 新 → 老（反向也写上，防止你从 2023-24 那边点）
    "Total Incoming Animals (Dogs) - including Dogs in Council Facility on July 1":
        "Total Number of Dogs entering Facility during Year",
    "Total Outgoing Animals (Dogs)":
        "Total Number of Dogs leaving Facility during Year",

    # ====== ② 进出场（Cats）======
    "Total Number of Cats entering Facility during Year":
        "Total Incoming Animals (Cats) - including Cats in Council Facility on July 1",
    "Total Number of Cats leaving Facility during Year":
        "Total Outgoing Animals (Cats)",
    "Total Incoming Animals (Cats) - including Cats in Council Facility on July 1":
        "Total Number of Cats entering Facility during Year",
    "Total Outgoing Animals (Cats)":
        "Total Number of Cats leaving Facility during Year",

    # ====== ③ Euthanase / Dangerous / Restricted ======
    # 这几年的写法来回变 / 有的多了“/ Other”，有的还少了右括号
    "Euthanase due to Dangerous/ Restricted/ Other":
        "Euthanase due to Dangerous/ Restricted",
    "Euthanase due to Dangerous/ Restricted":
        "Euthanase due to Dangerous/ Restricted/ Other",

    # cats 那半段通常用 .1 收尾，我们也映射一下
    "Euthanase due to Dangerous/ Restricted/ Other.1":
        "Euthanase due to Dangerous/ Restricted.1",
    "Euthanase due to Dangerous/ Restricted.1":
        "Euthanase due to Dangerous/ Restricted/ Other.1",

    # ====== ④ Other(...) 这一长串 ======
    # 有的年有 Adopt，有的年括号里少了个 )
    "Other (Including Stolen from Facility/Died at Facility/Adopt":
        "Other (Including Stolen from Facility/Died at Facility/Adopt)",
    "Other (Including Stolen from Facility/Died at Facility/Adopt)":
        "Other (Including Stolen from Facility/Died at Facility/Adopt",

    # Cats 那半段
    "Other (Including Stolen from Facility/Died at Facility/Adopt.1":
        "Other (Including Stolen from Facility/Died at Facility/Adopt).1",
    "Other (Including Stolen from Facility/Died at Facility/Adopt).1":
        "Other (Including Stolen from Facility/Died at Facility/Adopt.1",

    # 2020-21 之后这列换成了不带 Adopt 的版本，我们也接一下
    "Other (Including Stolen from Facility/Died at Facility)":
        "Other (Including Stolen from Facility/Died at Facility/Adopt)",
    "Other (Including Stolen from Facility/Died at Facility).1":
        "Other (Including Stolen from Facility/Died at Facility/Adopt).1",
}



def clean_metric_name(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s2 = s
    s2 = re.sub(r"on July 1\s+\d{4}", "on July 1", s2, flags=re.IGNORECASE)
    s2 = re.sub(r"\s+", " ", s2).strip()
    return s2


@app.callback(
    Output("trend-line", "figure"),
    Input("nsw-map", "clickData"),
    Input("metric-dropdown", "value"),
)
def update_trend(clickData, metric_name):
    if not clickData:
        return go.Figure(
            layout=dict(
                title="Click an LGA on the map to see its trend",
                xaxis_title="Year",
                yaxis_title=metric_name or "Value",
                template="plotly_white",
            )
        )

    # 当前点击的 LGA
    lga_clicked = clickData["points"][0]["location"]
    lga_clicked_norm = norm(lga_clicked)

    # 用户真正点的名字（保持原样，不要翻译掉！）
    user_metric = metric_name

    records = []

    for f in excel_files:
        # ① 取年份标签
        m = re.search(r"(\d{4}-\d{2})", f)
        year_label = m.group(1) if m else f

        # ② 这一年的表
        df_y = load_year_df(f)
        cols_this_year = list(df_y.columns)

        # ③ 为这一年准备一串“我都能接受的名字”
        #    1. 用户点的
        candidates = [user_metric]

        #    2. 如果这个名字在别名表里，那个别名也要试
        if user_metric in METRIC_ALIASES:
            candidates.append(METRIC_ALIASES[user_metric])

        #    3. 再把所有出现过的“老/新版本的进出场名字”也塞进去（防 2019-20 ↔ 2020-21 那次大改名）
        canonical_metrics = [
            "Total Number of Dogs entering Facility during Year",
            "Total Number of Dogs leaving Facility during Year",
            "Total Incoming Animals (Dogs) - including Dogs in Council Facility on July 1",
            "Total Outgoing Animals (Dogs)",
            "Total Number of Cats entering Facility during Year",
            "Total Number of Cats leaving Facility during Year",
            "Total Incoming Animals (Cats) - including Cats in Council Facility on July 1",
            "Total Outgoing Animals (Cats)",
        ]
        candidates.extend(canonical_metrics)

        # 全部清洗一遍，方便比
        clean_candidates = [clean_metric_name(c).lower() for c in candidates if c]

        # ④ 在这一年的列里找一个能对上的
        col_match = None
        for col in cols_this_year:
            col_clean = clean_metric_name(col).lower()
            if col_clean in clean_candidates:  # 完全一致
                col_match = col
                break
            # 或者列名很长，至少包含了我们点的那一部分
            if clean_candidates and clean_candidates[0] and clean_candidates[0] in col_clean:
                col_match = col
                break

        # 这一年实在没这个指标，就跳过这一年
        if not col_match:
            # print(f"[DEBUG] {year_label} no column for {user_metric}")
            continue

        # ⑤ 找这一年的这一行（LGA）
        row = df_y[df_y["LGA_KEY"] == lga_clicked]
        if row.empty:
            row = df_y[df_y["Council_norm"] == lga_clicked_norm]
        if row.empty:
            # 再宽松一点
            year_names = df_y["Council_norm"].dropna().unique().tolist()
            close = get_close_matches(lga_clicked_norm, year_names, n=1, cutoff=0.8)
            if close:
                row = df_y[df_y["Council_norm"] == close[0]]
        if row.empty:
            continue

        val = pd.to_numeric(row[col_match].iloc[0], errors="coerce")
        records.append((year_label, val))

    # ⑥ 真没数据
    if not records:
        return go.Figure(
            layout=dict(
                title=f"{lga_clicked} – no data for this metric across years",
                xaxis_title="Year",
                yaxis_title=user_metric,
                template="plotly_white",
            )
        )

    # ⑦ 把 '2017-18' → 2018 这样
    def parse_year_label(label: str) -> int:
        m = re.match(r"(\d{4})-(\d{2})", label)
        if m:
            base = int(m.group(1))
            tail = int(m.group(2))
            return 2000 + tail if tail < 100 else tail
        found = re.findall(r"\d{4}", label)
        return int(found[0]) if found else 0

    numeric_records = []
    for yl, v in records:
        yr = parse_year_label(yl)
        numeric_records.append((yr, v))

    numeric_records.sort(key=lambda x: x[0])
    years = [y for y, _ in numeric_records]
    vals = [v for _, v in numeric_records]

    # ⑧ 画图 + 强制整数刻度（去掉 2018.5 这种）
    fig = go.Figure(go.Scatter(
        x=years,
        y=vals,
        mode="lines+markers",
        line=dict(color="#08519c", width=3),
        marker=dict(size=8),
           # 有 NaN 也尽量连
    ))
    fig.update_layout(
        title=f"{lga_clicked} – {user_metric}",
        xaxis_title="Year",
        yaxis_title=user_metric,
        template="plotly_white",
        height=820,
        margin=dict(l=40, r=20, t=80, b=50),
        xaxis=dict(
            tickmode="array",
            tickvals=years,
            ticktext=[str(y) for y in years],
            dtick=1
        )
    )
    return fig







# ================== 9) run ==================
app = Dash(__name__)
server = app.server

if __name__ == "__main__":
    app.run_server(host="0.0.0.0", port=8050, debug=True)
