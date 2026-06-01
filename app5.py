"""
لوحة تحكم المبيعات الذكية - الكود الكامل
Smart Sales Dashboard - Full Version
=====================================
المتطلبات / Requirements:
    pip install streamlit pandas plotly anthropic openpyxl

تشغيل التطبيق / Run:
    streamlit run dashboard_app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import anthropic

# ─────────────────────────────────────────────
# 1. إعداد الصفحة
# ─────────────────────────────────────────────


st.set_page_config(
    page_title="لوحة تحكم المبيعات الذكية",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",  # أو "collapsed" إذا تبيها مخفية من البداية
)

# ─────────────────────────────────────────────
# 2. CSS مخصص - تصميم احترافي
# ─────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');

  html, body, [class*="css"] {
    font-family: 'Cairo', sans-serif !important;
    direction: rtl;
    }
  
   

  /* منع RTL على رسوم Plotly */
  .js-plotly-plot, .plotly, .plot-container {
    direction: ltr !important;
  }

  .main { background: #0a0f1e; }


  /* بطاقات KPI */
  .kpi-card {
    background: linear-gradient(135deg, rgba(255,255,255,0.05), rgba(255,255,255,0.02));
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px;
    padding: 20px 24px;
    text-align: center;
    transition: transform 0.2s, box-shadow 0.2s;
  }
  .kpi-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 12px 40px rgba(14,165,233,0.15);
  }
  .kpi-value { font-size: 2rem; font-weight: 900; margin: 8px 0 4px; }
  .kpi-label { font-size: 0.85rem; color: #64748b; }

  /* بطاقة الـ Insights */
  .insight-card {
    background: linear-gradient(135deg, #0f172a, #1e293b);
    border: 1px solid #0ea5e933;
    border-right: 4px solid #0ea5e9;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 12px;
    color: #e2e8f0;
    font-size: 0.95rem;
    line-height: 1.7;
  }

  /* عنوان القسم */
  .section-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #e2e8f0;
    margin-bottom: 16px;
    padding-bottom: 8px;
    border-bottom: 1px solid #1e293b;
  }

  /* شريط جانبي */
  [data-testid="stSidebar"] {
    background: #0f172a;
    border-left: 1px solid #1e293b;
  }

  /* hide streamlit branding */
  #MainMenu { visibility: hidden; }
  footer { visibility: hidden; }
  header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 3. الشريط الجانبي
# ─────────────────────────────────────────────
with st.expander("📂 رفع البيانات والإعدادات", expanded=False):
    uploaded_file = st.file_uploader(
        "اختر ملف Excel أو CSV",
        type=["xlsx", "xls", "csv"],
        help="ارفع ملف يحتوي على بيانات المبيعات"
    )
    show_raw_data = st.checkbox("عرض البيانات الخام", value=True)
    show_ai_insights = st.checkbox("تحليل الذكاء الاصطناعي", value=True)
    user_api_key = st.text_input(
        "🔑 مفتاح Claude API", 
        type="password",
        placeholder="sk-ant-..."
    )

# ─────────────────────────────────────────────
# 4. تحميل البيانات
# ─────────────────────────────────────────────
DEMO_DATA = {
    "المنتج":  ["منتج أ","منتج ب","منتج ج","منتج أ","منتج ب","منتج ج","منتج أ","منتج ب","منتج ج","منتج أ"],
    "المبيعات":[1200,    800,    1500,    900,    2000,    1100,    1800,    1300,    1700,    2200   ],
    "الكمية":  [30,      20,     40,      25,     50,      28,      45,      33,      42,      55     ],
    "المنطقة": ["شمال","جنوب","شمال","غرب","جنوب","غرب","شمال","جنوب","غرب","شمال"              ],
    "الشهر":   ["يناير","يناير","فبراير","فبراير","مارس","مارس","أبريل","أبريل","مايو","مايو"    ],
    "التقييم": [4.5,    3.8,    4.9,     4.1,    4.7,    3.5,    4.8,    4.2,    4.6,    5.0    ],
}

@st.cache_data
def load_data(file):
    """تحميل البيانات من الملف المرفوع"""
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        return df, None
    except Exception as e:
        return None, str(e)

def detect_columns(df):
    """كشف الأعمدة المهمة تلقائياً"""
    cols = {c.lower(): c for c in df.columns}
    mapping = {}

    # أعمدة المبيعات المحتملة
    sales_keywords = ["مبيعات", "sales", "revenue", "إيرادات", "المبيعات", "القيمة", "value", "amount"]
    for kw in sales_keywords:
        if kw in cols:
            mapping["sales"] = cols[kw]
            break

    # أعمدة المنتج
    product_keywords = ["منتج", "product", "item", "المنتج", "الصنف", "name"]
    for kw in product_keywords:
        if kw in cols:
            mapping["product"] = cols[kw]
            break

    # أعمدة المنطقة
    region_keywords = ["منطقة", "region", "area", "المنطقة", "city", "مدينة", "location"]
    for kw in region_keywords:
        if kw in cols:
            mapping["region"] = cols[kw]
            break

    # أعمدة التاريخ/الشهر
    date_keywords = ["شهر", "month", "date", "تاريخ", "الشهر", "التاريخ", "year", "سنة"]
    for kw in date_keywords:
        if kw in cols:
            mapping["date"] = cols[kw]
            break

    return mapping

# ─────────────────────────────────────────────
# 5. تحميل البيانات الفعلية أو التجريبية
# ─────────────────────────────────────────────
if uploaded_file:
    df, error = load_data(uploaded_file)
    if error:
        st.error(f"❌ خطأ في قراءة الملف: {error}")
        st.stop()
    col_map = detect_columns(df)
    using_demo = False
    st.sidebar.success(f"✅ تم تحميل {len(df)} سجل")
else:
    df = pd.DataFrame(DEMO_DATA)
    col_map = {
        "sales": "المبيعات",
        "product": "المنتج",
        "region": "المنطقة",
        "date": "الشهر",
    }
    using_demo = True

# إذا لم تُكتشف الأعمدة — اسمح للمستخدم باختيارها
if not col_map.get("sales") and not using_demo:
    st.warning("⚠️ لم يتم التعرف على الأعمدة تلقائياً. يرجى تحديدها:")
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    all_cols = df.columns.tolist()

    c1, c2, c3, c4 = st.columns(4)
    col_map["sales"]   = c1.selectbox("عمود المبيعات", numeric_cols)
    col_map["product"] = c2.selectbox("عمود المنتج",   all_cols)
    col_map["region"]  = c3.selectbox("عمود المنطقة",  ["—"] + all_cols)
    col_map["date"]    = c4.selectbox("عمود الشهر/التاريخ", ["—"] + all_cols)
    if col_map["region"] == "—": col_map.pop("region")
    if col_map["date"]   == "—": col_map.pop("date")

# اختصارات للأعمدة
SALES   = col_map.get("sales")
PRODUCT = col_map.get("product")
REGION  = col_map.get("region")
DATE    = col_map.get("date")

# ─────────────────────────────────────────────
# 6. فلاتر تفاعلية
# ─────────────────────────────────────────────
st.markdown("# 📊 لوحة تحكم المبيعات الذكية")
if using_demo:
    st.info("📌 يعرض التطبيق بيانات تجريبية — ارفع ملفك من الشريط الجانبي")

filter_cols = st.columns(4)

with filter_cols[0]:
    if PRODUCT and PRODUCT in df.columns:
        products = ["الكل"] + sorted(df[PRODUCT].dropna().unique().tolist())
        selected_product = st.selectbox("🏷️ المنتج", products)
    else:
        selected_product = "الكل"

with filter_cols[1]:
    if REGION and REGION in df.columns:
        regions = ["الكل"] + sorted(df[REGION].dropna().unique().tolist())
        selected_region = st.selectbox("🗺️ المنطقة", regions)
    else:
        selected_region = "الكل"

with filter_cols[2]:
    if DATE and DATE in df.columns:
        dates = ["الكل"] + sorted(df[DATE].dropna().unique().tolist())
        selected_date = st.selectbox("📅 الفترة", dates)
    else:
        selected_date = "الكل"

with filter_cols[3]:
    if SALES and SALES in df.columns:
        min_s = int(df[SALES].min())
        max_s = int(df[SALES].max())
        min_val, max_val = st.slider("💰 نطاق المبيعات", min_s, max_s, (min_s, max_s))
    else:
        min_val, max_val = 0, 999999999

# تطبيق الفلاتر
filtered = df.copy()
if selected_product != "الكل" and PRODUCT:
    filtered = filtered[filtered[PRODUCT] == selected_product]
if selected_region != "الكل" and REGION:
    filtered = filtered[filtered[REGION] == selected_region]
if selected_date != "الكل" and DATE:
    filtered = filtered[filtered[DATE] == selected_date]
if SALES and SALES in df.columns:
    filtered = filtered[(filtered[SALES] >= min_val) & (filtered[SALES] <= max_val)]

# ─────────────────────────────────────────────
# 7. بطاقات KPI
# ─────────────────────────────────────────────
st.markdown("---")
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

if SALES and SALES in filtered.columns:
    total   = filtered[SALES].sum()
    avg     = filtered[SALES].mean()
    maximum = filtered[SALES].max()
    count   = len(filtered)

    with kpi1:
        st.markdown(f"""
        <div class='kpi-card'>
          <div class='kpi-label'>💰 إجمالي المبيعات</div>
          <div class='kpi-value' style='color:#0ea5e9'>{total:,.0f}</div>
          <div class='kpi-label'>ريال سعودي</div>
        </div>""", unsafe_allow_html=True)

    with kpi2:
        st.markdown(f"""
        <div class='kpi-card'>
          <div class='kpi-label'>📦 عدد السجلات</div>
          <div class='kpi-value' style='color:#10b981'>{count:,}</div>
          <div class='kpi-label'>صفقة / سجل</div>
        </div>""", unsafe_allow_html=True)

    with kpi3:
        st.markdown(f"""
        <div class='kpi-card'>
          <div class='kpi-label'>📈 متوسط المبيعات</div>
          <div class='kpi-value' style='color:#f59e0b'>{avg:,.0f}</div>
          <div class='kpi-label'>لكل سجل</div>
        </div>""", unsafe_allow_html=True)

    with kpi4:
        st.markdown(f"""
        <div class='kpi-card'>
          <div class='kpi-label'>🏆 أعلى قيمة</div>
          <div class='kpi-value' style='color:#f43f5e'>{maximum:,.0f}</div>
          <div class='kpi-label'>أفضل أداء</div>
        </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 8. الرسوم البيانية
# ─────────────────────────────────────────────
st.markdown("---")

CHART_THEME = {
    "paper_bgcolor": "rgba(0,0,0,0)",
    "plot_bgcolor":  "rgba(0,0,0,0)",
    "font":          {"family": "Cairo", "color": "#94a3b8"},
    "xaxis":         {"gridcolor": "#1e293b", "zerolinecolor": "#1e293b"},
    "yaxis":         {"gridcolor": "#1e293b", "zerolinecolor": "#1e293b"},
    "colorway":      ["#0ea5e9","#f59e0b","#10b981","#f43f5e","#8b5cf6","#06b6d4"],
}

col_left, col_right = st.columns([3, 2])

# ── رسم المبيعات حسب المنتج ──────────────────
with col_left:
    st.markdown("<div class='section-title'>📦 المبيعات حسب المنتج</div>", unsafe_allow_html=True)
    if PRODUCT and SALES and PRODUCT in filtered.columns:
        by_product = (
            filtered.groupby(PRODUCT)[SALES]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
        )
        fig_bar = px.bar(
            by_product, x=PRODUCT, y=SALES,
            color=SALES,
            color_continuous_scale=["#0f172a","#0ea5e9"],
            text=SALES,
        )
        fig_bar.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            marker_line_width=0,
        )
        fig_bar.update_layout(
            **CHART_THEME,
            showlegend=False,
            coloraxis_showscale=False,
            margin=dict(t=10, b=10, l=0, r=0),
            height=300,
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("يرجى تحديد عمود المنتج والمبيعات")

# ── رسم الفطيرة حسب المنطقة ───────────────────
with col_right:
    st.markdown("<div class='section-title'>🗺️ توزيع المناطق</div>", unsafe_allow_html=True)
    if REGION and SALES and REGION in filtered.columns:
        by_region = filtered.groupby(REGION)[SALES].sum().reset_index()
        fig_pie = px.pie(
            by_region, values=SALES, names=REGION,
            hole=0.55,
            color_discrete_sequence=["#0ea5e9","#f59e0b","#10b981","#f43f5e","#8b5cf6"],
        )
        fig_pie.update_traces(
            textposition="outside",
            textinfo="label+percent",
        )
        fig_pie.update_layout(
            **CHART_THEME,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2),
            margin=dict(t=10, b=30, l=0, r=0),
            height=300,
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("لا يوجد عمود منطقة")

# ── رسم الخط الزمني ───────────────────────────
if DATE and SALES and DATE in filtered.columns:
    st.markdown("<div class='section-title'>📅 النمو عبر الزمن</div>", unsafe_allow_html=True)

    if PRODUCT and PRODUCT in filtered.columns:
        trend = (
            filtered.groupby([DATE, PRODUCT])[SALES]
            .sum()
            .reset_index()
        )
        fig_line = px.line(
            trend, x=DATE, y=SALES, color=PRODUCT,
            markers=True,
            color_discrete_sequence=["#0ea5e9","#f59e0b","#10b981","#f43f5e"],
        )
    else:
        trend = filtered.groupby(DATE)[SALES].sum().reset_index()
        fig_line = px.line(trend, x=DATE, y=SALES, markers=True)

    fig_line.update_traces(line_width=2.5, marker_size=8)
    fig_line.update_layout(
    **CHART_THEME,
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    margin=dict(t=10, b=10, l=0, r=0),
    height=280,
    yaxis_title=None,
    xaxis_title=None,
)

    st.plotly_chart(fig_line, use_container_width=True)

# ── Heatmap المنتج × المنطقة ──────────────────
if PRODUCT and REGION and SALES and all(c in filtered.columns for c in [PRODUCT, REGION, SALES]):
    st.markdown("<div class='section-title'>🔥 خريطة الحرارة: المنتج × المنطقة</div>", unsafe_allow_html=True)
    pivot = filtered.pivot_table(index=PRODUCT, columns=REGION, values=SALES, aggfunc="sum", fill_value=0)
    fig_heat = px.imshow(
        pivot,
        color_continuous_scale=[[0,"#0f172a"],[0.5,"#0ea5e9"],[1,"#f59e0b"]],
        aspect="auto",
        text_auto=True,
    )
    fig_heat.update_layout(
        **CHART_THEME,
        margin=dict(t=10, b=10, l=0, r=0),
        height=260,
        yaxis_title=None,
        xaxis_title=None,
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ─────────────────────────────────────────────
# 9. جدول البيانات
# ─────────────────────────────────────────────
if show_raw_data:
    st.markdown("---")
    st.markdown("<div class='section-title'>📋 جدول البيانات التفصيلي</div>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📄 البيانات المفلترة", "📊 ملخص إحصائي"])

    with tab1:
        st.dataframe(
    filtered,
    use_container_width=True,
    height=320,
        )
        # تحميل البيانات
        csv_data = filtered.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="⬇️ تحميل البيانات (CSV)",
            data=csv_data,
            file_name="sales_filtered.csv",
            mime="text/csv",
        )

    with tab2:
        if SALES and SALES in filtered.columns:
            st.dataframe(
                filtered[[SALES]].describe().T.style.format("{:.2f}"),
                use_container_width=True,
            )

# ─────────────────────────────────────────────
# 10. تحليل الذكاء الاصطناعي - Insights
# ─────────────────────────────────────────────
if show_ai_insights:
    st.markdown("---")
    st.markdown("<div class='section-title'>🤖 تحليل الذكاء الاصطناعي</div>", unsafe_allow_html=True)

    # إعداد ملخص البيانات للـ AI
    def build_summary(df, col_map):
        lines = [f"عدد السجلات: {len(df)}"]

        if col_map.get("sales") and col_map["sales"] in df.columns:
            s = df[col_map["sales"]]
            lines.append(f"إجمالي المبيعات: {s.sum():,.0f}")
            lines.append(f"متوسط المبيعات: {s.mean():,.0f}")
            lines.append(f"أعلى قيمة: {s.max():,.0f} | أدنى قيمة: {s.min():,.0f}")

        if col_map.get("product") and col_map["product"] in df.columns:
            top = df.groupby(col_map["product"])[col_map["sales"]].sum().sort_values(ascending=False)
            lines.append(f"أفضل منتج: {top.index[0]} ({top.iloc[0]:,.0f})")
            lines.append(f"أضعف منتج: {top.index[-1]} ({top.iloc[-1]:,.0f})")

        if col_map.get("region") and col_map["region"] in df.columns:
            top_r = df.groupby(col_map["region"])[col_map["sales"]].sum().sort_values(ascending=False)
            lines.append(f"أفضل منطقة: {top_r.index[0]} ({top_r.iloc[0]:,.0f})")

        return "\n".join(lines)

    summary_text = build_summary(filtered, col_map)

    ai_col1, ai_col2 = st.columns([3, 1])
    with ai_col1:
        st.markdown(f"""
        <div class='insight-card'>
          <strong>📊 ملخص البيانات الحالية:</strong><br>
          <pre style='margin:8px 0 0; color:#94a3b8; font-family:Cairo; font-size:0.85rem;'>{summary_text}</pre>
        </div>""", unsafe_allow_html=True)

    with ai_col2:
        analyze_btn = st.button("🔍 حلّل البيانات بالذكاء الاصطناعي", use_container_width=True)

    if analyze_btn:
        if not user_api_key:
            st.warning("⚠️ يرجى إدخال مفتاح API أولاً")
        elif not user_api_key.startswith("sk-ant-"):
            st.error("❌ المفتاح غير صحيح — يجب أن يبدأ بـ sk-ant-")
        else:
            with st.spinner("🤖 جارٍ التحليل..."):
                try:
                    client = anthropic.Anthropic(api_key=user_api_key)
                   
                    with client.messages.stream(
                        model="claude-sonnet-4-20250514",
                        max_tokens=1000,
                        messages=[{"role": "user", "content": f"حلل هذه البيانات وأعطني insights عملية بالعربي:\n{summary_text}"}]
                    ) as stream:
                        insight_placeholder = st.empty()
                        full_response = ""
                        for text in stream.text_stream:
                            full_response += text
                            insight_placeholder.markdown(f"""
                            <div class='insight-card'>
                              {full_response}▌
                            </div>""", unsafe_allow_html=True)
                        insight_placeholder.markdown(f"""
                        <div class='insight-card'>
                          {full_response}
                        </div>""", unsafe_allow_html=True)
                except anthropic.AuthenticationError:
                    st.error("❌ المفتاح غير صالح، تحقق منه مجدداً")
                except Exception as e:
                    st.error(f"❌ خطأ: {str(e)}")

# ─────────────────────────────────────────────
# 11. تذييل الصفحة
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align:center; color:#1e293b; font-size:0.8rem; padding:16px;'>
    Smart Sales Dashboard v2.0 — تحليل ذكي لبياناتك 📊
</div>
""", unsafe_allow_html=True)
