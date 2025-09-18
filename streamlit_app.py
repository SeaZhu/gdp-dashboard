
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(page_title="Perception Dashboard", layout="wide")

# ------------- Helpers -------------
@st.cache_data
def load_excel(fp: str):
    xls = pd.ExcelFile(fp)
    sheets = {name: xls.parse(name) for name in xls.sheet_names}
    return sheets

def to_pct(v):
    if pd.isna(v): 
        return None
    return v*100 if v <= 1 else v

def metric_block(title, value):
    st.metric(title, f"{to_pct(value):.1f}%")

def make_stacked_bar(df, y_col, pos_col, neu_col, neg_col, title):
    plot_df = df.copy()
    # ensure percentage scale 0-100
    for c in [pos_col, neu_col, neg_col]:
        plot_df[c] = plot_df[c].apply(to_pct)
    fig = go.Figure()
    fig.add_bar(name="Positive", x=plot_df[y_col], y=plot_df[pos_col])
    fig.add_bar(name="Neutral",  x=plot_df[y_col], y=plot_df[neu_col])
    fig.add_bar(name="Negative", x=plot_df[y_col], y=plot_df[neg_col])
    fig.update_layout(barmode="stack", title=title, yaxis_title="Percent", xaxis_title=None, height=420, margin=dict(l=10,r=10,t=40,b=10))
    return fig

def donut(values, labels, title):
    fig = px.pie(values=values, names=labels, hole=0.55)
    fig.update_layout(title=title, height=300, margin=dict(l=10,r=10,t=40,b=10), legend_orientation="h", legend_y=-0.15)
    return fig

# ------------- Sidebar -------------
st.sidebar.header("Data Source")
default_path = Path("data/perception_summary.xlsx")

sheets = load_excel(str(default_path))

# display sheet names for quick sanity
with st.sidebar.expander("Sheets Loaded", expanded=False):
    for k in sheets.keys():
        st.write("•", k)

top_n = st.sidebar.slider("Top N to show (Strengths / Improve)", 3, 15, 5)

# ------------- Overall Summary -------------
st.markdown("## Dataset Descriptive Summary")
overall = sheets.get("Overall Summary")
if overall is None or overall.empty:
    st.info("Sheet 'Overall Summary' not found.")
else:
    # Expect columns: Category, Percentage
    overall["Percentage"] = overall["Percentage"].apply(to_pct)
    c1, c2, c3 = st.columns(3)
    for col, label in zip([c1, c2, c3], ["Positive", "Neutral", "Negative"]):
        try:
            val = overall.loc[overall["Category"].str.lower()==label.lower(), "Percentage"].values[0]
            with col:
                metric_block(label, val)
        except Exception:
            pass

    # Charts
    left, right = st.columns([1.2, 1.3])
    with left:
        st.plotly_chart(donut(overall["Percentage"], overall["Category"], "Overall Perception"), use_container_width=True)
    with right:
        barfig = px.bar(overall, x="Category", y="Percentage", text=overall["Percentage"].round(1))
        barfig.update_layout(yaxis_title="Percent", xaxis_title=None, height=300, margin=dict(l=10,r=10,t=40,b=10))
        barfig.update_traces(textposition="outside")
        st.plotly_chart(barfig, use_container_width=True)

# ------------- Top Strengths -------------
st.markdown("## Strength Areas (by % Positive, with Neutral & Negative)")
strengths = sheets.get("Top Strengths")
if strengths is None or strengths.empty:
    st.info("Sheet 'Top Strengths' not found.")
else:
    # Expect columns: Question, Positive, Neutral, Negative
    strengths_disp = strengths.copy()
    strengths_disp["Positive"] = strengths_disp["Positive"].apply(to_pct)
    strengths_disp["Neutral"]  = strengths_disp["Neutral"].apply(to_pct)
    strengths_disp["Negative"] = strengths_disp["Negative"].apply(to_pct)
    strengths_disp = strengths_disp.sort_values("Positive", ascending=False).head(top_n)
    st.plotly_chart(make_stacked_bar(strengths_disp, "Question", "Positive", "Neutral", "Negative", "Top Strengths"), use_container_width=True)
    with st.expander("View data"):
        st.dataframe(strengths_disp.reset_index(drop=True))

# ------------- Areas to Improve -------------
st.markdown("## Areas to Improve (lowest % Positive, with Neutral & Negative)")
improve = sheets.get("Areas to Improve")
if improve is None or improve.empty:
    st.info("Sheet 'Areas to Improve' not found.")
else:
    # Expect columns: Question, Positive, Neutral, Negative
    improve_disp = improve.copy()
    improve_disp["Positive"] = improve_disp["Positive"].apply(to_pct)
    improve_disp["Neutral"]  = improve_disp["Neutral"].apply(to_pct)
    improve_disp["Negative"] = improve_disp["Negative"].apply(to_pct)
    improve_disp = improve_disp.sort_values("Positive", ascending=True).head(top_n)
    st.plotly_chart(make_stacked_bar(improve_disp, "Question", "Positive", "Neutral", "Negative", "Areas to Improve"), use_container_width=True)
    with st.expander("View data"):
        st.dataframe(improve_disp.reset_index(drop=True))

# ------------- Yearly Trend -------------
st.markdown("## Yearly Trend")
trend = sheets.get("Yearly Trend")
if trend is None or trend.empty:
    st.info("Sheet 'Yearly Trend' not found.")
else:
    # Accept either 'Positive %' or columns Positive/Neutral/Negative
    cols = [c.lower() for c in trend.columns]
    if "positive %" in cols:
        # simple single-line chart
        t = trend.rename(columns={trend.columns[cols.index("positive %")]: "Positive %"})
        fig = px.line(t, x=t.columns[0], y="Positive %", markers=True, title="Positive % over time")
        fig.update_layout(yaxis_title="Percent", xaxis_title=None, height=340, margin=dict(l=10,r=10,t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("View data"):
            st.dataframe(t)
    else:
        # if full 3-series is provided
        for col in trend.columns:
            if col.lower() in ["positive", "neutral", "negative"]:
                trend[col] = trend[col].apply(to_pct)
        melted = trend.melt(id_vars=[trend.columns[0]], value_vars=[c for c in trend.columns if c.lower() in ["positive","neutral","negative"]],
                            var_name="Category", value_name="Percent")
        fig = px.line(melted, x=melted.columns[0], y="Percent", color="Category", markers=True, title="Perception over time")
        fig.update_layout(yaxis_title="Percent", xaxis_title=None, height=340, margin=dict(l=10,r=10,t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("View data"):
            st.dataframe(trend)

st.caption("First version • Built with Streamlit + Plotly. Adjustments welcome.")
