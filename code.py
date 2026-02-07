import streamlit as st
import pandas as pd
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import math

# ===================== CONFIG =====================
FILE_Q1 = r"C:\Users\45328750\Desktop\NMD\Market Forecast-Oct.xlsm"
FILE_Q2 = r"C:\Users\45328750\Desktop\NMD\Market Forecast-Nov.xlsm"
LABEL_Q1 = "Previous Rate"
LABEL_Q2 = "Current Rate"
SHEET_NAME = "Market Forecast Rate Indexes"
HEADER_ROW = 5  # Excel row 6

# ===================== HELPERS =====================
def month_diff(d1, d2):
    return (d2.year - d1.year) * 12 + (d2.month - d1.month)

def load_and_transform(file_path, label):
    df = pd.read_excel(
        file_path, sheet_name=SHEET_NAME, header=HEADER_ROW, engine="openpyxl"
    )
    df.columns = df.columns.astype(str).str.strip()
    df = df.rename(columns={
        df.columns[0]: "scenario_path",
        df.columns[1]: "rate_index"
    })
    df = df.dropna(subset=["rate_index"])
    df["scenario"] = (
        df["scenario_path"]
        .astype(str)
        .str.extract(r"/\s*([^/]+)\s*/")
        .iloc[:, 0]
        .str.strip()
    )
    df = df.dropna(subset=["scenario"])
    date_cols = df.columns[2:]
    parsed_dates = pd.to_datetime(date_cols, errors="coerce")
    valuation_date = parsed_dates.dropna()[0]

    long_df = df.melt(
        id_vars=["scenario", "rate_index"],
        value_vars=date_cols,
        var_name="date",
        value_name="rate"
    )

    long_df["date"] = pd.to_datetime(long_df["date"], errors="coerce")
    long_df = long_df.dropna(subset=["date", "rate"])
    long_df["month_offset"] = long_df["date"].apply(
        lambda x: month_diff(valuation_date, x)
    )
    long_df["snapshot"] = label
    return long_df

# ===================== STREAMLIT UI =====================
st.set_page_config(page_title="Market Forecast Rate Indexes", layout="wide")
st.title("Market Forecast Rate Indexes")

# File upload widgets
st.sidebar.header("Upload Excel Files")
file_q1 = st.sidebar.file_uploader("Upload Previous Rate File", type=["xls", "xlsx", "xlsm"])
file_q2 = st.sidebar.file_uploader("Upload Current Rate File", type=["xls", "xlsx", "xlsm"])

if file_q1 and file_q2:
    df_q1 = load_and_transform(file_q1, LABEL_Q1)
    df_q2 = load_and_transform(file_q2, LABEL_Q2)
    final_df = pd.concat([df_q1, df_q2], ignore_index=True)

    rate_indices = sorted(final_df["rate_index"].unique())
    scenarios = sorted(final_df["scenario"].unique())

    rate_index = st.selectbox("Rate Index", rate_indices)
    selected_scenarios = st.multiselect("Scenario", scenarios, default=scenarios[:1])

    def plot_curve_plotly(rate_index, scenarios):
        if not scenarios:
            st.warning("Please select at least one scenario.")
            return

        plot_df = final_df[
            (final_df["rate_index"] == rate_index) &
            (final_df["scenario"].isin(scenarios))
        ].sort_values("month_offset")

        n = len(scenarios)
        cols = min(2, n)
        rows = math.ceil(n / cols)

        subplot_titles = [f"{scenario}" for scenario in scenarios]

        fig = make_subplots(
            rows=rows, cols=cols, subplot_titles=subplot_titles,
            shared_yaxes=False, horizontal_spacing=0.08, vertical_spacing=0.12
        )

        for i, scenario in enumerate(scenarios):
            row = (i // cols) + 1
            col = (i % cols) + 1

            for snap in plot_df["snapshot"].unique():
                sub = plot_df[
                    (plot_df["scenario"] == scenario) &
                    (plot_df["snapshot"] == snap)
                ]

                if sub.empty:
                    continue

                fig.add_trace(
                    go.Scatter(
                        x=[f"T{int(x)}" for x in sub["month_offset"]],
                        y=sub["rate"],
                        mode="lines+markers",
                        name=f"{snap}",
                        legendgroup=f"{snap}",
                        showlegend=(i == 0),
                        hovertemplate="Month: %{x}<br>Rate: %{y:.2f}<extra></extra>"
                    ),
                    row=row, col=col
                )

            fig.update_xaxes(
                title_text="Month",
                showgrid=False,
                showline=True,
                linecolor="white",
                linewidth=2,
                row=row, col=col
            )

            fig.update_yaxes(
                title_text="Rate",
                showgrid=False,
                showline=True,
                linecolor="white",
                linewidth=2,
                row=row, col=col
            )

        fig.update_layout(
            height=400 * rows,
            width=700 * cols,
            title_text=f"{rate_index}",
            title_font_color='White',
            legend_title="Snapshot",
            margin=dict(t=60, l=40, r=40, b=40),
            paper_bgcolor="#3d3c3c",
            plot_bgcolor="#3d3c3c",
            font=dict(color="white")
        )

        st.plotly_chart(fig, use_container_width=True)

    plot_curve_plotly(rate_index, selected_scenarios)

else:
    st.info("Please upload both Excel files to proceed.")