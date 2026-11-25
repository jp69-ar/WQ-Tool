import streamlit as st
import pandas as pd
import base64
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from datetime import timedelta

# =====================================================
# PAGE SETUP
# =====================================================

st.set_page_config(
    page_title="WQ Tool",
    layout="wide",
    page_icon=":droplet:"
)

st.markdown("""
<style>
:root {
    color-scheme: light;
}

/* Main page scrolling */
[data-testid="stAppViewContainer"] {
    overflow-y: auto !important;
}

/* Sidebar stays in place */
section[data-testid="stSidebar"] {
    height: 100vh !important;
}

/* Sidebar content scrolls independently */
section[data-testid="stSidebar"] > div:first-child {
    height: 100vh !important;
    overflow-y: auto !important;
    overflow-x: hidden !important;
}

/* Leave room below the final sidebar control */
section[data-testid="stSidebar"] div[data-testid="stSidebarContent"] {
    padding-bottom: 5rem !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Water Quality Reporting Tool")

with st.expander("User Guide"):
    st.markdown("""
    It is recommended to run this app in light mode (can change with 3 dots in the top right corner), and to zoom out slightly in your browser.
    
    - Upload WQ data to generate report ready interactive figures and a data summary sheet.
    
    - Sites will likely require manual column mapping. Ensure the correct sheet/tab is corrected, and use the WQ data to map the correct columns within this app.
    
    - Note that when the units are contained within the parameter column (e.g. Turbidity (NTU)), the "Auto_Unit" function can be used by selecting it within the Unit Column within column mapping, which separates the units for data analysis and presentation.
    
    - Multiple water sources can be presented either within one worksheet, or split between multiple worksheets with one water source each - check which is applicable to your data.
    
    - This uses a default PCV and limit database which contains UK, EU, and WHO limit values, as well as water company internal standards.
    
    - To see more information on the details of the compliance limit, hover over it.
    
    :red-background[**Note - if an error appears after loading your dataset, it is likely that column mapping is configured incorrectly, or that there is insufficient data for that part of the app to work. Ensure all columns are selected correctly.**]
    
    """
)

# =====================================================
# FUNCTIONS
# =====================================================


def statistics_table(
    data,
    result_col,
    selected_component,
    sample_unit=""
):

    limit_value = None
    breach_count = 0

    if not component_limits.empty:

        limit_value = pd.to_numeric(
            component_limits["Limit Value"].iloc[0],
            errors="coerce"
        )

        if pd.notna(limit_value):
            breach_count = (
                data[result_col] > limit_value
            ).sum()

    stats = {
        "Component": [selected_component],
        "No. Samples": [len(data)],
        "No. Breaches": [breach_count],
        "% Exceeding": [
            100 * breach_count / len(data)
            if len(data) > 0 else 0
        ],
        "Units": [sample_unit],
        "Minimum": [data[result_col].min()],
        "P10": [data[result_col].quantile(0.10)],
        "Mean": [data[result_col].mean()],
        "P50": [data[result_col].quantile(0.50)],
        "P90": [data[result_col].quantile(0.90)],
        "P95": [data[result_col].quantile(0.95)],
        "P99": [data[result_col].quantile(0.99)],
        "Maximum": [data[result_col].max()]
    }

    if pd.notna(limit_value):
        stats["Limit"] = [limit_value]

    return pd.DataFrame(stats).round(3)
    
def build_export(
    stats_df,
    filtered_df,
    selected_site,
    selected_component,
    start_date,
    end_date,
    source_filename,
    component_limits,
    sample_unit
):

    output = BytesIO()

    summary_df = pd.DataFrame({
        "Item": [
            "Source File",
            "Water Source",
            "Component",
            "Unit",
            "Monitoring Start",
            "Monitoring End",
            "Sample Count",
            "Detected Results",
            "Non-Detect Results",
            "% Non-Detect"
        ],
        "Value": [
            source_filename,
            selected_site,
            selected_component,
            sample_unit,
            pd.to_datetime(start_date).strftime("%d %b %Y"),
            pd.to_datetime(end_date).strftime("%d %b %Y"),
            len(filtered_df),
            (~filtered_df["Non_Detect"]).sum(),
            filtered_df["Non_Detect"].sum(),
            round(
                filtered_df["Non_Detect"].mean() * 100,
                1
            )
        ]
    })
    
    compliance_rows = []

    if not component_limits.empty:

        sample_unit = ""

        if (
            UNIT_COL is not None
            and UNIT_COL in filtered_df.columns
        ):
            unit_values = (
                filtered_df[UNIT_COL]
                .dropna()
                .astype(str)
            )

            if not unit_values.empty:
                sample_unit = (
                    unit_values.mode().iloc[0]
                )
                
        for _, limit_row in component_limits.iterrows():

            limit_value = get_converted_limit(
                limit_row,
                sample_unit
            )

            if pd.notna(limit_value):

                exceedances = filtered_df[
                    filtered_df[RESULT_COL] > limit_value
                ]

                compliance_rows.append({
                    "Regulatory Source":
                        limit_row["Regulatory Source"],
                    "Limit":
                        limit_value,
                    "Unit":
                        sample_unit,
                    "Maximum Result":
                        filtered_df[RESULT_COL].max(),
                    "Breaches":
                        len(exceedances),
                    "Status":
                        (
                            "FAIL"
                            if len(exceedances) > 0
                            else "PASS"
                        )
                })

    compliance_df = pd.DataFrame(
        compliance_rows
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="01 Summary",
            index=False
        )

        stats_df.to_excel(
            writer,
            sheet_name="02 Statistics",
            index=False
        )
        
        if not compliance_df.empty:

            compliance_df.to_excel(
                writer,
                sheet_name="02 Statistics",
                startrow=len(stats_df) + 3,
                index=False
            )

        filtered_df.to_excel(
            writer,
            sheet_name="03 Data",
            index=False
        )
        
        if not compliance_df.empty:

            compliance_df.to_excel(
                writer,
                sheet_name="04 Compliance",
                index=False
            )

    return output.getvalue()

from reportlab.platypus import Image

def build_pdf_report(
    stats_df,
    summary_df,
    selected_site,
    selected_component,
    start_date,
    end_date,
    source_filename,
    sample_unit,
    figure_path=None
):

    from io import BytesIO
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        Image,
        PageBreak
    )
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4, landscape

    pdf_buffer = BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#0F4C81"),
        fontSize=24
    )

    heading_style = ParagraphStyle(
        "Heading",
        parent=styles["Heading2"],
        textColor=colors.HexColor("#0F4C81")
    )

    content = []

    # =====================================================
    # COVER PAGE
    # =====================================================

    content.append(
        Paragraph(
            "Water Quality Assessment Report",
            title_style
        )
    )

    content.append(Spacer(1, 25))

    metadata = [
        ["Source File", source_filename],
        ["Water Source", selected_site],
        ["Component", selected_component],
        ["Unit", sample_unit],
        [
            "Monitoring Period",
            f"{pd.to_datetime(start_date).strftime('%d %b %Y')} "
            f"to "
            f"{pd.to_datetime(end_date).strftime('%d %b %Y')}"
        ],
        [
            "Generated",
            pd.Timestamp.now().strftime("%d %b %Y %H:%M")
        ]
    ]

    cover_table = Table(
        metadata,
        colWidths=[130, 500]
    )

    cover_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#0F4C81")),
            ("TEXTCOLOR", (0,0), (0,-1), colors.white),
            ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.5, colors.black),
            ("ROWBACKGROUNDS", (1,0), (1,-1),
             [colors.whitesmoke, colors.lightgrey])
        ])
    )

    content.append(cover_table)

    content.append(Spacer(1, 20))

    if not stats_df.empty:

        samples = stats_df.iloc[0].get("Samples", "")
        maximum = stats_df.iloc[0].get("Maximum", "")

        summary_text = f"""
        <b>Executive Summary</b><br/><br/>
        No. samples analysed: {samples}<br/>
        Maximum recorded value: {value_with_unit(maximum, sample_unit)}<br/><br/>
        <u>Recorded Exceedances</u>
        """

        content.append(
            Paragraph(summary_text, styles["BodyText"])
        )

        for _, limit_row in component_limits.iterrows():

            limit_value = get_converted_limit(
                limit_row,
                sample_unit
            )

            if pd.notna(limit_value):

                exceedances = filtered[
                    filtered[RESULT_COL] > limit_value
                ]

                regulation = limit_row["Regulatory Source"]
                poc = limit_row.get("Point of Compliance", "")
                units = sample_unit

                content.append(
                    Paragraph(
                        f"{regulation} - {poc} "
                        f"({limit_value} {units}): "
                        f"{len(exceedances)} exceedances",
                        styles["BodyText"]
                    )
                )

    content.append(PageBreak())

    # =====================================================
    # FIGURE PAGE
    # =====================================================

    if figure_path:

        content.append(
            Paragraph(
                "Component Results Figure",
                heading_style
            )
        )

        content.append(Spacer(1, 10))

        img = Image(
            figure_path,
            width=800,
            height=450
        )

        content.append(img)

        content.append(PageBreak())

    # =====================================================
    # STATISTICS
    # =====================================================

    content.append(
        Paragraph(
            "Statistical Summary",
            heading_style
        )
    )

    content.append(Spacer(1, 10))

    stats_table = Table(
        [stats_df.columns.tolist()]
        + stats_df.values.tolist(),
        repeatRows=1
    )

    stats_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0,0), (-1,0),
             colors.HexColor("#0F4C81")),
            ("TEXTCOLOR", (0,0), (-1,0),
             colors.white),
            ("FONTNAME", (0,0), (-1,0),
             "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1),
             0.5, colors.black),
            ("FONTSIZE", (0,0), (-1,-1), 10),
            ("ROWBACKGROUNDS", (0,1), (-1,-1),
             [colors.whitesmoke, colors.lightgrey])
        ])
    )

    content.append(stats_table)

    content.append(Spacer(1, 20))

    # =====================================================
    # COMPLIANCE
    # =====================================================

    content.append(
        Paragraph(
            "Compliance Summary",
            heading_style
        )
    )

    content.append(Spacer(1, 10))

    if not summary_df.empty:

        compliance_table = Table(
            [summary_df.columns.tolist()]
            + summary_df.values.tolist(),
            repeatRows=1
        )

        compliance_table.setStyle(
            TableStyle([
                ("BACKGROUND", (0,0), (-1,0),
                 colors.HexColor("#0F4C81")),
                ("TEXTCOLOR", (0,0), (-1,0),
                 colors.white),
                ("FONTNAME", (0,0), (-1,0),
                 "Helvetica-Bold"),
                ("GRID", (0,0), (-1,-1),
                 0.5, colors.black),
                ("FONTSIZE", (0,0), (-1,-1), 6),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0,1), (-1,-1),
                 [colors.whitesmoke, colors.lightgrey])
            ])
        )

        content.append(compliance_table)

    else:

        content.append(
            Paragraph(
                "No compliance summary generated.",
                styles["BodyText"]
            )
        )


    # =====================================================
    # HEADER / FOOTER
    # =====================================================

    def header_footer(canvas, doc):

        canvas.saveState()

        canvas.setFont(
            "Helvetica-Bold",
            9
        )

        canvas.drawString(
            20,
            575,
            "Water Quality Assessment Report"
        )

        canvas.setFont(
            "Helvetica",
            8
        )

        canvas.drawRightString(
            810,
            15,
            f"Page {doc.page}"
        )

        canvas.restoreState()

    doc.build(
        content,
        onFirstPage=header_footer,
        onLaterPages=header_footer
    )

    pdf_buffer.seek(0)

    return pdf_buffer.getvalue()

def create_export_figure(
    filtered,
    selected_site,
    selected_component,
    date_col,
    result_col,
    unit_col
):

    mean = filtered[result_col].mean()
    p50 = filtered[result_col].quantile(0.50)
    p95 = filtered[result_col].quantile(0.95)

    units = ""

    if len(filtered):
        units = ""

        if (
            unit_col is not None
            and unit_col in filtered.columns
            and not filtered.empty
        ):
            units = str(filtered[unit_col].iloc[0])

    detect_df = filtered[~filtered["Non_Detect"]]
    nondetect_df = filtered[filtered["Non_Detect"]]

    fig = go.Figure()

    # DETECTS
    fig.add_trace(
        go.Scatter(
            x=detect_df[date_col],
            y=detect_df[result_col],
            mode="markers",
            name="Detected",
            marker=dict(
                color="#005A9C",
                size=10,
                line=dict(
                    color="#003B5C",
                    width=0.75
                )
            )
        )
    )

    # NON-DETECTS
    fig.add_trace(
        go.Scatter(
            x=nondetect_df[date_col],
            y=nondetect_df[result_col],
            mode="markers",
            name="Non-Detect (<)",
            marker=dict(
                color="red",
                size=7,
                symbol="circle",
                line=dict(
                    color="darkred",
                    width=0.5
                )
            )
        )
    )

    fig.add_trace(
        go.Scatter(
            x=[filtered[DATE_COL].min(), filtered[DATE_COL].max()],
            y=[mean, mean],
            mode="lines",
            name=(
                    f"Mean ({mean:.2f} {units})"
                    if units
                    else f"Mean ({mean:.2f})"
                ),
            line=dict(
                color="#000000",
                width=3,
                dash="dot"
            ),
            hovertemplate=f"Mean: {mean:.2f}<extra></extra>"
        )
    )
                 
    fig.add_trace(
        go.Scatter(
            x=[filtered[DATE_COL].min(), filtered[DATE_COL].max()],
            y=[p50, p50],
            mode="lines",
            name=(
                    f"P50 ({p50:.2f} {units})"
                    if units
                    else f"P50 ({p50:.2f})"
                ),
            line=dict(
                color="#4B0082",
                width=3,
                dash="dash"
            ),
            hovertemplate=f"P50: {p50:.2f}<extra></extra>"
        )
    )

    fig.add_trace(
        go.Scatter(
            x=[filtered[DATE_COL].min(), filtered[DATE_COL].max()],
            y=[p95, p95],
            mode="lines",
            name=(
                    f"P95 ({p95:.2f} {units})"
                    if units
                    else f"P95 ({p95:.2f})"
                ),
            line=dict(
                color="#FF00FF",
                width=3,
                dash="dash"
            ),
            hovertemplate=f"P95: {p95:.2f}<extra></extra>"
        )
    )

    fig.update_layout(
        width=2400,
        height=1350,
        template="simple_white",

        # SHOW LEGEND
        showlegend=True,

        legend=dict(
            x=1.02,
            y=1.0,
            font=dict(size=20)
        ),

        title=dict(
            text=f"<b>{selected_site.upper()}</b><br>{selected_component}",
            x=0.01,
            xanchor="left",
            font=dict(size=34)
        ),

        font=dict(
            family="Arial",
            size=24,
            color="black"
        ),

        paper_bgcolor="white",
        plot_bgcolor="white",

        margin=dict(
            l=140,
            r=250,
            t=160,
            b=140
        )
    )

    fig.update_xaxes(
        title="Sample Date",
        title_font=dict(size=28),
        tickfont=dict(size=20),
        showline=True,
        linewidth=3,
        linecolor="black",
        ticks="outside",
        tickwidth=2,
        ticklen=10,
        showgrid=True,
        gridcolor="#D9D9D9"
    )

    fig.update_yaxes(
        title=f"{selected_component} ({units})",
        title_font=dict(size=28),
        tickfont=dict(size=20),
        showline=True,
        linewidth=3,
        linecolor="black",
        ticks="outside",
        tickwidth=2,
        ticklen=10,
        showgrid=True,
        gridcolor="#D9D9D9"
    )

    return fig
    
def prepare_figure_for_png(
    source_fig,
    width=2400,
    height=1350
):
    """
    Return a separate Plotly figure configured so all titles,
    axes and legend entries remain inside the PNG canvas.
    """
    export_figure = go.Figure(source_fig)

    # Use a horizontal legend inside the image.
    # Long legends wrap onto additional rows rather than
    # extending beyond the right edge.
    legend_entries = sum(
        1
        for trace in export_figure.data
        if trace.showlegend is not False
        and getattr(trace, "name", None)
    )

    # Increase top margin when there are many legend entries.
    legend_rows = max(
        1,
        int(np.ceil(legend_entries / 4))
    )

    top_margin = min(
        500,
        150 + (legend_rows * 55)
    )

    export_figure.update_layout(
        autosize=False,
        width=width,
        height=height,
        template="simple_white",
        paper_bgcolor="white",
        plot_bgcolor="white",

        # Keep the title inside the canvas.
        title=dict(
            x=0.01,
            xanchor="left",
            y=0.98,
            yanchor="top",
            font=dict(size=30)
        ),

        # Keep the legend within the PNG boundaries.
        legend=dict(
            orientation="h",
            x=0,
            xanchor="left",
            y=1.02,
            yanchor="bottom",
            font=dict(size=16),
            title=dict(
                font=dict(size=16)
            ),
            entrywidth=0.24,
            entrywidthmode="fraction"
        ),

        margin=dict(
            l=140,
            r=80,
            t=top_margin,
            b=130,
            pad=10,
            autoexpand=True
        ),

        font=dict(
            family="Arial",
            size=18,
            color="black"
        )
    )

    # Allow Plotly to make extra space for long axis titles
    # and tick-label text.
    export_figure.update_xaxes(
        automargin=True,
        title_standoff=25
    )

    export_figure.update_yaxes(
        automargin=True,
        title_standoff=25
    )

    return export_figure    
    
@st.cache_data
def load_default_pcv():
    return pd.read_excel(
        "PCVs.xlsb",
        engine="pyxlsb"
    )


@st.cache_data
def load_default_pfas():
    return pd.read_excel(
        "PCVs.xlsb",
        sheet_name="PFAS",
        engine="pyxlsb"
    )


@st.cache_data
def load_default_pesticides():
    return pd.read_excel(
        "PCVs.xlsb",
        sheet_name="Pesticides",
        engine="pyxlsb"
    )
    
@st.cache_data
def load_excel_file(file_bytes):

    loaded = {}

    excel_file = pd.ExcelFile(
        BytesIO(file_bytes)
    )

    for sheet in excel_file.sheet_names:

        loaded[sheet] = pd.read_excel(
            excel_file,
            sheet_name=sheet
        )

    return loaded


@st.cache_data
def load_csv_file(file_bytes):

    return pd.read_csv(
        BytesIO(file_bytes)
    )
    
import re
    
def clean_component_name(text):
    text = str(text).strip().lower()

    # Standardise micro symbols
    text = (
        text
        .replace("µ", "u")
        .replace("μ", "u")
    )

    # Remove any bracketed information:
    # Copper (mg/l) -> Copper
    text = re.sub(
        r"\([^)]*\)",
        "",
        text
    )

    # Remove concentration units at the end:
    # Copper mg/l -> Copper
    # Copper mg l -> Copper
    # Copper ug/L -> Copper
    # Copper micrograms/l -> Copper
    text = re.sub(
        r"""
        \s+
        (?:
            mg |
            ug |
            ng |
            g |
            micrograms? |
            milligrams? |
            nanograms?
        )
        \s*
        (?:/|\s+)
        \s*
        (?:
            l |
            litre |
            litres |
            liter |
            liters
        )
        \s*$
        """,
        "",
        text,
        flags=re.IGNORECASE | re.VERBOSE
    )

    # Existing general normalisation
    text = text.replace("/", " ")
    text = text.replace("-", " ")
    text = " ".join(text.split())

    return text.strip() 

@st.cache_data
def build_limit_lookup(pcv_df):

    lookup = {}

    for _, row in pcv_df.iterrows():

        parameter = clean_component_name(
            row.get("Parameter", "")
        )

        row_dict = row.to_dict()

        lookup.setdefault(
            parameter,
            []
        ).append(row_dict)

        aliases = str(
            row.get(
                "Parameter Alias(es)",
                ""
            )
        )

        for alias in aliases.split(","):

            alias = clean_component_name(alias)

            if alias:

                lookup.setdefault(
                    alias,
                    []
                ).append(row_dict)

    return lookup            
    
@st.cache_data
def add_component_clean(df, det_col):

    df = df.copy()

    df["Component_Clean"] = (
        df[det_col]
        .astype(str)
        .apply(clean_component_name)
    )

    return df    

@st.cache_data
def clean_dataset(
    df,
    date_col,
    det_col,
    result_col,
    site_col,
    operator_col=None
):

    df = df.copy()

    df[date_col] = pd.to_datetime(
        df[date_col],
        errors="coerce"
    )

    df["Original_Result"] = (
        df[result_col]
        .astype(str)
    )

    # Default
    df["Non_Detect"] = False

    # Use separate operator column if available
    if (
        operator_col is not None
        and operator_col != "No Operator Column"
        and operator_col in df.columns
    ):

        NON_DETECT_FLAGS = {
            "<",
            "<=",
            "ND",
            "LOD",
            "BDL",
            "U",
            "UJ",
            "LT"
        }

        df["Non_Detect"] = (
            df[operator_col]
            .astype(str)
            .str.upper()
            .str.strip()
            .isin(NON_DETECT_FLAGS)
        )

    # Otherwise use existing behaviour
    else:

        df["Non_Detect"] = (
            df["Original_Result"]
            .str.strip()
            .str.startswith("<")
        )

    df[result_col] = (
        df["Original_Result"]
        .str.replace("<=", "", regex=False)
        .str.replace("<", "", regex=False)
    )

    df[result_col] = pd.to_numeric(
        df[result_col],
        errors="coerce"
    )

    # Half LOD substitution
    df.loc[
        df["Non_Detect"],
        result_col
    ] = (
        df.loc[
            df["Non_Detect"],
            result_col
        ] / 2
    )

    required_cols = [
        date_col,
        det_col,
        result_col
    ]

    if site_col is not None:
        required_cols.append(site_col)

    df = df.dropna(
        subset=required_cols
    )

    ph_mask = (
        df[det_col]
        .astype(str)
        .str.contains(
            "Hydrogen",
            case=False,
            na=False
        )
    )

    df = df[
        ~(ph_mask & (df[result_col] <= 0))
    ]

    df = df[
        df[result_col] > 0
    ]

    df = add_component_clean(
        df,
        det_col
    )

    return df
    
@st.cache_data
def build_compliance_summary(
    df,
    site_col,
    det_col,
    unit_col,
    result_col,
    selected_regulations,
    limit_lookup
):

    summary_rows = []

    for source in sorted(
        df[site_col].dropna().unique()
    ):

        source_df = df[
            df[site_col] == source
        ]

        for component in sorted(
            source_df["Component_Clean"]
            .dropna()
            .unique()
        ):

            component_df = source_df[
                source_df["Component_Clean"]
                == component
            ]

            results = pd.to_numeric(
                component_df[result_col],
                errors="coerce"
            ).dropna()

            if len(results) == 0:
                continue

            breaches = 0
            limit_value = None
            poc = ""
            regulation_source = ""

            limits = pd.DataFrame(
                limit_lookup.get(
                    component,
                    []
                )
            )

            if (
                not limits.empty
                and "Regulatory Source" in limits.columns
            ):
                limits = limits[
                    limits["Regulatory Source"]
                    .isin(selected_regulations)
                ]

            if not limits.empty:

                numeric_limits = pd.to_numeric(
                    limits["Limit Value"],
                    errors="coerce"
                )

                if numeric_limits.notna().any():

                    limit_idx = numeric_limits.idxmin()

                    limit_row = limits.loc[
                        limit_idx
                    ]

                    limit_value = numeric_limits.min()

                    poc = limit_row.get(
                        "Point of Compliance",
                        ""
                    )

                    regulation_source = limit_row.get(
                        "Regulatory Source",
                        ""
                    )

                    breaches = (
                        results > limit_value
                    ).sum()

            units = ""

            if (
                unit_col
                and unit_col in component_df.columns
            ):

                unit_values = (
                    component_df[unit_col]
                    .dropna()
                    .astype(str)
                )

                if not unit_values.empty:
                    units = unit_values.mode().iloc[0]

            display_name = (
                component_df[det_col]
                .mode()
                .iloc[0]
            )

            summary_rows.append({

                "Water Source": source,
                "Component": display_name,
                "Units": units,
                "Samples": len(results),
                "Minimum": round(results.min(), 3),
                "Mean": round(results.mean(), 3),
                "P50": round(results.quantile(0.50), 3),
                "P95": round(results.quantile(0.95), 3),
                "Maximum": round(results.max(), 3),
                "Limit Source": regulation_source,
                "Limit": limit_value,
                "Point of Compliance": poc,
                "Breaches": breaches,
                "% Exceeding": round(
                    100 * breaches / len(results),
                    2
                )
            })

    return pd.DataFrame(summary_rows)
    
@st.cache_data
def build_pfas_pesticide_summary(
    site_df,
    date_col,
    result_col,
    pfas_components,
    pesticide_components
):

    pfas_results = site_df[
        site_df["Component_Clean"]
        .isin(pfas_components)
    ].copy()

    pesticide_results = site_df[
        site_df["Component_Clean"]
        .isin(pesticide_components)
    ].copy()

    pfas_totals = (
        pfas_results
        .groupby(date_col)[result_col]
        .sum()
        .reset_index()
        .rename(
            columns={
                result_col: "PFAS Total"
            }
        )
    )

    pesticide_totals = (
        pesticide_results
        .groupby(date_col)[result_col]
        .sum()
        .reset_index()
        .rename(
            columns={
                result_col: "Pesticides Total"
            }
        )
    )

    return (
        pfas_results,
        pesticide_results,
        pfas_totals,
        pesticide_totals
    )
 
# =====================================================
# PROJECT SETUP
# =====================================================

loaded_datasets = {}

project_setup = st.expander(
    "Project Setup & Sheet Selection",
    expanded=True
)

with project_setup:

    setup1, setup2, setup3 = st.columns([2, 2, 3])

    with setup1:
        uploaded_files = st.file_uploader(
            "Water Quality Dataset(s)",
            type=["xlsx", "csv"],
            accept_multiple_files=True
        )

    with setup2:
        pcv_file = st.file_uploader(
            "PCV Database",
            type=["xlsx", "xlsb", "csv"]
        )       
        
if uploaded_files:
    
    for uploaded_file in uploaded_files:

        try:

            if uploaded_file.name.lower().endswith(".csv"):

                loaded_datasets[
                    uploaded_file.name
                ] = load_csv_file(
                    uploaded_file.getvalue()
                )

            else:

                sheets = load_excel_file(
                    uploaded_file.getvalue()
                )

                for sheet, df_sheet in sheets.items():

                    loaded_datasets[
                        f"{uploaded_file.name} | {sheet}"
                    ] = df_sheet
        except Exception as e:

            st.error(
                f"Error loading {uploaded_file.name}: {e}"
            )

    with setup3:

        selected_dataset = st.selectbox(
            "Main Dataset",
            list(loaded_datasets.keys())
        ) 

    # ================================================
    # LOAD PCV DATABASE
    # ================================================

    try:

        # User uploaded a replacement database
        if pcv_file is not None:

            if pcv_file.name.lower().endswith(".csv"):

                pcv_df = pd.read_csv(pcv_file)

            elif pcv_file.name.lower().endswith(".xlsb"):

                pcv_df = pd.read_excel(
                    pcv_file,
                    engine="pyxlsb"
                )      

            else:

                pcv_df = pd.read_excel(pcv_file)

        # Otherwise use bundled database
        else:

            pcv_df = load_default_pcv()  

    except Exception as e:

        st.sidebar.error(
            f"Could not load PCV database: {e}"
        )

        pcv_df = None
        
    if pcv_df is not None:

        pcv_df["Limit Value"] = pd.to_numeric(
            pcv_df["Limit Value"],
            errors="coerce"
        )
        
        limit_lookup = build_limit_lookup(
            pcv_df
        )
        
        
        try:

            if pcv_file is not None:

                pfas_df = pd.read_excel(
                    pcv_file,
                    sheet_name="PFAS",
                    engine="pyxlsb"
                )

                pesticide_df = pd.read_excel(
                    pcv_file,
                    sheet_name="Pesticides",
                    engine="pyxlsb"
                )

            else:

                pfas_df = load_default_pfas()

                pesticide_df = load_default_pesticides()

        except Exception as e:

            st.error(
                f"PFAS/Pesticide sheet load failed: {e}"
            )

            pfas_df = pd.DataFrame()
            pesticide_df = pd.DataFrame()

    # =====================================================
    # COLUMN AUTO-DETECTION
    # =====================================================

    COLUMN_ALIASES = {

        "DATE": [
            "Sampled Date",
            "Sample Date",
            "Date",
            "Date Collected",
            "Collection Date"
        ],

        "SITE": [
            "Description",
            "Site",
            "Location",
            "Sampling Point",
            "Site Name",
            "Water Source",
            "Sample Point",
            "Sample Point Name"
        ],

        "PARAMETER": [
            "Name (Component)",
            "Component",
            "Parameter",
            "Analyte",
            "Determinand",
            "Param Units"
        ],

        "UNIT": [
            "Units",
            "Unit"
        ],

        "RESULT": [
            "Result Value (Numeric)",
            "Result",
            "Value",
            "Concentration",
            "Numeric Result",
            "Auth Result"
        ],
        
        "OPERATOR": [
            "Operator",
            "Qualifier",
            "Result Operator",
            "Result Type",
            "Analysis Qualifier",
            "Qualifier Code",
            "Relation",
            "Condition",
            "Qu.",
            "Qu"
        ]
    }


    def find_column(df, aliases):

        for col in df.columns:

            if col.strip().lower() in [
                a.strip().lower()
                for a in aliases
            ]:

                return col

        return None
        
    def normalise_unit(unit):

        unit = (
            str(unit)
            .strip()
            .lower()
            .replace("µ", "u")
        )

        # remove analyte names

        if "mg" in unit:
            return "mg/l"

        if "ug" in unit:
            return "ug/l"

        if "ng" in unit:
            return "ng/l"

        return unit
        
    def add_auto_unit_column(df, component_col):
        """
        Extract concentration units from component names.

        Examples:
            Iron (ug Fe/l)       -> ug/l
            Iron (µg Fe/L)       -> ug/l
            Aluminium (mg Al/l)  -> mg/l
            Lead (ug/l)          -> ug/l
            Copper mg Cu/l       -> mg/l
        """
        df = df.copy()
        df["Auto_Unit"] = ""

        if (
            component_col is None
            or component_col not in df.columns
        ):
            return df

        parameter_text = (
            df[component_col]
            .fillna("")
            .astype(str)
            .str.replace("µ", "u", regex=False)
            .str.replace("μ", "u", regex=False)
        )

        # Captures the concentration prefix only: mg, ug, ng or g.
        #
        # Also permits optional text between the prefix and /l:
        #   ug Fe/l
        #   mg Al/l
        #   ug as Pb/l
        #   ug/l
        bracketed_unit = parameter_text.str.extract(
            r"""
            \(
            \s*
            (mg|ug|ng|g)
            \s*
            (?:
                (?:as\s+)?
                [A-Za-z][A-Za-z0-9+\-]*
                \s*
            )?
            /
            \s*
            (?:l|litre|liter)
            \s*
            \)
            """,
            flags=re.IGNORECASE | re.VERBOSE,
            expand=False
        )

        trailing_unit = parameter_text.str.extract(
            r"""
            \s+
            (mg|ug|ng|g)
            \s*
            (?:
                (?:as\s+)?
                [A-Za-z][A-Za-z0-9+\-]*
                \s*
            )?
            /
            \s*
            (?:l|litre|liter)
            \s*
            $
            """,
            flags=re.IGNORECASE | re.VERBOSE,
            expand=False
        )

        df["Auto_Unit"] = (
            bracketed_unit
            .fillna(trailing_unit)
            .fillna("")
            .str.lower()
            .str.strip()
        )

        # The regex captures only the concentration prefix,
        # so restore the standard denominator.
        has_unit = df["Auto_Unit"] != ""

        df.loc[
            has_unit,
            "Auto_Unit"
        ] = (
            df.loc[
                has_unit,
                "Auto_Unit"
            ]
            + "/l"
        )

        return df
    
    def clean_unit_value(unit):
        """Standardise unit text while retaining non-concentration units."""
        if unit is None or pd.isna(unit):
            return ""

        unit = str(unit).strip()
        replacements = {
            "µg/l": "ug/l",
            "μg/l": "ug/l",
            "ug/L": "ug/l",
            "mg/L": "mg/l",
            "ng/L": "ng/l",
            "g/L": "g/l"
        }
        return replacements.get(unit, unit)


    def get_mode_unit(data, unit_col="Resolved_Unit"):
        """Return the most common non-blank unit in a dataframe."""
        if data is None or data.empty or unit_col not in data.columns:
            return ""

        values = data[unit_col].fillna("").astype(str).str.strip()
        values = values[values != ""]
        if values.empty:
            return ""
        return clean_unit_value(values.mode().iloc[0])


    def value_with_unit(value, unit, decimals=3):
        """Format a numeric value consistently for reports and tables."""
        if pd.isna(value):
            return ""
        formatted = f"{value:.{decimals}f}"
        return f"{formatted} {unit}" if unit else formatted


    def unit_scale(unit):

        unit = normalise_unit(unit)

        scales = {
            "mg/l": 1,
            "ug/l": 0.001,
            "ng/l": 0.000001
        }

        return scales.get(unit)


    def convert_units(value, from_unit, to_unit):

        if pd.isna(value):
            return value

        from_scale = unit_scale(from_unit)
        to_scale = unit_scale(to_unit)

        if from_scale is None:
            return value

        if to_scale is None:
            return value

        return value * (from_scale / to_scale)


    def get_converted_limit(limit_row, sample_unit):

        raw_limit = pd.to_numeric(
            limit_row.get("Limit Value"),
            errors="coerce"
        )

        if pd.isna(raw_limit):
            return np.nan

        limit_unit = limit_row.get(
            "Units",
            ""
        )

        return convert_units(
            raw_limit,
            limit_unit,
            sample_unit
        )

    PFAS_COMPONENTS = set()
    PESTICIDE_COMPONENTS = set()

    PFAS_LIMIT = np.nan
    PESTICIDE_LIMIT = np.nan

    PFAS_LIMIT_UNIT = ""
    PESTICIDE_LIMIT_UNIT = ""

    PFAS_COMPANY_LIMITS = {}
    PESTICIDE_COMPANY_LIMITS = {}

    if not pfas_df.empty:

        PFAS_COMPONENTS = {
            clean_component_name(x)
            for x in pfas_df.iloc[:, 0].dropna()
            if str(x).lower() != "total limit"
        }

        try:
            limit_row = pfas_df.loc[
                pfas_df.iloc[:,0]
                .astype(str)
                .str.lower()
                == "total limit"
            ].iloc[0]

            PFAS_LIMIT = float(limit_row.iloc[1])
            PFAS_LIMIT_UNIT = str(limit_row.iloc[2])
            
            found_total = False

            for _, row in pfas_df.iterrows():

                name = str(row.iloc[0]).strip()

                if name.lower() == "total limit":
                    found_total = True
                    continue

                if not found_total:
                    continue

                if pd.isna(row.iloc[0]):
                    continue

                try:
                    PFAS_COMPANY_LIMITS[name] = {
                        "limit": float(row.iloc[1]),
                        "unit": str(row.iloc[2])
                    }
                except:
                    pass

        except:
            pass

    if not pesticide_df.empty:

        PESTICIDE_COMPONENTS = {
            clean_component_name(x)
            for x in pesticide_df.iloc[:, 0].dropna()
            if str(x).lower() != "total limit"
        }

        try:
            limit_row = pesticide_df.loc[
                pesticide_df.iloc[:,0]
                .astype(str)
                .str.lower()
                == "total limit"
            ].iloc[0]

            PESTICIDE_LIMIT = float(limit_row.iloc[1])
            PESTICIDE_LIMIT_UNIT = str(limit_row.iloc[2])
            
            found_total = False

            for _, row in pesticide_df.iterrows():

                name = str(row.iloc[0]).strip()

                if name.lower() == "total limit":
                    found_total = True
                    continue

                if not found_total:
                    continue

                if pd.isna(row.iloc[0]):
                    continue

                try:
                    PESTICIDE_COMPANY_LIMITS[name] = {
                        "limit": float(row.iloc[1]),
                        "unit": str(row.iloc[2])
                    }
                except:
                    pass          

        except:
            pass

    def get_component_limits(pcv_df, selected_component):

        component = clean_component_name(selected_component)

        def row_matches(row):

            parameter = clean_component_name(
                row.get("Parameter", "")
            )

            aliases = str(
                row.get("Parameter Alias(es)", "")
            )

            alias_list = [
                clean_component_name(a)
                for a in aliases.split(",")
                if str(a).strip()
            ]

            # Exact parameter match
            if parameter == component:
                return True

            # Exact alias match
            if component in alias_list:
                return True

            return False

        return pcv_df[
            pcv_df.apply(row_matches, axis=1)
        ]

    df = loaded_datasets[
        selected_dataset
    ].copy()
    
    DATE_COL = find_column(
        df,
        COLUMN_ALIASES["DATE"]
    )

    SITE_COL = find_column(
        df,
        COLUMN_ALIASES["SITE"]
    )

    DET_COL = find_column(
        df,
        COLUMN_ALIASES["PARAMETER"]
    )

    UNIT_COL = find_column(
        df,
        COLUMN_ALIASES["UNIT"]
    )

    RESULT_COL = find_column(
        df,
        COLUMN_ALIASES["RESULT"]
    )
    
    OPERATOR_COL = find_column(
        df,
        COLUMN_ALIASES["OPERATOR"]
    )

    # If no unit column exists, extract units from parameter names
    if UNIT_COL is None:

        df["Auto_Unit"] = ""

        if (
            DET_COL is not None
            and pd.notna(DET_COL)
            and DET_COL in df.columns
        ):

            parameter_text = (
                df[DET_COL]
                .fillna("")
                .astype(str)
                .str.replace("µ", "u", regex=False)
                .str.replace("μ", "u", regex=False)
            )

            # Extracts the concentration prefix from component names such as:
            # Iron (ug Fe/l)       -> ug
            # Aluminium (mg Al/l)  -> mg
            # Lead (ug/l)          -> ug
            # Copper ug Cu/l       -> ug
            bracketed_unit = parameter_text.str.extract(
                r"""
                \(
                \s*
                (mg|ug|ng|g)
                \s*
                (?:
                    (?:as\s+)?
                    [A-Za-z][A-Za-z0-9+\-]*
                    \s*
                )?
                /
                \s*
                (?:l|litre|liter)
                \s*
                \)
                """,
                flags=re.IGNORECASE | re.VERBOSE,
                expand=False
            )

            trailing_unit = parameter_text.str.extract(
                r"""
                \s+
                (mg|ug|ng|g)
                \s*
                (?:
                    (?:as\s+)?
                    [A-Za-z][A-Za-z0-9+\-]*
                    \s*
                )?
                /
                \s*
                (?:l|litre|liter)
                \s*
                $
                """,
                flags=re.IGNORECASE | re.VERBOSE,
                expand=False
            )

            df["Auto_Unit"] = (
                bracketed_unit
                .fillna(trailing_unit)
                .fillna("")
                .str.lower()
                .str.strip()
            )

            # Add the denominator back to extracted prefixes.
            has_auto_unit = df["Auto_Unit"].ne("")

            df.loc[
                has_auto_unit,
                "Auto_Unit"
            ] = (
                df.loc[
                    has_auto_unit,
                    "Auto_Unit"
                ]
                + "/l"
            )

        UNIT_COL = "Auto_Unit"

    original_count = len(df)
    
    # =====================================================
    # MANUAL COLUMN OVERRIDE
    # =====================================================

    with st.expander(
        "Column Mapping (main statistics/figure)"
    ):
        
        dataset_layout = st.radio(
            "Dataset Layout",
            [
                "Long Format",
                "Wide Format"
            ]
        )
        
        first_component_col = None

        if dataset_layout == "Wide Format":

            first_component_col = st.selectbox(
                "First Component Column",
                df.columns
            )

        DATE_COL = st.selectbox(
            "Date Column",
            df.columns,
            index=df.columns.get_loc(DATE_COL)
            if DATE_COL else 0
        )

        site_options = ["No Site Column"] + list(df.columns)

        SITE_COL = st.selectbox(
            "Site Column",
            site_options,
            index=(
                site_options.index(SITE_COL)
                if SITE_COL in site_options
                else 0
            )
        )

        if SITE_COL == "No Site Column":
            SITE_COL = None

        if dataset_layout == "Long Format":

            DET_COL = st.selectbox(
                "Component Column",
                df.columns,
                index=df.columns.get_loc(DET_COL)
                if DET_COL else 0
            )

            RESULT_COL = st.selectbox(
                "Result Column",
                df.columns,
                index=df.columns.get_loc(RESULT_COL)
                if RESULT_COL else 0
            )

        else:

            st.info(
                "Wide format selected. Component and Result columns will be created automatically."
            )


        unit_options = ["No Unit Column"] + list(df.columns)

        UNIT_COL = st.selectbox(
            "Unit Column",
            unit_options,
            index=(
                unit_options.index(UNIT_COL)
                if UNIT_COL in unit_options
                else 0
            )
        )
        
        operator_options = ["No Operator Column"] + list(df.columns)

        OPERATOR_COL = st.selectbox(
            "Operator Column",
            operator_options,
            index=(
                operator_options.index(OPERATOR_COL)
                if OPERATOR_COL in operator_options
                else 0
            )
        )
         
    # =====================================================
    # WIDE FORMAT CONVERSION
    # =====================================================

    if dataset_layout == "Wide Format":

        start_idx = df.columns.get_loc(
            first_component_col
        )

        component_cols = list(
            df.columns[start_idx:]
        )

        id_cols = list(
            df.columns[:start_idx]
        )

        df = df.melt(
            id_vars=id_cols,
            value_vars=component_cols,
            var_name="Component",
            value_name="Result"
        )

        DET_COL = "Component"
        RESULT_COL = "Result"

        if SITE_COL is not None and SITE_COL not in df.columns:
            SITE_COL = None

        UNIT_COL = None
        OPERATOR_COL = None         
    
    required = {
        "Date": DATE_COL,
        "Component": DET_COL,
        "Result": RESULT_COL
    }

    missing = [
        k
        for k, v in required.items()
        if v is None
    ]

    if missing:

        st.error(
            "Missing required columns: "
            + ", ".join(missing)
        )

        st.write(
            "Columns detected:"
        )

        st.write(
            df.columns.tolist()
        )

        st.stop()
        
    if SITE_COL is None:
        SITE_COL = "__SITE__"
        df[SITE_COL] = uploaded_file.name.rsplit(".", 1)[0]
    
    # =================================================
    # CLEANING
    # =================================================

    before_cleaning = len(df)

    df = clean_dataset(
        df,
        DATE_COL,
        DET_COL,
        RESULT_COL,
        SITE_COL,
        OPERATOR_COL
    )

    # One authoritative unit column used throughout the app.
    df["Resolved_Unit"] = ""
    if (
        UNIT_COL is not None
        and UNIT_COL != "No Unit Column"
        and UNIT_COL in df.columns
    ):
        df["Resolved_Unit"] = (
            df[UNIT_COL]
            .fillna("")
            .astype(str)
            .apply(clean_unit_value)
        )

    if "manual_component_units" not in st.session_state:
        st.session_state.manual_component_units = {}

    for component_key, manual_unit in st.session_state.manual_component_units.items():
        if manual_unit:
            df.loc[
                df["Component_Clean"] == component_key,
                "Resolved_Unit"
            ] = manual_unit

    UNIT_COL = "Resolved_Unit"

    cleaned_count = len(df)

    qa_removed = (
        before_cleaning -
        cleaned_count
    )

    zero_removed = 0

    # =================================================
    # SIDEBAR
    # =================================================

    st.sidebar.header("Filters")

    selected_site = st.sidebar.selectbox(
        "Source",
        sorted(df[SITE_COL].unique())
    )

    filtered = df[
        df[SITE_COL] == selected_site
    ].copy()

    component_lookup = (
        filtered.groupby("Component_Clean")[DET_COL]
        .agg(
            lambda x: min(
                x.dropna().astype(str),
                key=lambda value: (
                    len(clean_component_name(value)),
                    len(value)
                )
            )
        )
        .to_dict()
    )

    display_names = sorted(component_lookup.values())

    selected_display = st.sidebar.selectbox(
        "Component",
        display_names
    )

    selected_component = clean_component_name(
        selected_display
    )

    st.sidebar.markdown("---")
    st.sidebar.header("Limits and Units")
    st.sidebar.subheader("Limits")

    selected_regulations = st.sidebar.multiselect(
        "Regulations",
        [
            "WS(WQ) 2016",
            "PWS (Scotland) 2014",
            "EU rDWD",
            "WHO",
            "US EPA NPDWR",
            "Anglian Water",
            "Wessex Water",
            "Severn Trent (Boughton)",
            "South East Water"
        ],
        default=["WS(WQ) 2016"]
    )

    use_manual_pcvs = st.sidebar.checkbox(
        "Add Manual Limit Entries",
        value=False,
        help=(
            "Manually add component limits when a component name has not "
            "been matched automatically."
        )
    )

    manual_regulations = []
    manual_pcvs = []

    if use_manual_pcvs:
        manual_regulations = st.sidebar.multiselect(
            "Regulatory Source(s)",
            sorted(
                pcv_df["Regulatory Source"]
                .dropna()
                .unique()
            )
        )

        for regulation in manual_regulations:
            regulation_pcvs = sorted(
                pcv_df.loc[
                    pcv_df["Regulatory Source"] == regulation,
                    "Parameter"
                ]
                .dropna()
                .unique()
            )

            selected_pcvs = st.sidebar.multiselect(
                f"{regulation} Parameter(s)",
                regulation_pcvs,
                key=f"manual_pcv_{regulation}"
            )
            manual_pcvs.extend(selected_pcvs)

    st.sidebar.subheader("Units")

    existing_component_rows = df[
        df["Component_Clean"] == selected_component
    ]
    detected_component_unit = get_mode_unit(
        existing_component_rows,
        "Resolved_Unit"
    )

    use_manual_component_unit = st.sidebar.checkbox(
        "Set unit for selected component",
        value=False,
        help=(
            "Enable this to assign or override the unit for the selected "
            "component across figures, tables, compliance calculations "
            "and exported reports."
        ),
        key=f"use_manual_unit_{selected_component}"
    )

    selected_manual_unit = ""

    if use_manual_component_unit:
        common_unit_options = [
            "mg/l", "ug/l", "ng/l", "g/l", "NTU", "Hazen",
            "uS/cm", "mS/cm", "pH units", "degrees C", "%",
            "CFU/100 ml", "MPN/100 ml", "Custom..."
        ]

        saved_component_unit = st.session_state.manual_component_units.get(
            selected_component,
            detected_component_unit
        )

        if saved_component_unit and saved_component_unit not in common_unit_options:
            default_unit_option = "Custom..."
        elif saved_component_unit in common_unit_options:
            default_unit_option = saved_component_unit
        else:
            default_unit_option = common_unit_options[0]

        selected_unit_option = st.sidebar.selectbox(
            "Unit for selected component",
            common_unit_options,
            index=common_unit_options.index(default_unit_option),
            help="Select the unit that applies to the currently selected component.",
            key=f"unit_select_{selected_component}"
        )

        if selected_unit_option == "Custom...":
            selected_manual_unit = st.sidebar.text_input(
                "Custom unit",
                value=(
                    saved_component_unit
                    if default_unit_option == "Custom..."
                    else ""
                ),
                key=f"custom_unit_{selected_component}"
            ).strip()
        else:
            selected_manual_unit = selected_unit_option

        if selected_manual_unit:
            selected_manual_unit = clean_unit_value(selected_manual_unit)
            st.session_state.manual_component_units[
                selected_component
            ] = selected_manual_unit
            df.loc[
                df["Component_Clean"] == selected_component,
                "Resolved_Unit"
            ] = selected_manual_unit
            st.sidebar.caption(
                f"Using {selected_manual_unit} for {selected_display}"
            )
    elif detected_component_unit:
        st.sidebar.caption(
            f"Detected unit: {detected_component_unit}"
        )
    else:
        st.sidebar.caption(
            "No unit detected. Enable the checkbox above to assign one."
        )

    # Rebuild this view after applying the component-level override.
    filtered = df[df[SITE_COL] == selected_site].copy()
    
    st.sidebar.markdown("---")
    st.sidebar.header("Time")


    # Date limits for the currently selected water source.
    min_date = filtered[DATE_COL].min().date()
    max_date = filtered[DATE_COL].max().date()

    quick_period = st.sidebar.radio(
        "Quick Period",
        [
            "All Data",
            "Last 12 Months",
            "Last 3 Years",
            "Last 5 Years",
            "Custom"
        ],
        index=0,
        key="quick_period"
    )

    # Only show and use manual date inputs when Custom is selected.
    if quick_period == "Custom":

        start_date = st.sidebar.date_input(
            "Start Date",
            value=min_date,
            min_value=min_date,
            max_value=max_date,
            key="custom_start_date"
        )

        end_date = st.sidebar.date_input(
            "End Date",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="custom_end_date"
        )

        if start_date > end_date:
            st.sidebar.error(
                "Start Date must be on or before End Date."
            )
            st.stop()

    elif quick_period == "Last 12 Months":

        end_date = max_date
        start_date = max(
            min_date,
            (
                pd.Timestamp(max_date)
                - pd.DateOffset(months=12)
            ).date()
        )

    elif quick_period == "Last 3 Years":

        end_date = max_date
        start_date = max(
            min_date,
            (
                pd.Timestamp(max_date)
                - pd.DateOffset(years=3)
            ).date()
        )

    elif quick_period == "Last 5 Years":

        end_date = max_date
        start_date = max(
            min_date,
            (
                pd.Timestamp(max_date)
                - pd.DateOffset(years=5)
            ).date()
        )

    else:
        # All Data always uses the full available dataset period.
        start_date = min_date
        end_date = max_date

    if quick_period != "Custom":
        st.sidebar.caption(
            f"Using {start_date:%d %b %Y} to "
            f"{end_date:%d %b %Y}"
        )

    filtered = filtered[
        filtered["Component_Clean"] == selected_component
    ].copy()

    # Include every result occurring on the end date, including timestamps.
    filter_start = pd.to_datetime(start_date)
    filter_end_exclusive = (
        pd.to_datetime(end_date)
        + pd.Timedelta(days=1)
    )

    filtered = filtered[
        (filtered[DATE_COL] >= filter_start)
        &
        (filtered[DATE_COL] < filter_end_exclusive)
    ].copy()
            
    sample_unit = get_mode_unit(
        filtered,
        "Resolved_Unit"
    )

    component_limits = pd.DataFrame()

    if pcv_df is not None:

        # automatic match
        component_limits = pd.DataFrame(
            limit_lookup.get(
                selected_component,
                []
            )
        )

        # manual additions
        if manual_pcvs:

            manual_limits = pcv_df[
                pcv_df["Parameter"]
                .isin(manual_pcvs)
            ]

            component_limits = pd.concat(
                [
                    component_limits,
                    manual_limits
                ],
                ignore_index=True
            )

        component_limits = component_limits.drop_duplicates()

        # allow both main regulations and manually selected regulation groups
        allowed_regulations = set(selected_regulations)

        if manual_regulations:
            allowed_regulations.update(
                manual_regulations
            )

        if (
            not component_limits.empty
            and "Regulatory Source" in component_limits.columns
        ):
            component_limits = component_limits[
                component_limits["Regulatory Source"]
                .isin(allowed_regulations)
            ]    
            
    # =================================================
    # SUMMARY
    # =================================================

    monitoring_start = df[DATE_COL].min()
    monitoring_end = df[DATE_COL].max()

    c1, c2, c3, c4 = st.columns(
        [1,1,1,1.45]
    )

    c1.metric(
        "Results Imported",
        f"{original_count:,}"
    )

    c2.metric(
        "Components",
        df[DET_COL].nunique()
    )

    c3.metric(
        "Water Sources",
        df[SITE_COL].nunique()
    )

    # c4.metric(
        # "QA Removed",
        # f"{qa_removed:,}"
    # )

    # c5.metric(
        # "Zero Values Removed",
        # f"{zero_removed:,}"
    # )

    c4.metric(
        "Monitoring Period",
        f"{monitoring_start:%d %b %Y} - {monitoring_end:%d %b %Y}"
    )

    # =================================================
    # TABS
    # =================================================
    
    summary_df = pd.DataFrame()

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs(
        [
            "Component Data",
            "Component Statistics",
            "Component Figure",
            "Component Exceedances",
            "Custom Statistics",
            "Site Compliance Summary",
            "Source Comparison",
            "Summary Export",
            "PFAS / Pesticides"
        ]
    )

    # =================================================
    # DATA TAB
    # =================================================

    with tab1:

        display_cols = [
            DATE_COL,
            SITE_COL,
            DET_COL,
            RESULT_COL
        ]
        
        if (
            OPERATOR_COL is not None
            and OPERATOR_COL != "No Operator Column"
        ):
            display_cols.append(OPERATOR_COL)

        if (
            UNIT_COL is not None
            and UNIT_COL != "No Unit Column"
            and UNIT_COL in filtered.columns
        ):
            display_cols.insert(3, UNIT_COL)          

        st.dataframe(
            filtered[display_cols],
            use_container_width=True
        )

    # =================================================
    # STATISTICS TAB
    # =================================================

    with tab2:

        stats_df = statistics_table(
            filtered,
            RESULT_COL,
            selected_component,
            sample_unit
        )

        st.dataframe(
            stats_df,
            use_container_width=True
        )

        if not component_limits.empty:

            st.subheader(
                "Applicable Regulatory Limits"
            )

            st.dataframe(
                component_limits[
                    [
                        "Regulatory Source",
                        "Limit Value",
                        "Units",
                        "Point of Compliance"

                    ]
                ],
                use_container_width=True
            )

    # =================================================
    # FIGURE TAB
    # =================================================

    with tab3:

        mean = filtered[RESULT_COL].mean()
        p50 = filtered[RESULT_COL].quantile(0.50)
        p95 = filtered[RESULT_COL].quantile(0.95)
                        
        detect_df = filtered[~filtered["Non_Detect"]]
        nondetect_df = filtered[filtered["Non_Detect"]]

        fig = go.Figure()

        # = results
        fig.add_trace(
            go.Scatter(
                x=detect_df[DATE_COL],
                y=detect_df[RESULT_COL],
                mode="markers",
                name="Detected",
                marker=dict(
                    color="#005A9C",
                    size=14,
                    symbol="circle",
                    line=dict(
                        width=1,
                        color="#003B5C"
                    )
                )
            )
        )

        # < results
        fig.add_trace(
            go.Scatter(
                x=nondetect_df[DATE_COL],
                y=nondetect_df[RESULT_COL],
                mode="markers",
                name="Below LOD",
                marker=dict(
                    color="red",
                    size=7,
                    symbol="diamond",
                    line=dict(
                        width=0.5,
                        color="darkred"
                    )
                )
            )
        )


        fig.add_trace(
            go.Scatter(
                x=[filtered[DATE_COL].min(), filtered[DATE_COL].max()],
                y=[mean, mean],
                mode="lines",
                name=(
                    f"Mean ({mean:.2f} {sample_unit})"
                    if sample_unit
                    else f"Mean ({mean:.2f})"
                ),
                line=dict(
                    color="#000000",
                    dash="dot"
                ),
                hovertemplate=f"Mean: {mean:.2f}<extra></extra>"
            )
        )
        
        fig.add_trace(
            go.Scatter(
                x=[filtered[DATE_COL].min(), filtered[DATE_COL].max()],
                y=[p50, p50],
                mode="lines",
                name=(
                    f"P50 ({p50:.2f} {sample_unit})"
                    if sample_unit
                    else f"P50 ({p50:.2f})"
                ),
                line=dict(
                    color="#4B0082",
                    dash="dash"
                ),
                hovertemplate=f"P50: {p50:.2f}<extra></extra>"
            )
        )
        
        fig.add_trace(
            go.Scatter(
                x=[filtered[DATE_COL].min(), filtered[DATE_COL].max()],
                y=[p95, p95],
                mode="lines",
                name=(
                    f"P95 ({p95:.2f} {sample_unit})"
                    if sample_unit
                    else f"P95 ({p95:.2f})"
                ),
                line=dict(
                    color="#FF00FF",
                    dash="dash"
                ),
                hovertemplate=f"P95: {p95:.2f}<extra></extra>"
            )
        )

        if not component_limits.empty:

            colours = {
                "WS(WQ) 2016": "red",
                "EU rDWD": "blue",
                "WHO": "green",
                "Anglian Water": "orange",
                "Wessex Water": "black",
                "Severn Trent (Boughton)": "yellow"
            }
            
            colour_palette = [
                "#4DAF4A",  # green
                "#FF7F00",  # orange
                "#E41A1C",  # red
                "#377EB8",  # blue
                "#984EA3",  # purple                
                "#A65628",  # brown
                "#F781BF",  # pink
                "#999999",  # grey
            ]
            
            regulation_colour_map = {}

            for _, row in component_limits.iterrows():
                

                limit_value = get_converted_limit(
                    row,
                    sample_unit
                )              

                if pd.isna(limit_value):
                    continue

                poc = row.get("Point of Compliance", "")

                regulation = row.get("Regulatory Source", "")

                pcv_unit = str(
                    row.get("Units", "")
                )

                original_limit = row.get(
                    "Limit Value",
                    ""
                )

                display_unit = (
                    sample_unit
                    if sample_unit
                    else pcv_unit
                )
                                
                colour_key = (
                    str(regulation).strip(),
                    str(poc).strip()
                )

                if colour_key not in regulation_colour_map:
                    regulation_colour_map[colour_key] = colour_palette[
                        len(regulation_colour_map) % len(colour_palette)
                    ]

                line_colour = regulation_colour_map[colour_key]

                fig.add_trace(
                    go.Scatter(
                        x=[
                            filtered[DATE_COL].min(),
                            filtered[DATE_COL].max()
                        ],
                        y=[limit_value, limit_value],
                        mode="lines",
                        name=(
                            f"{regulation} - {poc} "
                            f"({limit_value:.3f} {display_unit})"
                        ),
                        line=dict(
                            color=line_colour,
                            width=3,
                            dash="dash"
                        ),
                        hovertemplate=(
                            f"{regulation} - {poc}<br>"
                            f"Original PCV: {original_limit} {pcv_unit}<br>"
                            f"Converted Limit: {limit_value:.3f} {display_unit}"
                            "<extra></extra>"
                        )
                    )
                )
                
                hover_text = (
                    f"<b>{row.get('Regulatory Source','')}</b><br>"
                    f"CAS Number: {row.get('CAS Number','N/A')}<br>"
                    f"Source Document: {row.get('Source Document','N/A')}<br>"
                    f"Source Section: {row.get('Source Section','N/A')}<br>"
                    f"Jurisdiction: {row.get('Jurisdiction','N/A')}<br>"
                    f"Point of Compliance: {row.get('Point of Compliance','N/A')}<br>"
                    f"Effective Date: {row.get('Effective Date','N/A')}<br>"
                    f"Original PCV: {original_limit} {pcv_unit}<br>"
                    f"Converted Limit: {limit_value:.3f} {display_unit}"
                )

                fig.add_trace(
                    go.Scatter(
                        x=[filtered[DATE_COL].max()],
                        y=[limit_value],
                        mode="markers",
                        marker=dict(
                            size=14,
                            color=line_colour,
                            opacity=0
                        ),
                        hovertemplate=hover_text + "<extra></extra>",
                        showlegend=False
                    )
                )


        fig.update_layout(

            template="simple_white",

            height=700,

            showlegend=True,
            
            legend=dict(
                title="Reported Result",
                x=1.02,
                y=1,
                font=dict(size=18)
            ),

            title=dict(
                text=f"<b>{selected_site.upper()}</b><br>{selected_display}",
                x=0.01,
                xanchor="left",
                font=dict(size=30)
            ),

            paper_bgcolor="white",
            plot_bgcolor="white",

            margin=dict(
                l=80,
                r=80,
                t=100,
                b=80
            )
        )

        fig.update_xaxes(
            title="Sample Date",
            range=[
                filtered[DATE_COL].min(),
                filtered[DATE_COL].max()
            ],
            title_font=dict(size=24),
            tickfont=dict(size=16),
            showline=True,
            linewidth=3,
            linecolor="black",
            ticks="outside",
            tickwidth=2,
            ticklen=8,
            showgrid=True,
            gridcolor="#D9D9D9"
        )

        fig.update_yaxes(
            title=(
                selected_component
                if sample_unit and f"({sample_unit})" in str(selected_component)
                else (
                    f"{selected_component} ({sample_unit})"
                    if sample_unit
                    else selected_component
                )
            ),
            title_font=dict(size=24),
            tickfont=dict(size=16),
            showline=True,
            linewidth=3,
            linecolor="black",
            ticks="outside",
            tickwidth=2,
            ticklen=8,
            showgrid=True,
            gridcolor="#D9D9D9"
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={
                "displaylogo": False,
                "toImageButtonOptions": {
                    "format": "png",
                    "filename": "water_quality_figure",
                    "width": 2400,
                    "height": 1350,
                    "scale": 2
                }
            }
        )
        
        st.info(
            "Click a legend item to hide/show a series.\n\n"
            "Double click a legend item to isolate it.\n\n"
            "Drag across the figure to zoom in on a certain section.\n\n"
            "Use toolbar in top right of the figure to pan, zoom, enable fullscreen, and export as a PNG. \n\n"
            "Double click on figure to reset zoom to original position."
        )

        # ---------------------------------------
        # HIGH QUALITY PNG EXPORT
        # ---------------------------------------

        export_fig = create_export_figure(
            filtered,
            selected_site,
            selected_component,
            DATE_COL,
            RESULT_COL,
            UNIT_COL
        )

        report_figures = {
            selected_component: fig
        }

        png_export_fig = prepare_figure_for_png(
            export_fig,
            width=2400,
            height=1350
        )    
                
    # =================================================
    # Limits/PCV
    # =================================================
        
    with tab4:

        if component_limits.empty:

            st.info(
                "No regulatory limits found for this component."
            )

        else:       

            for _, limit_row in component_limits.iterrows():

                limit_value = get_converted_limit(
                    limit_row,
                    sample_unit
                )
                
                poc = limit_row.get("Point of Compliance", "")
                regulation = limit_row["Regulatory Source"]  

                if pd.notna(limit_value):

                    exceedances = filtered[
                        filtered[RESULT_COL] > limit_value
                    ]

                else:

                    exceedances = pd.DataFrame()
                       
                st.subheader(
                    f"{selected_display} | "
                    f"{regulation} - {poc} "
                    f"({limit_value:.3f} {sample_unit})"
                )

                st.metric(
                    "Number of Exceedances",
                    len(exceedances)
                )

                if len(exceedances):

                    display_cols = [
                        SITE_COL,
                        DATE_COL,
                        DET_COL,
                        RESULT_COL
                    ]
                    
                    if UNIT_COL and UNIT_COL in exceedances.columns:
                        display_cols.append(UNIT_COL)

                    display_cols = [
                        c for c in display_cols
                        if c in exceedances.columns
                    ]

                    st.dataframe(
                        exceedances[display_cols],
                        use_container_width=True
                    )

                else:

                    st.success(
                        "No exceedances recorded."
                    )  


    # =================================================
    # CUSTOM SUMMARY BUILDER
    # =================================================

    with tab5:

        site_df = df[
            (df[SITE_COL] == selected_site)
            &
            (df[DATE_COL] >= pd.to_datetime(start_date))
            &
            (df[DATE_COL] < filter_end_exclusive)
        ].copy()
  
        available_components = sorted(
            site_df["Component_Clean"]
            .dropna()
            .unique()
        )

        st.subheader("Preset Tables")

        generate_envelope = st.button(
            "Generate AWTM Input Table"
        )

        st.divider()

        selected_components = st.multiselect(
            "Components",
            available_components,
            default=[]
        )

        stat_options = [
            "Samples",
            "Breaches",
            "% Exceeding",
            "Minimum",
            "P05",
            "P10",
            "Mean",
            "P50",
            "P90",
            "P95",
            "P99",
            "Maximum",
            "Limit"
        ]

        selected_stats = st.multiselect(
            "Statistics",
            stat_options,
            default=[]
        )

        # ============================================
        # ASTM TABLE
        # ============================================

        if (
            generate_envelope
            and not selected_components
            and not selected_stats
        ):

            envelope_components = {
                "Aluminium (total)": ["aluminium"],
                "Colour (true)": ["colour", "color"],
                "Iron (total) as Fe": ["iron"],
                "Manganese (total) as Mn": ["manganese"],
                "Turbidity": ["turbidity"],
                "Chlorophyll": ["chlorophyll"],
                "Total Organic Carbon": [
                    "toc",
                    "total organic carbon"
                ],
                "Total Dissolved Solids": [
                    "tds",
                    "total dissolved solids"
                ],
                "Temperature": ["temperature"],
                "pH": ["hydrogen ion"],
                "Alkalinity, as CaCO3": ["alkalinity"],
                "Calcium hardness, as CaCO3": [
                    "hardness"
                ],
                "Chloride": ["chloride"],
                "Sulphate SO4": [
                    "sulphate",
                    "sulfate"
                ],
                "Calcium as Ca": ["calcium"],
                "Conductivity": ["conductivity"],
                "Total Pesticides": ["pesticide"],
                "Bromide": ["bromide"]
            }

            envelope_rows = []

            for display_name, aliases in envelope_components.items():

                matched_component = None

                for component in available_components:

                    if any(
                        alias.lower() in component.lower()
                        for alias in aliases
                    ):
                        matched_component = component
                        break

                if matched_component is None:

                    envelope_rows.append({
                        "Component": display_name,
                        "Maximum": "",
                        "95%ile": "",
                        "Mean": "",
                        "5%ile": "",
                        "Minimum": "",
                        "Unit": ""
                    })

                    continue

                component_df = site_df[
                    site_df["Component_Clean"]
                    == matched_component
                ]

                results = pd.to_numeric(
                    component_df[RESULT_COL],
                    errors="coerce"
                ).dropna()

                if len(results) == 0:

                    envelope_rows.append({
                        "Component": display_name,
                        "Maximum": "",
                        "95%ile": "",
                        "Mean": "",
                        "5%ile": "",
                        "Minimum": "",
                        "Unit": ""
                    })

                    continue

                unit = ""

                if (
                    UNIT_COL
                    and UNIT_COL in component_df.columns
                ):

                    unit_values = (
                        component_df[UNIT_COL]
                        .dropna()
                        .astype(str)
                    )

                    if not unit_values.empty:
                        unit = unit_values.mode().iloc[0]

                envelope_rows.append({

                    "Component": display_name,

                    "Maximum":
                        round(results.max(), 3),

                    "95%ile":
                        round(
                            results.quantile(0.95),
                            3
                        ),

                    "Mean":
                        round(
                            results.mean(),
                            3
                        ),

                    "5%ile":
                        round(
                            results.quantile(0.05),
                            3
                        ),

                    "Minimum":
                        round(
                            results.min(),
                            3
                        ),

                    "Unit":
                        unit
                })

            envelope_df = pd.DataFrame(
                envelope_rows
            )

            st.dataframe(
                envelope_df,
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                "Download AWTM Input CSV",
                envelope_df.to_csv(index=False),
                "water_quality_envelope.csv",
                "text/csv"
            )

        # ============================================
        # CUSTOM TABLE
        # ============================================

        elif selected_components and selected_stats:

            custom_rows = []

            for component in selected_components:

                component_df = site_df[
                    site_df["Component_Clean"]
                    == component
                ]

                results = pd.to_numeric(
                    component_df[RESULT_COL],
                    errors="coerce"
                ).dropna()

                if len(results) == 0:
                    continue

                if component_df[DET_COL].mode().empty:
                    display_name = component
                else:
                    display_name = (
                        component_df[DET_COL]
                        .mode()
                        .iloc[0]
                    )

                limit_value = None
                breaches = 0

                if pcv_df is not None:

                    limits = pd.DataFrame(
                        limit_lookup.get(
                            component,
                            []
                        )
                    )

                    limits = limits[
                        limits["Regulatory Source"]
                        .isin(selected_regulations)
                    ]

                    if not limits.empty:

                        limit_value = pd.to_numeric(
                            limits["Limit Value"],
                            errors="coerce"
                        ).min()

                        if pd.notna(limit_value):

                            limit_value = get_converted_limit(
                                limit_row,
                                units
                            )

                            breaches = (
                                results > limit_value
                            ).sum()

                calculations = {

                    "Samples": len(results),

                    "Breaches": breaches,

                    "% Exceeding":
                        round(
                            100 * breaches / len(results),
                            2
                        ),

                    "Minimum":
                        round(results.min(), 3),

                    "P05":
                        round(
                            results.quantile(0.05),
                            3
                        ),

                    "P10":
                        round(
                            results.quantile(0.10),
                            3
                        ),

                    "Mean":
                        round(results.mean(), 3),

                    "P50":
                        round(
                            results.quantile(0.50),
                            3
                        ),

                    "P90":
                        round(
                            results.quantile(0.90),
                            3
                        ),

                    "P95":
                        round(
                            results.quantile(0.95),
                            3
                        ),

                    "P99":
                        round(
                            results.quantile(0.99),
                            3
                        ),

                    "Maximum":
                        round(results.max(), 3),

                    "Limit":
                        limit_value

                }

                row = {
                    "Component": display_name
                }

                for stat in selected_stats:
                    row[stat] = calculations[stat]

                custom_rows.append(row)

            custom_df = pd.DataFrame(custom_rows)

            st.dataframe(
                custom_df,
                use_container_width=True,
                hide_index=True
            )
            
            st.caption(
                f"Using data from {start_date:%d %b %Y} to {end_date:%d %b %Y}"
            )

            st.download_button(
                "Download CSV",
                custom_df.to_csv(index=False),
                "custom_summary.csv",
                "text/csv"
            )
                    
    ############################################################################
    # COMPLIANCE SUMMARY
    ############################################################################                   
                                                             
    with tab6:

        st.subheader(
            "Compliance Summary by Water Source"
        )

        generate_summary = st.button(
            "Generate Compliance Summary"
        )

        if generate_summary:

            summary_df = build_compliance_summary(
                df,
                SITE_COL,
                DET_COL,
                UNIT_COL,
                RESULT_COL,
                tuple(selected_regulations),
                limit_lookup
            )

            if not summary_df.empty:

                summary_df = summary_df.sort_values(
                    by="% Exceeding",
                    ascending=False,
                    na_position="last"
                ).reset_index(drop=True)

                st.dataframe(
                    summary_df,
                    use_container_width=True
                )

                st.info(
                    "Click on the 'Breaches' or '% Exceeding' columns to sort by the number, or proportion, of limit breaches for a given component."
                )

            else:

                st.info(
                    "No compliance results found."
                )


    # =================================================
    # COMPARISON TAB
    # =================================================

    with tab7:

        st.subheader("Source Comparison")

        dataset_names = list(loaded_datasets.keys())

        c1, c2 = st.columns(2)

        with c1:
            dataset_a_name = st.selectbox(
                "Dataset A",
                dataset_names,
                key="dataset_a"
            )

        with c2:
            dataset_b_name = st.selectbox(
                "Dataset B",
                dataset_names,
                index=min(1, len(dataset_names) - 1),
                key="dataset_b"
            )

        compare_df_a = loaded_datasets[dataset_a_name].copy()
        compare_df_b = loaded_datasets[dataset_b_name].copy()

        DATE_A_DEFAULT = find_column(
            compare_df_a,
            COLUMN_ALIASES["DATE"]
        )

        SITE_A_DEFAULT = find_column(
            compare_df_a,
            COLUMN_ALIASES["SITE"]
        )

        COMPONENT_A_DEFAULT = find_column(
            compare_df_a,
            COLUMN_ALIASES["PARAMETER"]
        )

        RESULT_A_DEFAULT = find_column(
            compare_df_a,
            COLUMN_ALIASES["RESULT"]
        )
        
        UNIT_A_DEFAULT = find_column(
            compare_df_a,
            COLUMN_ALIASES["UNIT"]
        )
        
        OPERATOR_A_DEFAULT = find_column(
            compare_df_a,
            COLUMN_ALIASES["OPERATOR"]
        )
                
        DATE_B_DEFAULT = find_column(
            compare_df_b,
            COLUMN_ALIASES["DATE"]
        )

        SITE_B_DEFAULT = find_column(
            compare_df_b,
            COLUMN_ALIASES["SITE"]
        )

        COMPONENT_B_DEFAULT = find_column(
            compare_df_b,
            COLUMN_ALIASES["PARAMETER"]
        )

        RESULT_B_DEFAULT = find_column(
            compare_df_b,
            COLUMN_ALIASES["RESULT"]
        )
        
        UNIT_B_DEFAULT = find_column(
            compare_df_b,
            COLUMN_ALIASES["UNIT"]
        ) 

        OPERATOR_B_DEFAULT = find_column(
            compare_df_b,
            COLUMN_ALIASES["OPERATOR"]
        )        

        # ---------------------------------------
        # COLUMN MAPPING
        # ---------------------------------------

        c1, c2 = st.columns(2)

        with c1:

            with st.expander("Source A Column Mapping"):
                
                LAYOUT_A = st.radio(
                    "Layout A",
                    [
                        "Long Format",
                        "Wide Format"
                    ],
                    key="LAYOUT_A"
                )

                FIRST_COMPONENT_A = None

                if LAYOUT_A == "Wide Format":

                    FIRST_COMPONENT_A = st.selectbox(
                        "First Component Column A",
                        compare_df_a.columns,
                        key="FIRST_COMPONENT_A"
                    )
                            

                DATE_A = st.selectbox(
                    "Date Column A",
                    compare_df_a.columns,
                    index=(
                        compare_df_a.columns.get_loc(DATE_A_DEFAULT)
                        if DATE_A_DEFAULT
                        else 0
                    ),
                    key="DATE_A"
                )

                site_options_a = (
                    ["No Site Column"]
                    + list(compare_df_a.columns)
                )

                SITE_A = st.selectbox(
                    "Site Column A",
                    site_options_a,
                    index=(
                        site_options_a.index(SITE_A_DEFAULT)
                        if SITE_A_DEFAULT in site_options_a
                        else 0
                    ),
                    key="SITE_A"
                )

                if LAYOUT_A == "Long Format":

                    COMPONENT_A = st.selectbox(
                        "Component Column A",
                        compare_df_a.columns,
                        index=(
                            compare_df_a.columns.get_loc(COMPONENT_A_DEFAULT)
                            if COMPONENT_A_DEFAULT
                            else 0
                        ),
                        key="COMPONENT_A"
                    )

                    RESULT_A = st.selectbox(
                        "Result Column A",
                        compare_df_a.columns,
                        index=(
                            compare_df_a.columns.get_loc(RESULT_A_DEFAULT)
                            if RESULT_A_DEFAULT
                            else 0
                        ),
                        key="RESULT_A"
                    )

                else:

                    st.info(
                        "Wide format selected. Component and Result columns will be created automatically."
                    )

                
                if LAYOUT_A == "Long Format":

                    UNIT_A_OPTIONS = ["No Unit Column"] + list(compare_df_a.columns)

                    UNIT_A = st.selectbox(
                        "Unit Column A",
                        UNIT_A_OPTIONS,
                        index=(
                            UNIT_A_OPTIONS.index(UNIT_A_DEFAULT)
                            if UNIT_A_DEFAULT in UNIT_A_OPTIONS
                            else 0
                        ),
                        key="UNIT_A"
                    )
                    
                else:

                    UNIT_A = "No Unit Column"
 
                if LAYOUT_A == "Long Format":

                    OPERATOR_A_OPTIONS = ["No Operator Column"] + list(compare_df_a.columns)

                    OPERATOR_A = st.selectbox(
                        "Operator Column A",
                        OPERATOR_A_OPTIONS,
                        index=(
                            OPERATOR_A_OPTIONS.index(OPERATOR_A_DEFAULT)
                            if OPERATOR_A_DEFAULT in OPERATOR_A_OPTIONS
                            else 0
                        ),
                        key="OPERATOR_A"
                    )

                else:
                    OPERATOR_A = "No Operator Column"                    

 
        with c2:

            with st.expander("Source B Column Mapping"):
                
                LAYOUT_B = st.radio(
                    "Layout B",
                    [
                        "Long Format",
                        "Wide Format"
                    ],
                    key="LAYOUT_B"
                )

                FIRST_COMPONENT_B = None

                if LAYOUT_B == "Wide Format":

                    FIRST_COMPONENT_B = st.selectbox(
                        "First Component Column B",
                        compare_df_b.columns,
                        key="FIRST_COMPONENT_B"
                    )

                DATE_B = st.selectbox(
                    "Date Column B",
                    compare_df_b.columns,
                    index=(
                        compare_df_b.columns.get_loc(DATE_B_DEFAULT)
                        if DATE_B_DEFAULT
                        else 0
                    ),
                    key="DATE_B"
                )

                site_options_b = (
                    ["No Site Column"]
                    + list(compare_df_b.columns)
                )

                SITE_B = st.selectbox(
                    "Site Column B",
                    site_options_b,
                    index=(
                        site_options_b.index(SITE_B_DEFAULT)
                        if SITE_B_DEFAULT in site_options_b
                        else 0
                    ),
                    key="SITE_B"
                )

                if LAYOUT_B == "Long Format":

                    COMPONENT_B = st.selectbox(
                        "Component Column B",
                        compare_df_b.columns,
                        index=(
                            compare_df_b.columns.get_loc(COMPONENT_B_DEFAULT)
                            if COMPONENT_B_DEFAULT
                            else 0
                        ),
                        key="COMPONENT_B"
                    )

                    RESULT_B = st.selectbox(
                        "Result Column B",
                        compare_df_b.columns,
                        index=(
                            compare_df_b.columns.get_loc(RESULT_B_DEFAULT)
                            if RESULT_B_DEFAULT
                            else 0
                        ),
                        key="RESULT_B"
                    )

                else:

                    st.info(
                        "Wide format selected. Component and Result columns will be created automatically."
                    )
                
                if LAYOUT_B == "Long Format":

                    UNIT_B_OPTIONS = ["No Unit Column"] + list(compare_df_b.columns)

                    UNIT_B = st.selectbox(
                        "Unit Column B",
                        UNIT_B_OPTIONS,
                        index=(
                            UNIT_B_OPTIONS.index(UNIT_B_DEFAULT)
                            if UNIT_B_DEFAULT in UNIT_B_OPTIONS
                            else 0
                        ),
                        key="UNIT_B"
                    )

                else:

                    UNIT_B = "No Unit Column"   

                if LAYOUT_B == "Long Format":

                    OPERATOR_B_OPTIONS = ["No Operator Column"] + list(compare_df_b.columns)

                    OPERATOR_B = st.selectbox(
                        "Operator Column B",
                        OPERATOR_B_OPTIONS,
                        index=(
                            OPERATOR_B_OPTIONS.index(OPERATOR_B_DEFAULT)
                            if OPERATOR_B_DEFAULT in OPERATOR_B_OPTIONS
                            else 0
                        ),
                        key="OPERATOR_B"
                    )

                else:
                    OPERATOR_B = "No Operator Column"
                                    
        # =====================================
        # SOURCE A WIDE FORMAT
        # =====================================

        if LAYOUT_A == "Wide Format":

            start_idx = compare_df_a.columns.get_loc(
                FIRST_COMPONENT_A
            )

            component_cols = list(
                compare_df_a.columns[start_idx:]
            )

            id_cols = list(
                compare_df_a.columns[:start_idx]
            )

            compare_df_a = compare_df_a.melt(
                id_vars=id_cols,
                value_vars=component_cols,
                var_name="Component",
                value_name="Result"
            )

            COMPONENT_A = "Component"
            RESULT_A = "Result"
            UNIT_A = "No Unit Column"
            OPERATOR_A = "No Operator Column"

        # =====================================
        # SOURCE B WIDE FORMAT
        # =====================================

        if LAYOUT_B == "Wide Format":

            start_idx = compare_df_b.columns.get_loc(
                FIRST_COMPONENT_B
            )

            component_cols = list(
                compare_df_b.columns[start_idx:]
            )

            id_cols = list(
                compare_df_b.columns[:start_idx]
            )

            compare_df_b = compare_df_b.melt(
                id_vars=id_cols,
                value_vars=component_cols,
                var_name="Component",
                value_name="Result"
            )

            COMPONENT_B = "Component"
            RESULT_B = "Result"  
            UNIT_B = "No Unit Column"
            OPERATOR_B = "No Operator Column"

        # =====================================
        # AUTO UNIT EXTRACTION
        # =====================================

        # Source A
        if (
            UNIT_A == "No Unit Column"
            or UNIT_A not in compare_df_a.columns
        ):
            compare_df_a = add_auto_unit_column(
                compare_df_a,
                COMPONENT_A
            )
            UNIT_A = "Auto_Unit"

        # Source B
        if (
            UNIT_B == "No Unit Column"
            or UNIT_B not in compare_df_b.columns
        ):
            compare_df_b = add_auto_unit_column(
                compare_df_b,
                COMPONENT_B
            )
            UNIT_B = "Auto_Unit"

        # ---------------------------------------
        # CLEAN DATA
        # ---------------------------------------

        compare_df_a[DATE_A] = pd.to_datetime(
            compare_df_a[DATE_A],
            errors="coerce"
        )

        compare_df_b[DATE_B] = pd.to_datetime(
            compare_df_b[DATE_B],
            errors="coerce"
        )

        # =====================================
        # SOURCE A CLEANING
        # =====================================

        compare_df_a["Original_Result"] = (
            compare_df_a[RESULT_A].astype(str)
        )

        NON_DETECT_FLAGS = {
            "<",
            "<=",
            "ND",
            "LOD",
            "BDL",
            "U",
            "UJ",
            "LT"
        }

        if (
            OPERATOR_A != "No Operator Column"
            and OPERATOR_A in compare_df_a.columns
        ):
            compare_df_a["Non_Detect"] = (
                compare_df_a[OPERATOR_A]
                .astype(str)
                .str.upper()
                .str.strip()
                .isin(NON_DETECT_FLAGS)
            )
        else:
            compare_df_a["Non_Detect"] = (
                compare_df_a["Original_Result"]
                .str.strip()
                .str.startswith("<")
            )

        compare_df_a[RESULT_A] = (
            compare_df_a["Original_Result"]
            .str.replace("<", "", regex=False)
        )

        compare_df_a[RESULT_A] = pd.to_numeric(
            compare_df_a[RESULT_A],
            errors="coerce"
        )

        compare_df_a.loc[
            compare_df_a["Non_Detect"],
            RESULT_A
        ] = (
            compare_df_a.loc[
                compare_df_a["Non_Detect"],
                RESULT_A
            ] / 2
        )

        compare_df_a = compare_df_a[
            compare_df_a[RESULT_A] > 0
        ]
        
        compare_df_a["Component_Clean"] = (
            compare_df_a[COMPONENT_A]
            .astype(str)
            .apply(clean_component_name)
        )        

        # =====================================
        # SOURCE B CLEANING
        # =====================================

        compare_df_b["Original_Result"] = (
            compare_df_b[RESULT_B].astype(str)
        )

        if (
            OPERATOR_B != "No Operator Column"
            and OPERATOR_B in compare_df_b.columns
        ):
            compare_df_b["Non_Detect"] = (
                compare_df_b[OPERATOR_B]
                .astype(str)
                .str.upper()
                .str.strip()
                .isin(NON_DETECT_FLAGS)
            )
        else:
            compare_df_b["Non_Detect"] = (
                compare_df_b["Original_Result"]
                .str.strip()
                .str.startswith("<")
            )

        compare_df_b[RESULT_B] = (
            compare_df_b["Original_Result"]
            .str.replace("<", "", regex=False)
        )

        compare_df_b[RESULT_B] = pd.to_numeric(
            compare_df_b[RESULT_B],
            errors="coerce"
        )

        compare_df_b.loc[
            compare_df_b["Non_Detect"],
            RESULT_B
        ] = (
            compare_df_b.loc[
                compare_df_b["Non_Detect"],
                RESULT_B
            ] / 2
        )

        compare_df_b = compare_df_b[
            compare_df_b[RESULT_B] > 0
        ]
        
        compare_df_b["Component_Clean"] = (
            compare_df_b[COMPONENT_B]
            .astype(str)
            .apply(clean_component_name)
        )

        # ---------------------------------------
        # COMPONENTS
        # ---------------------------------------

        c1, c2 = st.columns(2)

        with c1:
            
            compare_df_a = compare_df_a.dropna(
                subset=[COMPONENT_A]
            )

            component_lookup_a = (
                compare_df_a.groupby("Component_Clean")[COMPONENT_A]
                .agg(
                    lambda x: (
                        x.mode().iloc[0]
                        if not x.mode().empty
                        else x.dropna().iloc[0]
                    )
                )
                .to_dict()
            )

            display_names_a = sorted(component_lookup_a.values())

            component_display_a = st.selectbox(
                "Component A",
                display_names_a,
                key="component_a_compare"
            )

            component_a = clean_component_name(
                component_display_a
            )

        with c2:
            
            compare_df_b = compare_df_b.dropna(
                subset=[COMPONENT_B]
            )

            component_lookup_b = (
                compare_df_b.groupby("Component_Clean")[COMPONENT_B]
                .agg(
                    lambda x: (
                        x.mode().iloc[0]
                        if not x.mode().empty
                        else x.dropna().iloc[0]
                    )
                )
                .to_dict()
            )

            display_names_b = sorted(
                component_lookup_b.values()
            )

            component_display_b = st.selectbox(
                "Component B",
                display_names_b,
                key="component_b_compare"
            )

            component_b = clean_component_name(
                component_display_b
            )

        unit_col_a, unit_col_b = st.columns(2)
        with unit_col_a:
            comparison_manual_unit_a = st.text_input(
                "Unit for Component A",
                value=st.session_state.manual_component_units.get(component_a, ""),
                key=f"comparison_unit_a_{component_a}"
            ).strip()
        with unit_col_b:
            comparison_manual_unit_b = st.text_input(
                "Unit for Component B",
                value=st.session_state.manual_component_units.get(component_b, ""),
                key=f"comparison_unit_b_{component_b}"
            ).strip()

        if comparison_manual_unit_a:
            comparison_manual_unit_a = clean_unit_value(comparison_manual_unit_a)
            st.session_state.manual_component_units[component_a] = comparison_manual_unit_a
        if comparison_manual_unit_b:
            comparison_manual_unit_b = clean_unit_value(comparison_manual_unit_b)
            st.session_state.manual_component_units[component_b] = comparison_manual_unit_b

        compare_a = compare_df_a[
            compare_df_a["Component_Clean"] == component_a
        ].copy()

        compare_b = compare_df_b[
            compare_df_b["Component_Clean"]
            == component_b
        ].copy()
        
        units_a = ""
        if (
            UNIT_A != "No Unit Column"
            and UNIT_A in compare_a.columns
        ):
            unit_values = (
                compare_a[UNIT_A]
                .fillna("")
                .astype(str)
                .apply(clean_unit_value)
            )
            unit_values = unit_values[unit_values.str.strip() != ""]

            if not unit_values.empty:
                units_a = unit_values.mode().iloc[0]

        units_b = ""
        if (
            UNIT_B != "No Unit Column"
            and UNIT_B in compare_b.columns
        ):
            unit_values = (
                compare_b[UNIT_B]
                .fillna("")
                .astype(str)
                .apply(clean_unit_value)
            )
            unit_values = unit_values[unit_values.str.strip() != ""]

            if not unit_values.empty:
                units_b = unit_values.mode().iloc[0]

        if comparison_manual_unit_a:
            units_a = comparison_manual_unit_a
        if comparison_manual_unit_b:
            units_b = comparison_manual_unit_b

        # ---------------------------------------
        # SITE FILTERS
        # ---------------------------------------

        site_col1, site_col2 = st.columns(2)

        # =====================================
        # SOURCE A SITE
        # =====================================

        if SITE_A != "No Site Column":

            available_a = (
                ["All Sites"]
                + sorted(
                    compare_a[SITE_A]
                    .dropna()
                    .unique()
                )
            )

            with site_col1:

                selected_a = st.selectbox(
                    "Source A Site",
                    available_a,
                    key="site_a_compare"
                )

            if selected_a != "All Sites":

                compare_a = compare_a[
                    compare_a[SITE_A] == selected_a
                ]

            title_a = selected_a

        else:

            title_a = dataset_a_name

        # =====================================
        # SOURCE B SITE
        # =====================================

        if SITE_B != "No Site Column":

            available_b = (
                ["All Sites"]
                + sorted(
                    compare_b[SITE_B]
                    .dropna()
                    .unique()
                )
            )

            with site_col2:

                selected_b = st.selectbox(
                    "Source B Site",
                    available_b,
                    key="site_b_compare"
                )

            if selected_b != "All Sites":

                compare_b = compare_b[
                    compare_b[SITE_B] == selected_b
                ]

            title_b = selected_b

        else:

            title_b = dataset_b_name
            
        # ---------------------------------------
        # APPLY SIDEBAR DATE RANGE
        # ---------------------------------------

        comparison_start = pd.to_datetime(start_date)

        # Use the start of the following day so results containing a time
        # on the selected end date are not accidentally excluded.
        comparison_end_exclusive = (
            pd.to_datetime(end_date)
            + pd.Timedelta(days=1)
        )

        compare_a = compare_a[
            (compare_a[DATE_A] >= comparison_start)
            &
            (compare_a[DATE_A] < comparison_end_exclusive)
        ].copy()

        compare_b = compare_b[
            (compare_b[DATE_B] >= comparison_start)
            &
            (compare_b[DATE_B] < comparison_end_exclusive)
        ].copy()

        st.caption(
            f"Comparison period: "
            f"{comparison_start:%d %b %Y} to "
            f"{pd.to_datetime(end_date):%d %b %Y}"
        )

        if compare_a.empty:
            st.warning(
                f"No Source A results are available between "
                f"{comparison_start:%d %b %Y} and "
                f"{pd.to_datetime(end_date):%d %b %Y}."
            )

        if compare_b.empty:
            st.warning(
                f"No Source B results are available between "
                f"{comparison_start:%d %b %Y} and "
                f"{pd.to_datetime(end_date):%d %b %Y}."
            )

        if compare_a.empty or compare_b.empty:
            st.stop()            

        # ---------------------------------------
        # FIGURE
        # ---------------------------------------
        
        # Split detects and non-detects

        compare_a_detect = compare_a[
            ~compare_a["Non_Detect"]
        ].copy()

        compare_a_nondetect = compare_a[
            compare_a["Non_Detect"]
        ].copy()

        compare_b_detect = compare_b[
            ~compare_b["Non_Detect"]
        ].copy()

        compare_b_nondetect = compare_b[
            compare_b["Non_Detect"]
        ].copy()        

        mean_a = compare_a[RESULT_A].mean()
        p50_a = compare_a[RESULT_A].quantile(0.50)
        p95_a = compare_a[RESULT_A].quantile(0.95)

        mean_b = compare_b[RESULT_B].mean()
        p50_b = compare_b[RESULT_B].quantile(0.50)
        p95_b = compare_b[RESULT_B].quantile(0.95)
        
        limits_a = pd.DataFrame()
        limits_b = pd.DataFrame()

        if pcv_df is not None:

            # Automatic matches
            limits_a = pd.DataFrame(
                limit_lookup.get(
                    component_a,
                    []
                )
            )

            limits_b = pd.DataFrame(
                limit_lookup.get(
                    component_b,
                    []
                )
            )

            # Manual PCV additions
            if manual_pcvs:

                manual_limits = pcv_df[
                    pcv_df["Parameter"].isin(manual_pcvs)
                ]

                limits_a = pd.concat(
                    [
                        limits_a,
                        manual_limits
                    ],
                    ignore_index=True
                ).drop_duplicates()

                limits_b = pd.concat(
                    [
                        limits_b,
                        manual_limits
                    ],
                    ignore_index=True
                ).drop_duplicates()

            # Allow both normal regulation selections
            # and manually-added regulation groups
            allowed_regulations = set(
                selected_regulations
            )

            if manual_regulations:
                allowed_regulations.update(
                    manual_regulations
                )

            if (
                not limits_a.empty
                and "Regulatory Source" in limits_a.columns
            ):
                limits_a = limits_a[
                    limits_a["Regulatory Source"]
                    .isin(allowed_regulations)
                ]

            if (
                not limits_b.empty
                and "Regulatory Source" in limits_b.columns
            ):
                limits_b = limits_b[
                    limits_b["Regulatory Source"]
                    .isin(allowed_regulations)
                ]

        fig_compare = make_subplots(
            rows=1,
            cols=2,
            shared_yaxes=True,
            subplot_titles=(title_a, title_b),
            horizontal_spacing=0.03
        )

        # SOURCE A DETECTS

        fig_compare.add_trace(
            go.Scatter(
                x=compare_a_detect[DATE_A],
                y=compare_a_detect[RESULT_A],
                mode="markers",
                marker=dict(
                    color="#005A9C",
                    size=8
                ),
                name=f"{title_a} Detected"
            ),
            row=1,
            col=1
        )

        # SOURCE A NON DETECTS

        fig_compare.add_trace(
            go.Scatter(
                x=compare_a_nondetect[DATE_A],
                y=compare_a_nondetect[RESULT_A],
                mode="markers",
                marker=dict(
                    color="red",
                    size=6,
                    symbol="diamond",
                    line=dict(
                        color="darkred",
                        width=0.5
                    )
                ),
                name=f"{title_a} Below LOD"
            ),
            row=1,
            col=1
        )

        for value, colour, dash, label in [
            (mean_a, "#000000", "dot", f"{title_a} Mean ({mean_a:.2f} {units_a})"),
            (p50_a, "#4B0082", "dash", f"{title_a} P50 ({p50_a:.2f} {units_a})"),
            (p95_a, "#FF00FF", "dash", f"{title_a} P95 ({p95_a:.2f} {units_a})")
        ]:

            fig_compare.add_trace(
                go.Scatter(
                    x=[
                        compare_a[DATE_A].min(),
                        compare_a[DATE_A].max()
                    ],
                    y=[value, value],
                    mode="lines",
                    name=label,
                    line=dict(
                        color=colour,
                        width=2,
                        dash=dash
                    )
                ),
                row=1,
                col=1
            )
            
        colour_palette = [
            "#4DAF4A",
            "#FF7F00",
            "#E41A1C",
            "#377EB8",
            "#984EA3"
        ]

        regulation_colour_map = {}

        for _, row in limits_a.iterrows():

            limit_value = get_converted_limit(
                row,
                units_a
            )

            if pd.isna(limit_value):
                continue

            regulation = row.get(
                "Regulatory Source",
                ""
            )

            poc = row.get(
                "Point of Compliance",
                ""
            )

            colour_key = (
                str(regulation).strip(),
                str(poc).strip()
            )

            if colour_key not in regulation_colour_map:
                regulation_colour_map[colour_key] = colour_palette[
                    len(regulation_colour_map)
                    % len(colour_palette)
                ]

            line_colour = regulation_colour_map[colour_key]

            fig_compare.add_trace(
                go.Scatter(
                    x=[
                        compare_a[DATE_A].min(),
                        compare_a[DATE_A].max()
                    ],
                    y=[limit_value, limit_value],
                    mode="lines",
                    name=f"{regulation} - {poc} ({limit_value} {units_a})",
                    line=dict(
                        color=line_colour,
                        width=3,
                        dash="dash"
                    )
                ),
                row=1,
                col=1
            )

        # SOURCE B DETECTS

        fig_compare.add_trace(
            go.Scatter(
                x=compare_b_detect[DATE_B],
                y=compare_b_detect[RESULT_B],
                mode="markers",
                marker=dict(
                    color="#FF7F00",
                    size=8
                ),
                name=f"{title_b} Detected",
                showlegend=True
            ),
            row=1,
            col=2
        )

        # SOURCE B NON DETECTS

        fig_compare.add_trace(
            go.Scatter(
                x=compare_b_nondetect[DATE_B],
                y=compare_b_nondetect[RESULT_B],
                mode="markers",
                marker=dict(
                    color="red",
                    size=6,
                    symbol="diamond",
                    line=dict(
                        color="darkred",
                        width=0.5
                    )
                ),
                name=f"{title_b} Below LOD",
                showlegend=True
            ),
            row=1,
            col=2
        )

        for value, colour, dash, label in [
            (mean_b, "#000000", "dot", f"{title_b} Mean ({mean_b:.2f} {units_b})"),
            (p50_b, "#4B0082", "dash", f"{title_b} P50 ({p50_b:.2f} {units_b})"),
            (p95_b, "#FF00FF", "dash", f"{title_b} P95 ({p95_b:.2f} {units_b})")
        ]:

            fig_compare.add_trace(
                go.Scatter(
                    x=[
                        compare_b[DATE_B].min(),
                        compare_b[DATE_B].max()
                    ],
                    y=[value, value],
                    mode="lines",
                    name=label,
                    showlegend=True,
                    line=dict(
                        color=colour,
                        width=2,
                        dash=dash
                    )
                ),
                row=1,
                col=2
            )

        for _, row in limits_b.iterrows():

            limit_value = get_converted_limit(
                row,
                units_b
            )

            if pd.isna(limit_value):
                continue

            regulation = row.get(
                "Regulatory Source",
                ""
            )

            poc = row.get(
                "Point of Compliance",
                ""
            )

            colour_key = (
                str(regulation).strip(),
                str(poc).strip()
            )

            if colour_key not in regulation_colour_map:
                regulation_colour_map[colour_key] = colour_palette[
                    len(regulation_colour_map)
                    % len(colour_palette)
                ]

            line_colour = regulation_colour_map[colour_key]

            fig_compare.add_trace(
                go.Scatter(
                    x=[
                        compare_b[DATE_B].min(),
                        compare_b[DATE_B].max()
                    ],
                    y=[limit_value, limit_value],
                    mode="lines",
                    name=f"{regulation} - {poc} ({limit_value} {units_b})",
                    line=dict(
                        color=line_colour,
                        width=3,
                        dash="dash"
                    ),
                    showlegend=False
                ),
                row=1,
                col=2
            )

        fig_compare.update_layout(
            template="simple_white",
            height=700,
            showlegend=True,
            legend=dict(
                title="Reported Result",
                x=1.02,
                y=1
            ),
            margin=dict(
                l=20,
                r=20,
                t=80,
                b=40
            ),
            title=dict(
                text="Source Comparison",
                x=0.01
            )
        )

        fig_compare.update_xaxes(
            title_text="Sample Date",
            range=[
                comparison_start,
                pd.to_datetime(end_date)
            ],
            showgrid=True,
            showline=True,
            linewidth=2,
            linecolor="black",
            gridcolor="#D9D9D9"
        )

        fig_compare.update_yaxes(
            title_text=(
                f"{component_display_a} ({units_a})"
                if units_a else component_display_a
            ),
            showgrid=True,
            showline=True,
            linewidth=2,
            linecolor="black",
            gridcolor="#D9D9D9",
            row=1,
            col=1
        )
        fig_compare.update_yaxes(
            title_text=(
                f"{component_display_b} ({units_b})"
                if units_b else component_display_b
            ),
            showgrid=True,
            showline=True,
            linewidth=2,
            linecolor="black",
            gridcolor="#D9D9D9",
            row=1,
            col=2
        )

        st.plotly_chart(
            fig_compare,
            use_container_width=True
        )

        st.subheader("Comparison Statistics")

        comparison_stats = pd.DataFrame({
            "Statistic": [
                "Units",
                "Samples",
                "Minimum",
                "Mean",
                "P50",
                "P95",
                "Maximum"
            ],
            title_a: [
                units_a,
                len(compare_a),
                round(compare_a[RESULT_A].min(), 3),
                round(mean_a, 3),
                round(p50_a, 3),
                round(p95_a, 3),
                round(compare_a[RESULT_A].max(), 3)
            ],
            title_b: [
                units_b,
                len(compare_b),
                round(compare_b[RESULT_B].min(), 3),
                round(mean_b, 3),
                round(p50_b, 3),
                round(p95_b, 3),
                round(compare_b[RESULT_B].max(), 3)
            ]
        })

        st.dataframe(
            comparison_stats,
            use_container_width=True
        )

    # =======================================
    # EXPORT
    # =======================================

    with tab8:

        if st.button(
            "Generate Export Files",
            key="generate_exports"
        ):

            # --------------------------------------------------
            # ALWAYS BUILD COMPLIANCE SUMMARY FOR PDF EXPORT
            # --------------------------------------------------

            summary_df = build_compliance_summary(
                df,
                SITE_COL,
                DET_COL,
                UNIT_COL,
                RESULT_COL,
                tuple(selected_regulations),
                limit_lookup
            )

            if not summary_df.empty:
                summary_df = summary_df.sort_values(
                    by="% Exceeding",
                    ascending=False,
                    na_position="last"
                ).reset_index(drop=True)

            # --------------------------------------------------
            # EXCEL EXPORT
            # --------------------------------------------------

            excel_data = build_export(
                stats_df,
                filtered,
                selected_site,
                selected_component,
                start_date,
                end_date,
                uploaded_file.name,
                component_limits,
                sample_unit
            )

            # --------------------------------------------------
            # FIGURE EXPORT
            # --------------------------------------------------

            import tempfile

            png_bytes = png_export_fig.to_image(
                format="png",
                width=2400,
                height=1350,
                scale=2
            )

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".png"
            ) as tmp:

                tmp.write(png_bytes)
                chart_path = tmp.name

            # --------------------------------------------------
            # PDF EXPORT
            # --------------------------------------------------

            pdf_data = build_pdf_report(
                stats_df,
                summary_df,
                selected_site,
                selected_component,
                start_date,
                end_date,
                uploaded_file.name,
                sample_unit,
                chart_path
            )

            # --------------------------------------------------
            # FILE NAMES
            # --------------------------------------------------

            export_site = uploaded_file.name.rsplit(".", 1)[0]

            start_txt = pd.to_datetime(
                start_date
            ).strftime("%Y%m%d")

            end_txt = pd.to_datetime(
                end_date
            ).strftime("%Y%m%d")

            safe_component = (
                selected_component
                .replace("/", "-")
                .replace("\\", "-")
            )

            excel_filename = (
                f"{export_site}_"
                f"{safe_component}_"
                f"{start_txt}_to_{end_txt}.xlsx"
            )

            pdf_filename = (
                f"{export_site}_"
                f"{safe_component}_"
                f"{start_txt}_to_{end_txt}.pdf"
            )

            # --------------------------------------------------
            # DOWNLOAD BUTTONS
            # --------------------------------------------------

            st.download_button(
                "Download Excel Report",
                data=excel_data,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.download_button(
                "Download PDF Report",
                data=pdf_data,
                file_name=pdf_filename,
                mime="application/pdf"
            )
            
        st.info("Ensure that the site compliance summary has been generated before generating any reports!")

    # =================================================
    # PFAS / PESTICIDE TOTALS
    # =================================================

    with tab9:

        st.subheader("PFAS & Pesticide Totals")
        
        use_company_totals = st.checkbox(
            "Apply company-specific PFAS and pesticide total limits",
            value=False
        )

        ACTIVE_PFAS_LIMIT = PFAS_LIMIT
        ACTIVE_PFAS_UNIT = PFAS_LIMIT_UNIT

        ACTIVE_PESTICIDE_LIMIT = PESTICIDE_LIMIT
        ACTIVE_PESTICIDE_UNIT = PESTICIDE_LIMIT_UNIT

        if use_company_totals:

            available_companies = sorted(
                set(PFAS_COMPANY_LIMITS.keys())
                | set(PESTICIDE_COMPANY_LIMITS.keys())
            )

            selected_company = st.selectbox(
                "Company",
                available_companies
            )

            if selected_company in PFAS_COMPANY_LIMITS:

                ACTIVE_PFAS_LIMIT = (
                    PFAS_COMPANY_LIMITS[selected_company]["limit"]
                )

                ACTIVE_PFAS_UNIT = (
                    PFAS_COMPANY_LIMITS[selected_company]["unit"]
                )

            if selected_company in PESTICIDE_COMPANY_LIMITS:

                ACTIVE_PESTICIDE_LIMIT = (
                    PESTICIDE_COMPANY_LIMITS[selected_company]["limit"]
                )

                ACTIVE_PESTICIDE_UNIT = (
                    PESTICIDE_COMPANY_LIMITS[selected_company]["unit"]
                )

            st.info(
                f"Using {selected_company} totals: "
                f"PFAS = {ACTIVE_PFAS_LIMIT} {ACTIVE_PFAS_UNIT}, "
                f"Pesticides = {ACTIVE_PESTICIDE_LIMIT} {ACTIVE_PESTICIDE_UNIT}"
            )

        site_df = df[
            (df[SITE_COL] == selected_site)
            &
            (df[DATE_COL] >= pd.to_datetime(start_date))
            &
            (df[DATE_COL] < filter_end_exclusive)
        ].copy()

        # -----------------------------------------
        # PFAS / PESTICIDE MATCHING
        # -----------------------------------------

        (
            pfas_results,
            pesticide_results,
            pfas_totals,
            pesticide_totals
        ) = build_pfas_pesticide_summary(
            site_df,
            DATE_COL,
            RESULT_COL,
            tuple(PFAS_COMPONENTS),
            tuple(PESTICIDE_COMPONENTS)
        )
        
        matched_pfas = sorted(
            pfas_results[DET_COL]
            .dropna()
            .unique()
        )

        matched_pesticides = sorted(
            pesticide_results[DET_COL]
            .dropna()
            .unique()
        )

        totals_df = pd.merge(
            pfas_totals,
            pesticide_totals,
            on=DATE_COL,
            how="outer"
        ).sort_values(DATE_COL)

        total_units = ""

        if (
            UNIT_COL
            and UNIT_COL in site_df.columns
        ):

            unit_values = (
                site_df[UNIT_COL]
                .dropna()
                .astype(str)
            )

            if not unit_values.empty:
                total_units = unit_values.mode().iloc[0]

        # -----------------------------------------
        # FIGURE
        # -----------------------------------------

        fig_totals = go.Figure()

        if not pfas_totals.empty:

            fig_totals.add_trace(
                go.Scatter(
                    x=pfas_totals[DATE_COL],
                    y=pfas_totals["PFAS Total"],
                    mode="lines+markers",
                    name="PFAS Total",
                    line=dict(
                        color="purple",
                        width=3
                    )
                )
            )

            pfas_mean = pfas_totals["PFAS Total"].mean()

            fig_totals.add_trace(
                go.Scatter(
                    x=[
                        pfas_totals[DATE_COL].min(),
                        pfas_totals[DATE_COL].max()
                    ],
                    y=[
                        pfas_mean,
                        pfas_mean
                    ],
                    mode="lines",
                    name=(
                        f"PFAS Mean "
                        f"({pfas_mean:.3f} {total_units})"
                    ),
                    line=dict(
                        color="purple",
                        width=2,
                        dash="dot"
                    )
                )
            )

        if not pesticide_totals.empty:

            fig_totals.add_trace(
                go.Scatter(
                    x=pesticide_totals[DATE_COL],
                    y=pesticide_totals["Pesticides Total"],
                    mode="lines+markers",
                    name="Pesticides Total",
                    line=dict(
                        color="green",
                        width=3
                    )
                )
            )

            pesticide_mean = (
                pesticide_totals[
                    "Pesticides Total"
                ].mean()
            )

            fig_totals.add_trace(
                go.Scatter(
                    x=[
                        pesticide_totals[DATE_COL].min(),
                        pesticide_totals[DATE_COL].max()
                    ],
                    y=[
                        pesticide_mean,
                        pesticide_mean
                    ],
                    mode="lines",
                    name=(
                        f"Pesticides Mean "
                        f"({pesticide_mean:.3f} {total_units})"
                    ),
                    line=dict(
                        color="darkgreen",
                        width=2,
                        dash="dot"
                    )
                )
            )

        if pd.notna(ACTIVE_PFAS_LIMIT):

            limit_name = (
                f"{selected_company} PFAS Limit"
                if use_company_totals
                else "PFAS Limit"
            )

            fig_totals.add_trace(
                go.Scatter(
                    x=[
                        totals_df[DATE_COL].min(),
                        totals_df[DATE_COL].max()
                    ],
                    y=[
                        ACTIVE_PFAS_LIMIT,
                        ACTIVE_PFAS_LIMIT
                    ],
                    mode="lines",
                    name=(
                        f"{limit_name} "
                        f"({ACTIVE_PFAS_LIMIT:.3f} {ACTIVE_PFAS_UNIT})"
                    ),
                    line=dict(
                        color="purple",
                        width=3,
                        dash="dash"
                    )
                )
            )

        if pd.notna(ACTIVE_PESTICIDE_LIMIT):

            limit_name = (
                f"{selected_company} Pesticide Limit"
                if use_company_totals
                else "Pesticide Limit"
            )

            fig_totals.add_trace(
                go.Scatter(
                    x=[
                        totals_df[DATE_COL].min(),
                        totals_df[DATE_COL].max()
                    ],
                    y=[
                        ACTIVE_PESTICIDE_LIMIT,
                        ACTIVE_PESTICIDE_LIMIT
                    ],
                    mode="lines",
                    name=(
                        f"{limit_name} "
                        f"({ACTIVE_PESTICIDE_LIMIT:.3f} {ACTIVE_PESTICIDE_UNIT})"
                    ),
                    line=dict(
                        color="darkgreen",
                        width=3,
                        dash="dash"
                    )
                )
            )

        fig_totals.update_layout(
            template="simple_white",
            height=700,
            title=(
                f"{selected_site}"
                "<br>PFAS & Pesticide Totals"
                + (
                    f"<br>Using {selected_company} Standards"
                    if use_company_totals
                    else ""
                )
            ),
            xaxis_title="Sample Date",
            yaxis_title=f"Total Concentration ({ACTIVE_PFAS_UNIT})",
            legend_title="Series"
        )

        st.plotly_chart(
            fig_totals,
            use_container_width=True
        )

        # -----------------------------------------
        # SUMMARY TABLES
        # -----------------------------------------
        
        st.subheader("Sample Data")

        start_date_display = totals_df[DATE_COL].min()
        end_date_display = totals_df[DATE_COL].max()

        event_summary = pd.DataFrame({
            "Statistic": [
                "Monitoring Start",
                "Monitoring End",
                "PFAS Samples",
                "Pesticide Samples"
            ],
            "Value": [
                (
                    start_date_display.strftime("%d %b %Y")
                    if pd.notna(start_date_display)
                    else "N/A"
                ),
                (
                    end_date_display.strftime("%d %b %Y")
                    if pd.notna(end_date_display)
                    else "N/A"
                ),
                len(pfas_totals),
                len(pesticide_totals)
            ]
        })

        st.dataframe(
            event_summary,
            use_container_width=True,
            hide_index=True
        )


        col1, col2 = st.columns(2)

        with col1:

            st.subheader("PFAS Summary")

            if not pfas_totals.empty:

                pfas_summary = pd.DataFrame({

                    "Statistic": [
                        "Matched Components",
                        "Samples",
                        "Mean",
                        "P50",
                        "P95",
                        "Maximum",
                        "Limit Source",
                        "Limit",
                        "Breaches"
                    ],

                    "Value": [
                        len(matched_pfas),
                        len(pfas_totals),
                        f"{pfas_totals['PFAS Total'].mean():.3f} {total_units}",
                        f"{pfas_totals['PFAS Total'].quantile(0.50):.3f} {total_units}",
                        f"{pfas_totals['PFAS Total'].quantile(0.95):.3f} {total_units}",
                        f"{pfas_totals['PFAS Total'].max():.3f} {total_units}",
                        (
                            selected_company
                            if use_company_totals
                            else "Default PFAS Total Limit"
                        ),
                        f"{PFAS_LIMIT:.3f} {total_units}",
                        (
                            pfas_totals["PFAS Total"]
                            > ACTIVE_PFAS_LIMIT
                        ).sum()
                    ]
                })

                st.dataframe(
                    pfas_summary,
                    use_container_width=True,
                    hide_index=True
                )

        with col2:

            st.subheader("Pesticides Summary")

            if not pesticide_totals.empty:

                pesticide_summary = pd.DataFrame({

                    "Statistic": [
                        "Matched Components",
                        "Samples",
                        "Mean",
                        "P50",
                        "P95",
                        "Maximum",
                        "Limit Source",
                        "Limit",
                        "Breaches"
                    ],

                    "Value": [
                        len(matched_pesticides),
                        len(pesticide_totals),
                        f"{pesticide_totals['Pesticides Total'].mean():.3f} {total_units}",
                        f"{pesticide_totals['Pesticides Total'].quantile(0.50):.3f} {total_units}",
                        f"{pesticide_totals['Pesticides Total'].quantile(0.95):.3f} {total_units}",
                        f"{pesticide_totals['Pesticides Total'].max():.3f} {total_units}",
                        (
                            selected_company
                            if use_company_totals
                            else "Default Pesticide Total Limit"
                        ),
                        f"{ACTIVE_PESTICIDE_LIMIT:.3f} {total_units}",
                        (
                            pesticide_totals["Pesticides Total"]
                            > ACTIVE_PESTICIDE_LIMIT
                        ).sum()
                    ]
                })

                st.dataframe(
                    pesticide_summary,
                    use_container_width=True,
                    hide_index=True
                )
                
        st.subheader("All Data")        
        
        totals_display = totals_df.copy()

        if "PFAS Total" in totals_display.columns:
            totals_display["PFAS Total"] = (
                totals_display["PFAS Total"]
                .round(3)
                .astype(str)
                + f" {ACTIVE_PFAS_UNIT}"
            )

        if "Pesticides Total" in totals_display.columns:
            totals_display["Pesticides Total"] = (
                totals_display["Pesticides Total"]
                .round(3)
                .astype(str)
                + f" {total_units}"
            )

        st.dataframe(
            totals_display,
            use_container_width=True,
            hide_index=True
        )