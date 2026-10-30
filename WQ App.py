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
    page_title="Water Quality Reporting Tool",
    layout="wide"
)

st.markdown("""
<style>
:root {
    color-scheme: light;
}
</style>
""", unsafe_allow_html=True)

st.title("Water Quality Reporting Tool.")

st.markdown(
    """
    It is recommended to run this app in light mode (can change with 3 dots in the top right corner), and to zoom out slightly in your browser.
    
    - Upload WQ data to generate report ready interactive figures and a data summary sheet.
    
    - Sites will likely require manual column mapping. Ensure the correct sheet/tab is corrected, and use the WQ data to map the correct columns within this app.
    
    - Note that when the units are contained within the parameter column (e.g. Turbidity (NTU)), the Auto Unit function can be used by selecting it within the Unit Column within column mapping, which separates the units for data analysis and presentation
    
    - Multiple water sources can be presented either within one worksheet, or split between multiple worksheets with one water source each - check which is applicable to your data.
    
    - This uses a default PCV and limit database - although your own database can be altered. Ensure that the same column formatting is used within PCVs.xlsb, which is contained within the source folder for this applet.
    
    - To see more information on the details of the compliance limit, hover over it.
    
    """
)

# =====================================================
# FUNCTIONS
# =====================================================

def statistics_table(data, result_col, selected_component):

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
        "Samples": [len(data)],
        "Breaches": [breach_count],
        "% Exceeding": [
            100 * breach_count / len(data)
            if len(data) > 0 else 0
        ],
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
    component_limits
):

    output = BytesIO()

    summary_df = pd.DataFrame({
        "Item": [
            "Source File",
            "Water Source",
            "Component",
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

        for _, limit_row in component_limits.iterrows():

            limit_value = pd.to_numeric(
                limit_row["Limit Value"],
                errors="coerce"
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
                        limit_row["Units"],
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
        Maximum recorded value: {maximum}<br/><br/>
        <u>Recorded Exceedances</u>
        """

        content.append(
            Paragraph(summary_text, styles["BodyText"])
        )

        for _, limit_row in component_limits.iterrows():

            limit_value = pd.to_numeric(
                limit_row["Limit Value"],
                errors="coerce"
            )

            if pd.notna(limit_value):

                exceedances = filtered[
                    filtered[RESULT_COL] > limit_value
                ]

                regulation = limit_row["Regulatory Source"]
                poc = limit_row.get("Point of Compliance", "")
                units = limit_row.get("Units", "")

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
            name=f"Mean ({mean:.2f})",
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
            name=f"P50 ({p50:.2f})",
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
            name=f"P95 ({p95:.2f})",
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
    
# =====================================================
# FILE UPLOAD
# =====================================================

uploaded_files = st.file_uploader(
    "Upload Water Quality Dataset(s)",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

pcv_file = st.file_uploader(
    "Upload Regulatory Limits Database (optional)",
    type=["xlsx", "xlsb", "csv"]
)

loaded_datasets = {}

if uploaded_files:
    
    for uploaded_file in uploaded_files:

        try:

            if uploaded_file.name.lower().endswith(".csv"):

                loaded_datasets[
                    uploaded_file.name
                ] = pd.read_csv(uploaded_file)

            else:

                excel_file = pd.ExcelFile(
                    uploaded_file
                )

                for sheet in excel_file.sheet_names:

                    loaded_datasets[
                        f"{uploaded_file.name} | {sheet}"
                    ] = pd.read_excel(
                        excel_file,
                        sheet_name=sheet
                    )

        except Exception as e:

            st.error(
                f"Error loading {uploaded_file.name}: {e}"
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

            st.sidebar.success(
                "Using uploaded PCV database"
            )

        # Otherwise use bundled database
        else:

            pcv_df = pd.read_excel(
                "PCVs.xlsb",
                engine="pyxlsb"
            )

            st.sidebar.success(
                "Currently using default PCV database file"
            )

    except Exception as e:

        st.sidebar.error(
            f"Could not load PCV database: {e}"
        )

        pcv_df = None


    if pcv_df is not None:

        st.sidebar.caption(
            f"{len(pcv_df)} limits loaded"
        )
        
    if pcv_df is not None:

        pcv_df["Limit Value"] = pd.to_numeric(
            pcv_df["Limit Value"],
            errors="coerce"
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

                pfas_df = pd.read_excel(
                    "PCVs.xlsb",
                    sheet_name="PFAS",
                    engine="pyxlsb"
                )

                pesticide_df = pd.read_excel(
                    "PCVs.xlsb",
                    sheet_name="Pesticides",
                    engine="pyxlsb"
                )

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
        
    import re

    def clean_component_name(text):
        text = str(text).lower()

        # remove brackets and contents
        text = re.sub(r"\(.*?\)", "", text)

        # replace separators
        text = text.replace("/", " ")
        text = text.replace("-", " ")

        # collapse spaces
        text = " ".join(text.split())

        # standardise metals
        text = text.replace("iron", "iron")
        text = text.replace("aluminium", "aluminium")
        text = text.replace("manganese", "manganese")

        return text.strip()

    PFAS_COMPONENTS = set()
    PESTICIDE_COMPONENTS = set()

    PFAS_LIMIT = np.nan
    PESTICIDE_LIMIT = np.nan

    if not pfas_df.empty:

        PFAS_COMPONENTS = {
            clean_component_name(x)
            for x in pfas_df.iloc[:, 0].dropna()
            if str(x).lower() != "total limit"
        }

        try:
            PFAS_LIMIT = float(
                pfas_df.loc[
                    pfas_df.iloc[:,0]
                    .astype(str)
                    .str.lower()
                    == "total limit",
                    pfas_df.columns[1]
                ].iloc[0]
            )
        except:
            pass

    if not pesticide_df.empty:

        PESTICIDE_COMPONENTS = {
            clean_component_name(x)
            for x in pesticide_df.iloc[:, 0].dropna()
            if str(x).lower() != "total limit"
        }

        try:
            PESTICIDE_LIMIT = float(
                pesticide_df.loc[
                    pesticide_df.iloc[:,0]
                    .astype(str)
                    .str.lower()
                    == "total limit",
                    pesticide_df.columns[1]
                ].iloc[0]
            )
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
        
    selected_dataset = st.selectbox(
        "Main Dataset",
        list(loaded_datasets.keys())
    )

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
    
    # If no unit column exists, extract units from parameter names

    if UNIT_COL is None:

        df["Auto_Unit"] = ""

        if (
            DET_COL is not None
            and pd.notna(DET_COL)
            and DET_COL in df.columns
        ):

            df["Auto_Unit"] = (
                df[DET_COL]
                .astype(str)
                .str.extract(r"\(([^)]*)\)", expand=False)
                .fillna("")
            )

        UNIT_COL = "Auto_Unit"

    original_count = len(df)
    
    # =====================================================
    # MANUAL COLUMN OVERRIDE
    # =====================================================

    with st.expander(
        "Column Mapping (main statistics/figure)"
    ):

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

    df[DATE_COL] = pd.to_datetime(
        df[DATE_COL],
        errors="coerce"
    )

    # Preserve original result text
    df["Original_Result"] = df[RESULT_COL].astype(str)

    # Flag results reported with <
    df["Non_Detect"] = (
        df["Original_Result"]
        .str.strip()
        .str.startswith("<")
    )

    # Extract numeric component
    df[RESULT_COL] = (
        df["Original_Result"]
        .astype(str)
        .str.replace("<", "", regex=False)
    )

    df[RESULT_COL] = pd.to_numeric(
        df[RESULT_COL],
        errors="coerce"
    )

    # Apply half-LOD substitution for non-detects
    df.loc[
        df["Non_Detect"],
        RESULT_COL
    ] = (
        df.loc[
            df["Non_Detect"],
            RESULT_COL
        ] / 2
    )
    
    required_cols = [
        DATE_COL,
        DET_COL,
        RESULT_COL
    ]

    if SITE_COL is not None:
        required_cols.append(SITE_COL)

    df = df.dropna(subset=required_cols)

    ph_mask = (
        df[DET_COL]
        .str.contains(
            "Hydrogen",
            case=False,
            na=False
        )
    )

    df = df[
        ~(ph_mask & (df[RESULT_COL] <= 0))
    ]
    
    # Remove zero and negative results

    before_zero_removal = len(df)

    df = df[
        df[RESULT_COL] > 0
    ]

    zero_removed = (
        before_zero_removal
        - len(df)
    )
    
    cleaned_count = len(df)

    qa_removed = (
        original_count -
        cleaned_count
    )

    # =================================================
    # SIDEBAR
    # =================================================

    st.sidebar.header("Filters")

    selected_site = st.sidebar.selectbox(
        "Water Source",
        sorted(df[SITE_COL].unique())
    )

    filtered = df[
        df[SITE_COL] == selected_site
    ].copy()
    
    filtered["Component_Clean"] = (
        filtered[DET_COL]
        .apply(clean_component_name)
    )

    component_lookup = (
        filtered.groupby("Component_Clean")[DET_COL]
        .agg(lambda x: x.mode().iloc[0])
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
    
    selected_regulations = st.sidebar.multiselect(
        "Regulations",
        [
            "DWI",
            "EU rDWD",
            "WHO"
        ],
        default=["DWI"]
    )

    min_date = filtered[DATE_COL].min().date()
    max_date = filtered[DATE_COL].max().date()

    quick_period = st.sidebar.radio(
        "Quick Period",
        [
            "Custom",
            "Last 12 Months",
            "Last 3 Years",
            "Last 5 years",
            "All Data"
        ]
    )

    if quick_period == "Last 12 Months":
        default_start = max_date - timedelta(days=365)

    elif quick_period == "Last 3 Years":
        default_start = max_date - timedelta(days=1095)
        
    elif quick_period == "Last 5 Years":
        default_start = max_date - timedelta(days=1825)

    else:
        default_start = min_date

    start_date = st.sidebar.date_input(
        "Start Date",
        value=default_start
    )

    end_date = st.sidebar.date_input(
        "End Date",
        value=max_date
    )

    filtered = filtered[
        filtered["Component_Clean"]
        == selected_component
    ]

    filtered = filtered[
        (filtered[DATE_COL] >= pd.to_datetime(start_date))
        &
        (filtered[DATE_COL] <= pd.to_datetime(end_date))
    ]

    component_limits = pd.DataFrame()

    if pcv_df is not None:

        component_limits = get_component_limits(
            pcv_df,
            selected_component
        )

        component_limits = component_limits[
            component_limits["Regulatory Source"]
            .isin(selected_regulations)
        ]
       
    st.write(selected_component)

    st.write("Matches found:")
    st.write(len(component_limits))

    # =================================================
    # SUMMARY
    # =================================================

    monitoring_start = df[DATE_COL].min()
    monitoring_end = df[DATE_COL].max()

    c1, c2, c3, c4, c5 = st.columns(5)

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

    c4.metric(
        "QA Removed",
        f"{qa_removed:,}"
    )

    c5.metric(
        "Zero Values Removed",
        f"{zero_removed:,}"
    )

    st.info(
        f"""
        Monitoring Period:
        {monitoring_start.strftime('%d %b %Y')}
        to
        {monitoring_end.strftime('%d %b %Y')}
        """
    )

    # =================================================
    # TABS
    # =================================================

    tab1, tab2, tab2a, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
        [
            "Component Data",
            "Component Statistics",
            "Custom Statistics",
            "Component Figure",
            "Component Exceedances",
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

        if UNIT_COL != "No Unit Column":
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
            selected_component
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
    # CUSTOM SUMMARY BUILDER
    # =================================================

    with tab2a:

        site_df = df[
            (df[SITE_COL] == selected_site)
            &
            (df[DATE_COL] >= pd.to_datetime(start_date))
            &
            (df[DATE_COL] <= pd.to_datetime(end_date))
        ].copy()

        site_df["Component_Clean"] = (
            site_df[DET_COL]
            .apply(clean_component_name)
        )
  
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

                    limits = get_component_limits(
                        pcv_df,
                        component
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

    # =================================================
    # FIGURE TAB
    # =================================================

    with tab3:

        mean = filtered[RESULT_COL].mean()
        p50 = filtered[RESULT_COL].quantile(0.50)
        p95 = filtered[RESULT_COL].quantile(0.95)

        units = ""

        if len(filtered):
               if UNIT_COL == "No Unit Column":
                    units = ""
               else:
                    units = filtered[UNIT_COL].iloc[0]
                        
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
                name=f"Mean ({mean:.2f})",
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
                name=f"P50 ({p50:.2f})",
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
                name=f"P95 ({p95:.2f})",
                line=dict(
                    color="#FF00FF",
                    dash="dash"
                ),
                hovertemplate=f"P95: {p95:.2f}<extra></extra>"
            )
        )


        if not component_limits.empty:

            colours = {
                "DWI": "red",
                "EU rDWD": "blue",
                "WHO": "green"
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
                

                limit_value = pd.to_numeric(
                    row["Limit Value"],
                    errors="coerce"
                )

                if pd.isna(limit_value):
                    continue

                poc = row.get("Point of Compliance", "")
                
                regulation = row.get("Regulatory Source", "")
                                
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
                        name=f"{regulation} - {poc} ({limit_value} {units})",
                        line=dict(
                            color=line_colour,
                            width=3,
                            dash="dash"
                        ),
                        hovertemplate=(
                            f"{regulation} - {poc}<br>"
                            f"Limit: {limit_value} {units}"
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
                    f"Limit: {limit_value} {row.get('Units','')}"
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
                if units and f"({units})" in str(selected_component)
                else (
                    f"{selected_component} ({units})"
                    if units
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
            use_container_width=True
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

        fig.update_layout(
        width=2200,
        height=1200,
        font=dict(size=20),
        title_font=dict(size=30),
        margin=dict(
            l=100,
            r=150,
            t=120,
            b=100
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        )
    )

    png_bytes = fig.to_image(
        format="png",
        width=2200,
        height=1200,
        scale=2
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

                limit_value = pd.to_numeric(
                    limit_row["Limit Value"],
                    errors="coerce"
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
                    f"({limit_value} {limit_row.get('Units','')})"
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
    # COMPLIANCE SUMMARY TAB
    # =================================================

    with tab5:

        st.subheader(
            "Compliance Summary by Water Source"
        )

        summary_rows = []

        for source in sorted(
            df[SITE_COL].dropna().unique()
        ):

            source_df = df[
                df[SITE_COL] == source
            ]
            
            source_df = source_df.copy()

            source_df["Component_Clean"] = (
                source_df[DET_COL]
                .apply(clean_component_name)
            )

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
                    component_df[RESULT_COL],
                    errors="coerce"
                ).dropna()

                if len(results) == 0:
                    continue

                min_value = results.min()
                mean_value = results.mean()
                p50_value = results.quantile(0.50)
                p95_value = results.quantile(0.95)
                max_value = results.max()

                breaches = 0
                limit_value = None

                if pcv_df is not None:

                    limits = get_component_limits(
                        pcv_df,
                        component
                    )


                    limits = limits[
                        limits["Regulatory Source"]
                        .isin(selected_regulations)
                    ]

                    poc = ""
                    
                    regulation_source = ""

                    if not limits.empty:
                        limit_value = pd.to_numeric(
                            limits["Limit Value"],
                            errors="coerce"
                        ).min()

                        numeric_limits = pd.to_numeric(
                            limits["Limit Value"],
                            errors="coerce"
                        )

                        if numeric_limits.notna().any():

                            limit_idx = numeric_limits.idxmin()
                            
                            limit_row = limits.loc[limit_idx]

                            poc = limit_row.get(
                                "Point of Compliance",
                                ""
                            )

                            regulation_source = limit_row.get(
                                "Regulatory Source",
                                ""
)

                        else:

                            poc = ""

                        if pd.notna(limit_value):
                            breaches = (
                                results > limit_value
                            ).sum()
                        else:
                            breaches = 0

                units = ""

                if UNIT_COL:
                    unit_values = (
                        component_df[UNIT_COL]
                        .dropna()
                        .astype(str)
                    )

                    if not unit_values.empty:
                        units = unit_values.mode().iloc[0]
                
                display_name = (
                    component_df[DET_COL]
                    .mode()
                    .iloc[0]
                )
                
                summary_rows.append(
                    {
                        "Water Source": source,
                        "Component": display_name,
                        "Units": units,
                        "Samples": len(results),
                        "Minimum": round(min_value, 3),
                        "Mean": round(mean_value, 3),
                        "P50": round(p50_value, 3),
                        "P95": round(p95_value, 3),
                        "Maximum": round(max_value, 3),
                        "Limit Source": regulation_source,
                        "Limit": limit_value,
                        "Point of Compliance": poc,
                        "Breaches": breaches,
                        "% Exceeding": round(
                            100 * (breaches / len(results)),
                            2
                        ) if len(results) > 0 else 0
                    }
                )

        summary_df = pd.DataFrame(summary_rows)

        st.dataframe(
            summary_df,
            use_container_width=True
        )


    # =================================================
    # COMPARISON TAB
    # =================================================

    with tab6:

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

        # ---------------------------------------
        # COLUMN MAPPING
        # ---------------------------------------

        c1, c2 = st.columns(2)

        with c1:

            with st.expander("Source A Column Mapping"):

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
        with c2:

            with st.expander("Source B Column Mapping"):

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

        # =====================================
        # SOURCE B CLEANING
        # =====================================

        compare_df_b["Original_Result"] = (
            compare_df_b[RESULT_B].astype(str)
        )

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

        # ---------------------------------------
        # COMPONENTS
        # ---------------------------------------

        c1, c2 = st.columns(2)

        with c1:
            
            compare_df_a = compare_df_a.dropna(
                subset=[COMPONENT_A]
            )

            compare_df_a["Component_Clean"] = (
                compare_df_a[COMPONENT_A]
                .apply(clean_component_name)
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

            compare_df_b["Component_Clean"] = (
                compare_df_b[COMPONENT_B]
                .apply(clean_component_name)
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
            unit_values = compare_a[UNIT_A].dropna()

            if not unit_values.empty:
                units_a = str(unit_values.mode().iloc[0])

        units_b = ""

        if (
            UNIT_B != "No Unit Column"
            and UNIT_B in compare_b.columns
        ):
            unit_values = compare_b[UNIT_B].dropna()

            if not unit_values.empty:
                units_b = str(unit_values.mode().iloc[0])

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
        # FIGURE
        # ---------------------------------------

        mean_a = compare_a[RESULT_A].mean()
        p50_a = compare_a[RESULT_A].quantile(0.50)
        p95_a = compare_a[RESULT_A].quantile(0.95)

        mean_b = compare_b[RESULT_B].mean()
        p50_b = compare_b[RESULT_B].quantile(0.50)
        p95_b = compare_b[RESULT_B].quantile(0.95)
        
        limits_a = pd.DataFrame()
        limits_b = pd.DataFrame()

        if pcv_df is not None:

            limits_a = get_component_limits(
                pcv_df,
                component_a
            )

            limits_b = get_component_limits(
                pcv_df,
                component_b
            )

            limits_a = limits_a[
                limits_a["Regulatory Source"]
                .isin(selected_regulations)
            ]

            limits_b = limits_b[
                limits_b["Regulatory Source"]
                .isin(selected_regulations)
            ]

        fig_compare = make_subplots(
            rows=1,
            cols=2,
            shared_yaxes=True,
            subplot_titles=(title_a, title_b),
            horizontal_spacing=0.03
        )

        # SOURCE A DATA

        fig_compare.add_trace(
            go.Scatter(
                x=compare_a[DATE_A],
                y=compare_a[RESULT_A],
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

        for value, colour, dash, label in [
            (mean_a, "#000000", "dot", f"{title_a} Mean ({mean_a:.2f})"),
            (p50_a, "#4B0082", "dash", f"{title_a} P50 ({p50_a:.2f})"),
            (p95_a, "#FF00FF", "dash", f"{title_a} P95 ({p95_a:.2f})")
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
            
        for _, row in limits_a.iterrows():

            limit_value = pd.to_numeric(
                row["Limit Value"],
                errors="coerce"
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

            fig_compare.add_trace(
                go.Scatter(
                    x=[
                        compare_a[DATE_A].min(),
                        compare_a[DATE_A].max()
                    ],
                    y=[limit_value, limit_value],
                    mode="lines",
                    name=f"{regulation} - {poc} ({limit_value})",
                    line=dict(
                        width=3,
                        dash="dash"
                    )
                ),
                row=1,
                col=1
            )

        # SOURCE B DATA

        fig_compare.add_trace(
            go.Scatter(
                x=compare_b[DATE_B],
                y=compare_b[RESULT_B],
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

        for value, colour, dash, label in [
            (mean_b, "#000000", "dot", f"{title_b} Mean ({mean_b:.2f})"),
            (p50_b, "#4B0082", "dash", f"{title_b} P50 ({p50_b:.2f})"),
            (p95_b, "#FF00FF", "dash", f"{title_b} P95 ({p95_b:.2f})")
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

            limit_value = pd.to_numeric(
                row["Limit Value"],
                errors="coerce"
            )

            if pd.isna(limit_value):
                continue

            fig_compare.add_trace(
                go.Scatter(
                    x=[
                        compare_b[DATE_B].min(),
                        compare_b[DATE_B].max()
                    ],
                    y=[limit_value, limit_value],
                    mode="lines",
                    showlegend=False,
                    line=dict(
                        width=3,
                        dash="dash"
                    )
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
            showgrid=True,
            showline=True,
            linewidth=2,
            linecolor="black",
            gridcolor="#D9D9D9"
        )

        fig_compare.update_yaxes(
            title_text=(
                f"{component_display_a} ({units_a})"
                if units_a
                else component_display_a
            ),
            showgrid=True,
            showline=True,
            linewidth=2,
            linecolor="black",
            gridcolor="#D9D9D9"
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

        with tab7:

            excel_data = build_export(
                stats_df,
                filtered,
                selected_site,
                selected_component,
                start_date,
                end_date,
                uploaded_file.name,
                component_limits
            )
            
            import tempfile

            png_bytes = fig.to_image(
                format="png",
                scale=3
            )

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".png"
            ) as tmp:

                tmp.write(png_bytes)

                chart_path = tmp.name          
            
            pdf_data = build_pdf_report(
                stats_df,
                summary_df,
                selected_site,
                selected_component,
                start_date,
                end_date,
                uploaded_file.name,
                chart_path
            )
                        
            export_site = uploaded_file.name.rsplit(".", 1)[0]

            start_txt = pd.to_datetime(start_date).strftime("%Y%m%d")

            end_txt = pd.to_datetime(end_date).strftime("%Y%m%d")

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

            st.download_button(
                "Download Excel Report",
                data=excel_data,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            pdf_filename = (
                f"{export_site}_"
                f"{safe_component}_"
                f"{start_txt}_to_{end_txt}.pdf"
            )

            st.download_button(
                "Download PDF Report",
                data=pdf_data,
                file_name=pdf_filename,
                mime="application/pdf"
            )

            file_name=excel_filename
            

# =================================================
# PFAS / PESTICIDE TOTALS
# =================================================

with tab8:

    st.subheader("PFAS & Pesticide Totals")

    site_df = df[
        (df[SITE_COL] == selected_site)
        &
        (df[DATE_COL] >= pd.to_datetime(start_date))
        &
        (df[DATE_COL] <= pd.to_datetime(end_date))
    ].copy()

    site_df["Component_Clean"] = (
        site_df[DET_COL]
        .apply(clean_component_name)
    )

    # -----------------------------------------
    # PFAS / PESTICIDE MATCHING
    # -----------------------------------------

    pfas_results = site_df[
        site_df["Component_Clean"]
        .isin(PFAS_COMPONENTS)
    ].copy()

    pesticide_results = site_df[
        site_df["Component_Clean"]
        .isin(PESTICIDE_COMPONENTS)
    ].copy()

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

    col1, col2 = st.columns(2)

    with col1:

        with st.expander(
            f"Matched PFAS Components ({len(matched_pfas)})",
            expanded=False
        ):

            st.dataframe(
                pd.DataFrame({
                    "PFAS Components": matched_pfas
                }),
                use_container_width=True,
                hide_index=True
            )

    with col2:

        with st.expander(
            f"Matched Pesticides ({len(matched_pesticides)})",
            expanded=False
        ):

            st.dataframe(
                pd.DataFrame({
                    "Pesticide Components": matched_pesticides
                }),
                use_container_width=True,
                hide_index=True
            )

    # -----------------------------------------
    # TOTALS
    # -----------------------------------------

    pfas_totals = (
        pfas_results
        .groupby(DATE_COL)[RESULT_COL]
        .sum()
        .reset_index()
        .rename(
            columns={
                RESULT_COL: "PFAS Total"
            }
        )
    )

    pesticide_totals = (
        pesticide_results
        .groupby(DATE_COL)[RESULT_COL]
        .sum()
        .reset_index()
        .rename(
            columns={
                RESULT_COL: "Pesticides Total"
            }
        )
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

    if pd.notna(PFAS_LIMIT):

        fig_totals.add_trace(
            go.Scatter(
                x=[
                    totals_df[DATE_COL].min(),
                    totals_df[DATE_COL].max()
                ],
                y=[
                    PFAS_LIMIT,
                    PFAS_LIMIT
                ],
                mode="lines",
                name=(
                    f"PFAS Limit "
                    f"({PFAS_LIMIT:.3f} {total_units})"
                ),
                line=dict(
                    color="purple",
                    width=3,
                    dash="dash"
                )
            )
        )

    if pd.notna(PESTICIDE_LIMIT):

        fig_totals.add_trace(
            go.Scatter(
                x=[
                    totals_df[DATE_COL].min(),
                    totals_df[DATE_COL].max()
                ],
                y=[
                    PESTICIDE_LIMIT,
                    PESTICIDE_LIMIT
                ],
                mode="lines",
                name=(
                    f"Pesticide Limit "
                    f"({PESTICIDE_LIMIT:.3f} {total_units})"
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
        ),
        xaxis_title="Sample Date",
        yaxis_title=f"Total Concentration ({total_units})",
        legend_title="Series"
    )

    st.plotly_chart(
        fig_totals,
        use_container_width=True
    )

    # -----------------------------------------
    # SUMMARY TABLES
    # -----------------------------------------

    event_summary = pd.DataFrame({
        "Statistic": [
            "Monitoring Start",
            "Monitoring End",
            "PFAS Components Included",
            "Pesticide Components Included",
            "PFAS Samples",
            "Pesticide Samples"
        ],
        "Value": [
            totals_df[DATE_COL].min().strftime("%d %b %Y"),
            totals_df[DATE_COL].max().strftime("%d %b %Y"),
            len(matched_pfas),
            len(matched_pesticides),
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
                    f"{PFAS_LIMIT:.3f} {total_units}",
                    (
                        pfas_totals["PFAS Total"]
                        > PFAS_LIMIT
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
                    f"{PESTICIDE_LIMIT:.3f} {total_units}",
                    (
                        pesticide_totals["Pesticides Total"]
                        > PESTICIDE_LIMIT
                    ).sum()
                ]
            })

            st.dataframe(
                pesticide_summary,
                use_container_width=True,
                hide_index=True
            )
    
    totals_display = totals_df.copy()

    if "PFAS Total" in totals_display.columns:
        totals_display["PFAS Total"] = (
            totals_display["PFAS Total"]
            .round(3)
            .astype(str)
            + f" {total_units}"
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