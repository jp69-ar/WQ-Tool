import streamlit as st
import pandas as pd
import base64
import plotly.express as px
import plotly.graph_objects as go
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

st.title("Water Quality Reporting Tool")

st.markdown(
    """
    - Upload WQ data to generate report ready interactive figures and a data summary sheet.
    
    - Sites will likely require manual column mapping. Ensure the correct sheet/tab is corrected, and use the WQ data to map the correct columns within this app.
    
    - Note that when the units are contained within the parameter column (e.g. Turbidity (NTU)), the Auto Unit function can be used by selecting it within the Unit Column within column mapping, which separates the units for data analysis and presentation
    
    - Multiple water sources can be presented either within one worksheet, or split between multiple worksheets with one water source each - check which is applicable to your data.
    
    - This uses a default PCV and limit database - although your own database can be altered. Ensure that the same column formatting is used within PCVs.xlsb, which is contained within the source folder for this applet.
    
    - To see more information on the details of the complaince limit, hover over it.
    
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
        "Median": [data[result_col].median()],
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

uploaded_file = st.file_uploader(
    "Upload Water Quality Dataset",
    type=["xlsx", "csv"]
)

pcv_file = st.file_uploader(
    "Upload Regulatory Limits Database (optional)",
    type=["xlsx", "xlsb", "csv"]
)

if uploaded_file is not None:

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
    # ================================================
    # LOAD WATER QUALITY DATA
    # ================================================

    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            excel_file = pd.ExcelFile(uploaded_file)

            selected_sheet = st.selectbox(
                "Worksheet",
                excel_file.sheet_names
            )

            df = pd.read_excel(
                excel_file,
                sheet_name=selected_sheet
            )

    except Exception as e:
        st.error(
            f"Error loading file: {e}"
        )
        st.stop()
        
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
        "Column Mapping"
    ):

        DATE_COL = st.selectbox(
            "Date Column",
            df.columns,
            index=df.columns.get_loc(DATE_COL)
            if DATE_COL else 0
        )

        SITE_COL = st.selectbox(
            "Site Column",
            df.columns,
            index=df.columns.get_loc(SITE_COL)
            if SITE_COL else 0
        )

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
    "Site": SITE_COL,
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
    
    df = df.dropna(
        subset=[
            DATE_COL,
            SITE_COL,
            DET_COL,
            RESULT_COL
        ]
    )

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
            "All Data"
        ]
    )

    if quick_period == "Last 12 Months":
        default_start = max_date - timedelta(days=365)

    elif quick_period == "Last 3 Years":
        default_start = max_date - timedelta(days=1095)

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

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "Data",
            "Statistics",
            "Figure",
            "Exceedances",
            "Compliance Summary",
            "Export"
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
    # FIGURE TAB
    # =================================================

    with tab3:

        mean = filtered[RESULT_COL].mean()
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
                "No regulatory limits loaded."
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
                    f"{regulation} - {poc} ({limit_value} {limit_row.get('Units','')})"
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
                        "P95": round(p95_value, 3),
                        "Maximum": round(max_value, 3),
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

    # =======================================
    # EXPORT
    # =======================================

        with tab6:

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
                        
            selected_site = uploaded_file.name.rsplit(".", 1)[0]

            start_txt = pd.to_datetime(start_date).strftime("%Y%m%d")

            end_txt = pd.to_datetime(end_date).strftime("%Y%m%d")

            safe_component = (
                selected_component
                .replace("/", "-")
                .replace("\\", "-")
            )

            excel_filename = (
                f"{selected_site}_"
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
                f"{selected_site}_"
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
            


