from pathlib import Path
import zipfile

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="ShopMetrics Pro",
    page_icon="🛍️",
    layout="wide",
)

st.title("🛍️ ShopMetrics Pro")
st.caption("আপনার Excel বা CSV ডেটা সহজে সাজান, সারাংশ দেখুন এবং বিশ্লেষণ করুন।")

st.sidebar.header("ডেটা ও প্রদর্শন")
theme_mode = st.sidebar.selectbox(
    "ড্যাশবোর্ড থিম",
    options=["লাইট মোড", "ডার্ক মোড"],
    key="theme_mode",
)

is_dark_mode = theme_mode == "ডার্ক মোড"
background_color = "#0E1117" if is_dark_mode else "#FFFFFF"
secondary_background_color = "#262730" if is_dark_mode else "#F0F2F6"
text_color = "#FAFAFA" if is_dark_mode else "#31333F"
chart_template = "plotly_dark" if is_dark_mode else "plotly_white"

st.markdown(
    f"""
    <style>
        [data-testid="stAppViewContainer"],
        [data-testid="stHeader"] {{
            background-color: {background_color};
        }}
        [data-testid="stSidebar"] {{
            background-color: {secondary_background_color};
        }}
        [data-testid="stAppViewContainer"] [data-testid="stMarkdownContainer"],
        [data-testid="stAppViewContainer"] [data-testid="stWidgetLabel"],
        [data-testid="stAppViewContainer"] [data-testid="stMetricLabel"],
        [data-testid="stAppViewContainer"] [data-testid="stMetricValue"],
        [data-testid="stAppViewContainer"] [data-testid="stCaptionContainer"],
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{
            color: {text_color};
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.sidebar.file_uploader(
    "Excel বা CSV ফাইল আপলোড করুন",
    type=["xlsx", "xls", "csv"],
    help="Excel workbook হলে নিচে তার একটি sheet নির্বাচন করতে পারবেন।",
)

if uploaded_file is None:
    st.info("শুরু করতে সাইডবার থেকে একটি Excel বা CSV ফাইল আপলোড করুন।")
    st.stop()


def make_unique_columns(columns: pd.Index) -> list[str]:
    """Create readable, unique column names for reliable downstream selection."""
    unique_names = []
    seen_names: dict[str, int] = {}

    for position, raw_name in enumerate(columns, start=1):
        name = str(raw_name).strip() or f"Column {position}"
        occurrence = seen_names.get(name, 0)
        seen_names[name] = occurrence + 1
        unique_names.append(name if occurrence == 0 else f"{name} ({occurrence + 1})")

    return unique_names


def prepare_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    dataframe = dataframe.dropna(how="all").dropna(axis=1, how="all").copy()
    dataframe.columns = make_unique_columns(dataframe.columns)

    for column in dataframe.columns:
        column_dtype = dataframe[column].dtype
        if not (
            pd.api.types.is_object_dtype(column_dtype)
            or pd.api.types.is_string_dtype(column_dtype)
        ):
            continue

        values = dataframe[column].astype("string").str.strip()
        non_empty = values.notna() & values.ne("")
        if not non_empty.any():
            dataframe[column] = values
            continue

        column_hint = column.casefold()
        numeric_hint = any(
            word in column_hint
            for word in (
                "sales",
                "revenue",
                "price",
                "amount",
                "total",
                "cost",
                "quantity",
                "qty",
                "count",
                "profit",
                "balance",
            )
        )
        numeric_values = values.str.replace(",", "", regex=False)
        if numeric_hint:
            numeric_values = numeric_values.str.replace(
                r"[$£€৳₹]", "", regex=True
            )
        parsed_numbers = pd.to_numeric(numeric_values, errors="coerce")

        if parsed_numbers[non_empty].notna().mean() >= 0.9:
            dataframe[column] = parsed_numbers
            continue

        if any(
            word in column_hint
            for word in (
                "date",
                "time",
                "month",
                "year",
                "created",
                "updated",
                "timestamp",
                "day",
            )
        ):
            parsed_dates = pd.to_datetime(values, errors="coerce")
            if parsed_dates[non_empty].notna().mean() >= 0.8:
                dataframe[column] = parsed_dates
                continue

        dataframe[column] = values

    return dataframe.reset_index(drop=True)


file_extension = Path(uploaded_file.name).suffix.casefold()
selected_sheet = None

try:
    if file_extension == ".csv":
        raw_df = pd.read_csv(uploaded_file)
    else:
        workbook = pd.ExcelFile(uploaded_file)
        if not workbook.sheet_names:
            st.error("Excel workbook-এ কোনো sheet পাওয়া যায়নি।")
            st.stop()
        selected_sheet = st.sidebar.selectbox(
            "Excel sheet নির্বাচন করুন",
            options=workbook.sheet_names,
        )
        raw_df = pd.read_excel(workbook, sheet_name=selected_sheet)
except ImportError:
    if file_extension == ".xlsx":
        st.error("XLSX ফাইল পড়তে openpyxl ইনস্টল করুন: pip install openpyxl")
    elif file_extension == ".xls":
        st.error("পুরোনো XLS ফাইল পড়তে xlrd ইনস্টল করুন: pip install xlrd")
    else:
        st.error("CSV ফাইলটি পড়ার encoding বা format-এ সমস্যা হয়েছে।")
    st.stop()
except (
    UnicodeDecodeError,
    pd.errors.EmptyDataError,
    pd.errors.ParserError,
    ValueError,
    OSError,
    zipfile.BadZipFile,
) as error:
    st.error(f"ফাইলটি পড়া যায়নি। ফাইলের ধরন ও ডেটা যাচাই করুন। বিস্তারিত: {error}")
    st.stop()

df = prepare_dataframe(raw_df)
if df.empty:
    st.error("নির্বাচিত ফাইলে বা sheet-এ দেখানোর মতো কোনো ডেটা নেই।")
    st.stop()

st.sidebar.caption(f"ডেটার উৎস: {uploaded_file.name}")
if selected_sheet is not None:
    st.sidebar.caption(f"নির্বাচিত sheet: {selected_sheet}")

numeric_columns = df.select_dtypes(include="number").columns.tolist()
datetime_columns = df.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()
categorical_columns = [
    column
    for column in df.select_dtypes(include=["object", "string", "category"]).columns
    if df[column].nunique(dropna=True) <= 100
]

filtered_df = df.copy()

if categorical_columns:
    preferred_categories = [
        column
        for column in categorical_columns
        if any(word in column.casefold() for word in ("category", "type", "status"))
    ]
    filter_column = st.sidebar.selectbox(
        "ক্যাটাগরি/গ্রুপ ফিল্টার",
        options=["কোনো ফিল্টার নয়"] + categorical_columns,
        index=(
            categorical_columns.index(preferred_categories[0]) + 1
            if preferred_categories
            else 0
        ),
    )
    if filter_column != "কোনো ফিল্টার নয়":
        filter_options = sorted(
            df[filter_column].dropna().astype(str).unique().tolist()
        )
        selected_values = st.sidebar.multiselect(
            f"{filter_column} নির্বাচন করুন",
            options=filter_options,
            default=filter_options,
        )
        filtered_df = filtered_df[
            filtered_df[filter_column].astype(str).isin(selected_values)
        ]

if datetime_columns:
    date_column = st.sidebar.selectbox("তারিখ কলাম", options=datetime_columns)
    date_values = filtered_df[date_column].dropna()
    if date_values.empty:
        st.sidebar.caption("নির্বাচিত ডেটায় তারিখের মান নেই।")
    else:
        earliest_date = date_values.min().date()
        latest_date = date_values.max().date()
        selected_dates = st.sidebar.date_input(
            "তারিখের পরিসর",
            value=(earliest_date, latest_date),
            min_value=earliest_date,
            max_value=latest_date,
        )
        if isinstance(selected_dates, tuple):
            start_date, end_date = selected_dates
        else:
            start_date = end_date = selected_dates
        filtered_df = filtered_df[
            filtered_df[date_column].dt.date.between(start_date, end_date)
        ]

st.subheader("📌 ডেটার সারসংক্ষেপ")
summary_columns = st.columns(4)
summary_columns[0].metric("মোট সারি", f"{len(filtered_df):,}")
summary_columns[1].metric("মোট কলাম", f"{len(df.columns):,}")
summary_columns[2].metric("সংখ্যাসূচক কলাম", f"{len(numeric_columns):,}")
summary_columns[3].metric(
    "ফাঁকা ঘর",
    f"{int(filtered_df.isna().sum().sum()):,}",
)

if filtered_df.empty:
    st.info("এই ফিল্টারে কোনো ডেটা নেই। সাইডবারের ফিল্টার পরিবর্তন করুন।")
    st.stop()

if numeric_columns:
    st.subheader("📈 সংখ্যাসূচক পরিসংখ্যান")
    selected_numeric = st.selectbox(
        "যে সংখ্যাসূচক কলামের সারাংশ দেখতে চান",
        options=numeric_columns,
    )
    numeric_series = filtered_df[selected_numeric].dropna()
    if numeric_series.empty:
        st.info(f"{selected_numeric} কলামে কোনো সংখ্যাসূচক মান নেই।")
    else:
        metric_columns = st.columns(4)
        metric_columns[0].metric("মোট", f"{numeric_series.sum():,.2f}")
        metric_columns[1].metric("গড়", f"{numeric_series.mean():,.2f}")
        metric_columns[2].metric("সর্বোচ্চ", f"{numeric_series.max():,.2f}")
        metric_columns[3].metric("সর্বনিম্ন", f"{numeric_series.min():,.2f}")

        with st.expander("সম্পূর্ণ পরিসংখ্যান দেখুন"):
            st.dataframe(
                filtered_df[numeric_columns].describe().T,
                width="stretch",
            )

st.subheader("📊 ইন্টারেক্টিভ চার্ট")
chart_dimensions = categorical_columns + [
    column for column in datetime_columns if column not in categorical_columns
]

if chart_dimensions and numeric_columns:
    chart_controls = st.columns(3)
    x_column = chart_controls[0].selectbox(
        "গ্রুপ/সময় কলাম",
        options=chart_dimensions,
    )
    y_column = chart_controls[1].selectbox(
        "মানের কলাম",
        options=numeric_columns,
    )
    preferred_chart = "Line" if x_column in datetime_columns else "Bar"
    chart_type = chart_controls[2].selectbox(
        "চার্টের ধরন",
        options=["Bar", "Line", "Pie"],
        index=["Bar", "Line", "Pie"].index(preferred_chart),
    )

    if x_column in datetime_columns:
        chart_data = (
            filtered_df.groupby(x_column, as_index=False, dropna=False)[y_column]
            .sum()
            .sort_values(x_column)
        )
    else:
        chart_data = (
            filtered_df.groupby(x_column, as_index=False, dropna=False)[y_column]
            .sum()
            .sort_values(y_column, ascending=False)
        )

    if chart_type == "Bar":
        chart_data = chart_data.head(30)
        chart = px.bar(
            chart_data,
            x=x_column,
            y=y_column,
            title=f"{x_column} অনুযায়ী {y_column}",
            labels={x_column: x_column, y_column: y_column},
        )
    elif chart_type == "Line":
        chart = px.line(
            chart_data,
            x=x_column,
            y=y_column,
            markers=True,
            title=f"{x_column} অনুযায়ী {y_column}",
            labels={x_column: x_column, y_column: y_column},
        )
    else:
        chart_data = chart_data.head(12)
        chart = px.pie(
            chart_data,
            names=x_column,
            values=y_column,
            title=f"{y_column}-এর ভাগ অনুযায়ী {x_column}",
        )

    chart.update_layout(
        template=chart_template,
        paper_bgcolor=background_color,
        plot_bgcolor=background_color,
        font_color=text_color,
    )
    st.plotly_chart(chart, width="stretch")
elif categorical_columns:
    count_column = st.selectbox("গণনার জন্য কলাম", options=categorical_columns)
    count_data = (
        filtered_df[count_column]
        .fillna("ফাঁকা")
        .astype(str)
        .value_counts()
        .head(30)
        .rename_axis(count_column)
        .reset_index(name="সারি সংখ্যা")
    )
    count_chart = px.bar(
        count_data,
        x=count_column,
        y="সারি সংখ্যা",
        title=f"{count_column} অনুযায়ী সারি সংখ্যা",
    )
    count_chart.update_layout(
        template=chart_template,
        paper_bgcolor=background_color,
        plot_bgcolor=background_color,
        font_color=text_color,
    )
    st.plotly_chart(count_chart, width="stretch")
else:
    st.info("চার্ট তৈরির জন্য একটি সংখ্যাসূচক কলাম এবং একটি গ্রুপ/তারিখ কলাম প্রয়োজন।")

st.subheader("🧾 সাজানো ডেটা")
st.caption(
    f"ফিল্টার করা {len(filtered_df):,}টি সারি দেখানো হচ্ছে। "
    "কলামের শিরোনামে ক্লিক করে টেবিল সাজাতে পারবেন।"
)
st.dataframe(
    filtered_df,
    width="stretch",
    hide_index=True,
    height=420,
)

st.download_button(
    "⬇️ ফিল্টার করা ডেটা CSV হিসেবে ডাউনলোড",
    data=filtered_df.to_csv(index=False).encode("utf-8-sig"),
    file_name="shopmetrics_filtered_data.csv",
    mime="text/csv",
)
