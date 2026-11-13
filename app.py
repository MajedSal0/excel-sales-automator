import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Excel Sales Automator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Sales Automator")
st.write("Upload your sales file and generate an automated report.")

uploaded_file = st.file_uploader(
    "Choose an Excel or CSV file",
    type=["xlsx", "csv"]
)

if uploaded_file is not None:


    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)



    # Check for empty file
    if df.empty:
        st.error("The uploaded file contains no data.")
        st.stop()

    # Column Mapping
    st.subheader("Column Mapping")

    if len(df.columns) < 3:
        st.error("Your file must contain at least 3 columns.")
        st.stop()

    product_column = st.selectbox(
        "Product Column",
        df.columns,
        index=0
    )

    quantity_column = st.selectbox(
        "Quantity Column",
        df.columns,
        index=1
    )

    price_column = st.selectbox(
        "Price Column",
        df.columns,
        index=2
    )

    # Validate column selection
    if len({product_column, quantity_column, price_column}) != 3:
        st.error("Please select three different columns.")
        st.stop()

    # Convert selected columns to numeric values
    df[quantity_column] = pd.to_numeric(
        df[quantity_column], errors="coerce"
    )

    df[price_column] = pd.to_numeric(
        df[price_column], errors="coerce"
    )

    # Validate data
    if df[[quantity_column, price_column]].isna().any().any():
        st.error("Quantity and Price must contain valid numbers.")
        st.stop()

    # Calculate sales
    df["Total"] = df[quantity_column] * df[price_column]

    total_sales = df["Total"].sum()
    units_sold = df[quantity_column].sum()

    # Find best-selling product by total quantity
    product_quantities = df.groupby(product_column)[quantity_column].sum()
    best_seller = product_quantities.idxmax()

    st.subheader("Sales Dashboard")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            "Total Revenue",
            f"{total_sales:,.2f} SAR"
        )

    with col2:
        st.metric(
            "Units Sold",
            f"{units_sold:,}"
        )

    with col3:
        st.metric(
            "Best Seller",
            best_seller
        )

    st.subheader("Revenue by Product")
    revenue_by_product = df.groupby(product_column)["Total"].sum()
    st.bar_chart(revenue_by_product)
    st.subheader("Sales Data")
    st.dataframe(df, use_container_width=True)


    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        summary = pd.DataFrame({
            "Metric": [
                "Total Revenue",
                "Units Sold",
                "Best Seller"
            ],
            "Value": [
                total_sales,
                units_sold,
                best_seller
            ]
        })

        summary.to_excel(
            writer,
            index=False,
            sheet_name="Sales Summary"
        )

        df.to_excel(
            writer,
            index=False,
            sheet_name="Sales Data"
        )
                # Format Excel worksheets
        for sheet in writer.book.worksheets:

            sheet.freeze_panes = "A2"

            # Style header row
            for cell in sheet[1]:
                cell.font = Font(
                    bold=True,
                    color="FFFFFF"
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="17365D"
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )

            # Adjust column widths
            for column in sheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column
                )

                column_letter = get_column_letter(column[0].column)

                sheet.column_dimensions[column_letter].width = min(
                    max(max_length + 5, 15),
                    45
                )


    st.download_button(
        label="📥 Download Report",
        data=output.getvalue(),
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )