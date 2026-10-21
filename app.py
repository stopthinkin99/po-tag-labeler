import streamlit as st

from core import load_po_tag_map, process_pdf


st.set_page_config(page_title="PO Tag Labeler", page_icon="🏷️", layout="wide")

st.title("PO Tag Labeler")
st.write(
    "Upload the order Excel file and the shipping-label PDF. The app matches "
    "FedEx REF/PO values and UPS Purchase No values to the Excel Cust Po No, "
    "then prints the matching tag number in a bold circle on the label."
)

left, right = st.columns(2)
with left:
    excel_file = st.file_uploader("1. Upload Excel", type=["xlsx", "xlsm"])
with right:
    pdf_file = st.file_uploader("2. Upload label PDF", type=["pdf"])

settings_col1, settings_col2, settings_col3 = st.columns(3)
with settings_col1:
    position = st.selectbox(
        "Tag position",
        options=["auto-safe-top", "top-right", "top-left", "top-center"],
        index=0,
    )
with settings_col2:
    font_size = st.slider("Number font size", min_value=7, max_value=16, value=10)
with settings_col3:
    remove_old_tag_box = st.checkbox("Remove old large TAG NO box", value=True)

if st.button("Process Labels", type="primary", use_container_width=True):
    if not excel_file or not pdf_file:
        st.error("Please upload both the Excel file and the PDF.")
    else:
        try:
            with st.spinner("Matching FedEx/UPS labels and adding circled tag numbers..."):
                mapping = load_po_tag_map(excel_file.getvalue())
                output_pdf, results = process_pdf(
                    pdf_file.getvalue(),
                    mapping,
                    position=position,
                    font_size=font_size,
                    remove_old_tag_box=remove_old_tag_box,
                )

            matched = sum(r.status.startswith("Matched") for r in results)
            unmatched = len(results) - matched
            ups_ocr = sum(r.source == "UPS Purchase No (OCR)" for r in results)

            metric1, metric2, metric3, metric4 = st.columns(4)
            metric1.metric("PDF pages", len(results))
            metric2.metric("Matched", matched)
            metric3.metric("UPS OCR matches", ups_ocr)
            metric4.metric("Needs review", unmatched)

            if unmatched:
                st.warning(
                    "The output was created, but some pages were not matched. "
                    "Review the table below before using the PDF."
                )
            else:
                st.success("Every label was matched successfully.")

            output_name = pdf_file.name.rsplit(".", 1)[0] + "_with_tags.pdf"
            st.download_button(
                "Download Tagged PDF",
                data=output_pdf,
                file_name=output_name,
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )

            st.subheader("Match report")
            st.dataframe(
                [
                    {
                        "Page": r.page_number,
                        "PO / REF": r.po_number or "",
                        "Tag NO": ", ".join(r.tags),
                        "Read from": r.source,
                        "Status": r.status,
                    }
                    for r in results
                ],
                use_container_width=True,
                hide_index=True,
            )

        except Exception as exc:
            st.exception(exc)
