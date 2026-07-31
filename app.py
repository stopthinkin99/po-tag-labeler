import streamlit as st

from core import load_po_tag_map, process_pdf


st.set_page_config(page_title="PO Tag Labeler", page_icon="🏷️", layout="wide")

st.title("PO Tag Labeler")
st.write(
    "Upload the order Excel file and the shipping-label PDF. The app matches "
    "each label's REF/PO number to the Excel Cust Po No and prints the matching "
    "Tag NO on the label."
)

left, right = st.columns(2)
with left:
    excel_file = st.file_uploader("1. Upload Excel", type=["xlsx", "xlsm"])
with right:
    pdf_file = st.file_uploader("2. Upload label PDF", type=["pdf"])

settings_col1, settings_col2 = st.columns(2)
with settings_col1:
    position = st.selectbox(
        "Tag position",
        options=["top-right", "top-left", "top-center"],
        index=0,
    )
with settings_col2:
    font_size = st.slider("Tag font size", min_value=10, max_value=28, value=18)

if st.button("Process Labels", type="primary", use_container_width=True):
    if not excel_file or not pdf_file:
        st.error("Please upload both the Excel file and the PDF.")
    else:
        try:
            with st.spinner("Matching labels and adding tag numbers..."):
                mapping = load_po_tag_map(excel_file.getvalue())
                output_pdf, results = process_pdf(
                    pdf_file.getvalue(),
                    mapping,
                    position=position,
                    font_size=font_size,
                )

            matched = sum(r.status == "Matched" for r in results)
            unmatched = len(results) - matched

            metric1, metric2, metric3 = st.columns(3)
            metric1.metric("PDF pages", len(results))
            metric2.metric("Matched", matched)
            metric3.metric("Needs review", unmatched)

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
                        "Status": r.status,
                    }
                    for r in results
                ],
                use_container_width=True,
                hide_index=True,
            )

        except Exception as exc:
            st.exception(exc)
