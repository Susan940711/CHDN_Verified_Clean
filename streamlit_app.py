import tempfile
from pathlib import Path

import streamlit as st

from chdn_clean import run_clean

st.set_page_config(page_title="CHDN Clean Step", layout="wide")
st.title("🧹 CHDN Cleaning Step")

st.info("This app runs only the cleaning stage and writes a cleaned workbook for the next report step.")

uploaded_file = st.file_uploader("Upload the source EPI workbook", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    output_path = Path(__file__).resolve().with_name("cleaned_output.xlsx")
    if output_path.exists():
        output_path.unlink()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        temp_input = Path(tmp_file.name)

    with st.spinner("Running the cleaning step..."):
        status = run_clean(temp_input, output_path)

    if status == 0 and output_path.exists():
        st.success("Cleaning step completed successfully.")
        with open(output_path, "rb") as fh:
            data = fh.read()
        st.download_button(
            label="Download cleaned workbook",
            data=data,
            file_name="CHDN_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("The cleaning step did not finish successfully.")
else:
    st.warning("Please upload a workbook to begin.")
