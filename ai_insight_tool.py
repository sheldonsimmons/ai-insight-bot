import streamlit as st
import pandas as pd
import openai
import io
import docx
from docx import Document
from fpdf import FPDF
import xlsxwriter
import tempfile

# ✅ Initialize OpenAI client
client = openai.OpenAI(
    api_key=st.secrets["openai_api_key"],
    project="proj_f0kWLTOK35ZQ5JKBmb2Y5Su9"
)

# ✅ UI: Title and Info
st.title("📊 Spreadsheet + Document Insight Bot")
st.markdown("""
Upload an **Excel (.xlsx)** or **Word (.docx)** file and ask natural language questions about its content.  
Your data is processed **in memory only** and never stored or transmitted.
""")

# ✅ Initialize session state for chat history and content
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "content_for_gpt" not in st.session_state:
    st.session_state.content_for_gpt = ""
if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""

# ✅ File uploader for .xlsx and .docx
uploaded_file = st.file_uploader("Upload your file (.xlsx or .docx)", type=["xlsx", "docx"])

def is_excel(filename):
    return filename.lower().endswith(".xlsx")

def is_word(filename):
    return filename.lower().endswith(".docx")

MAX_ROWS = 500

if uploaded_file:
    try:
        file_name = uploaded_file.name

        if is_excel(file_name):
            df = pd.read_excel(uploaded_file)
            st.success("✅ Excel file uploaded successfully.")

            if len(df) > MAX_ROWS:
                st.warning(f"This file has {len(df)} rows. For best performance, only the first {MAX_ROWS} rows will be used.")
                df = df.head(MAX_ROWS)

            st.markdown("#### Preview (First Rows Analyzed)")
            st.dataframe(df)
            st.session_state.df = df
            st.session_state.content_for_gpt = df.to_string(index=False)

        elif is_word(file_name):
            doc = docx.Document(uploaded_file)
            full_text = "\n".join([para.text for para in doc.paragraphs])
            st.success("✅ Word document uploaded successfully.")
            st.text_area("📄 Document Preview", value=full_text[:1000], height=200)
            st.session_state.content_for_gpt = full_text[:3000]

        else:
            st.warning("Unsupported file type.")
            st.session_state.content_for_gpt = ""

    except Exception as e:
        st.error(f"❌ Error reading file:\n\n{e}")

# ✅ Ask a question about the content
if st.session_state.content_for_gpt:
    question = st.text_input("Ask a question about the content:")
    if question:
        with st.spinner("💡 Thinking..."):
            messages = [
                {"role": "system", "content": "You're an expert data and business assistant. Be helpful, concise, and insightful."},
                {"role": "user", "content": f"Here is the content to analyze:\n{st.session_state.content_for_gpt}"}
            ]
            for chat in st.session_state.chat_history:
                messages.append(chat)
            messages.append({"role": "user", "content": question})

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=messages,
                temperature=0.2,
                max_tokens=700
            )
            answer = response.choices[0].message.content
            st.session_state.chat_history.append({"role": "user", "content": question})
            st.session_state.chat_history.append({"role": "assistant", "content": answer})
            st.session_state.last_answer = answer
            st.markdown("**💬 AI Response:**")
            st.markdown(answer)

            # ✅ Default output as formatted Excel
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                worksheet_name = "AI Response"
                if 'df' in st.session_state:
                    df_answer = pd.DataFrame([x.split(" - ", 1) for x in answer.split("\n") if " - " in x], columns=["Customer", "Notes"])
                else:
                    df_answer = pd.DataFrame({"AI Response": answer.split("\n")})
                df_answer.to_excel(writer, index=False, sheet_name=worksheet_name)
                workbook = writer.book
                worksheet = writer.sheets[worksheet_name]
                header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1'})
                for col_num, value in enumerate(df_answer.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            excel_buffer.seek(0)
            st.download_button("⬇️ Download as .xlsx (default)", data=excel_buffer, file_name="ai_response.xlsx")

            # ✅ Offer additional download options
            file_format = st.selectbox("Also download as:", ["Text", "Word (.docx)", "PDF"])
            if file_format:
                if file_format == "Text":
                    st.download_button("⬇️ Download as .txt", data=answer, file_name="ai_response.txt")

                elif file_format == "Word (.docx)":
                    doc_buffer = io.BytesIO()
                    docx_file = Document()
                    docx_file.add_paragraph(answer)
                    docx_file.save(doc_buffer)
                    doc_buffer.seek(0)
                    st.download_button("⬇️ Download as .docx", data=doc_buffer, file_name="ai_response.docx")

                elif file_format == "PDF":
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)
                    for line in answer.split("\n"):
                        pdf.multi_cell(0, 10, line)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                        pdf.output(tmp_file.name)
                        tmp_file.seek(0)
                        st.download_button("⬇️ Download as .pdf", data=tmp_file.read(), file_name="ai_response.pdf")

# ✅ AI Summary toggle
if st.session_state.content_for_gpt:
    if st.button("Summarize File with AI"):
        with st.spinner("Summarizing your data..."):
            summary_prompt = f"""
You're an AI business analyst. Give a short summary of this customer data. Focus only on the contents below — do not refer to previous questions. Your job is to:

- Highlight themes and trends in the customer notes
- Mention average opportunity score and spend if possible
- Suggest smart next steps for marketing or sales

Here’s the data:
{st.session_state.content_for_gpt}
            """
            summary_response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You're an expert data and business analyst. Be helpful, concise, and insightful."},
                    {"role": "user", "content": summary_prompt}
                ],
                temperature=0.3,
                max_tokens=700
            )
            summary = summary_response.choices[0].message.content
            st.markdown("### 🔍 AI Summary")
            st.info(summary)

# ✅ Optional: Display full chat log
if st.session_state.chat_history:
    with st.expander("🗂 View Chat History"):
        for i, msg in enumerate(st.session_state.chat_history):
            st.markdown(f"**{msg['role'].capitalize()}:** {msg['content']}")
