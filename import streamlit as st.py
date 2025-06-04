import streamlit as st
import pandas as pd
import openai
import io
import docx

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

# ✅ File uploader for .xlsx and .docx
uploaded_file = st.file_uploader("Upload your file (.xlsx or .docx)", type=["xlsx", "docx"])

def is_excel(filename):
    return filename.lower().endswith(".xlsx")

def is_word(filename):
    return filename.lower().endswith(".docx")

if uploaded_file:
    try:
        file_name = uploaded_file.name

        if is_excel(file_name):
            df = pd.read_excel(uploaded_file)
            st.success("✅ Excel file uploaded successfully.")
            st.dataframe(df.head(51))  # Preview first 10 rows

            # 👇 Use full DataFrame content (limit for safety if large)
            max_rows = 75
            if len(df) > max_rows:
                st.warning(f"Only the first {max_rows} rows will be analyzed.")
                content_for_gpt = df.head(max_rows).to_string(index=False)
            else:
                content_for_gpt = df.to_string(index=False)

        elif is_word(file_name):
            doc = docx.Document(uploaded_file)
            full_text = "\n".join([para.text for para in doc.paragraphs])
            st.success("✅ Word document uploaded successfully.")
            st.text_area("📄 Document Preview", value=full_text[:1000], height=200)

            # 👇 Limit the document text for GPT (to avoid token overload)
            content_for_gpt = full_text[:3000]

        else:
            st.warning("Unsupported file type.")
            content_for_gpt = None

        # ✅ Ask AI a question
        if content_for_gpt:
            question = st.text_input("Ask a question about the content:")
            if question:
                with st.spinner("💡 Thinking..."):
                    prompt = f"""
You are a helpful data and document analyst. Use the following content to answer the user's question.

Content:
{content_for_gpt}

Question: {question}
Answer:
"""
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.2,
                        max_tokens=500
                    )
                    answer = response.choices[0].message.content
                    st.markdown(f"**💬 AI Response:** {answer}")

    except Exception as e:
        st.error(f"❌ Error reading file:\n\n{e}")
