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
st.title("\ud83d\udcca Spreadsheet + Document Insight Bot")
st.markdown("""
Upload an **Excel (.xlsx)** or **Word (.docx)** file and ask natural language questions about its content.  
Your data is processed **in memory only** and never stored or transmitted.
""")

# ✅ Initialize session state for chat history and content
debug = False
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "content_for_gpt" not in st.session_state:
    st.session_state.content_for_gpt = ""

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
            st.success("\u2705 Excel file uploaded successfully.")
            st.markdown("#### Preview (First 100 Rows)")
            st.dataframe(df.head(100))

            max_rows = 100
            if len(df) > max_rows:
                st.warning(f"Only the first {max_rows} rows will be analyzed.")
                content = df.head(max_rows).to_string(index=False)
            else:
                content = df.to_string(index=False)

            st.session_state.content_for_gpt = content

        elif is_word(file_name):
            doc = docx.Document(uploaded_file)
            full_text = "\n".join([para.text for para in doc.paragraphs])
            st.success("\u2705 Word document uploaded successfully.")
            st.text_area("\ud83d\udcc4 Document Preview", value=full_text[:1000], height=200)
            st.session_state.content_for_gpt = full_text[:3000]

        else:
            st.warning("Unsupported file type.")
            st.session_state.content_for_gpt = ""

    except Exception as e:
        st.error(f"\u274c Error reading file:\n\n{e}")

# ✅ Ask a question about the content
if st.session_state.content_for_gpt:
    question = st.text_input("Ask a question about the content:")
    if question:
        st.session_state.chat_history.append({"role": "user", "content": question})
        with st.spinner("\ud83d\udca1 Thinking..."):
            messages = [
                {"role": "system", "content": "You're an expert data and business assistant. Be helpful, concise, and insightful."},
                {"role": "user", "content": f"Here is the content to analyze:\n{st.session_state.content_for_gpt}"},
                *st.session_state.chat_history
            ]
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=messages,
                temperature=0.2,
                max_tokens=700
            )
            answer = response.choices[0].message.content
            st.session_state.chat_history.append({"role": "assistant", "content": answer})
            st.markdown(f"**\ud83d\udcac AI Response:** {answer}")

# ✅ AI Summary toggle
if st.session_state.content_for_gpt:
    if st.button("\ud83e\uddd1\u200d\ud83e\uddec Summarize This File with AI"):
        with st.spinner("Summarizing your data..."):
            summary_prompt = f"""
You're an AI business analyst. Give a short summary of this customer data.
Highlight:
- Key themes or trends in the notes
- Average opportunity score and spend
- Potential next steps based on what you see

Data:
{st.session_state.content_for_gpt}
            """
            summary_response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": summary_prompt}],
                temperature=0.3,
                max_tokens=500
            )
            summary = summary_response.choices[0].message.content
            st.markdown("### \ud83d\udd0d AI Summary")
            st.info(summary)

# ✅ Optional: Display full chat log (for debugging)
if debug and st.session_state.chat_history:
    st.write("### Debug: Chat History")
    for i, msg in enumerate(st.session_state.chat_history):
        st.write(f"{i+1}. {msg['role']}: {msg['content']}")
