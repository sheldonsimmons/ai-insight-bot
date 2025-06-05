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

# Session state to hold chat history
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

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
            st.markdown("#### Preview (First 200 Rows)")
            st.dataframe(df.head(100))

            max_rows = 100
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
            content_for_gpt = full_text[:3000]

        else:
            st.warning("Unsupported file type.")
            content_for_gpt = None

        if content_for_gpt:
            question = st.text_input("Ask a question about the content:")
            if question:
                st.session_state.chat_history.append({"role": "user", "content": question})
                with st.spinner("💡 Thinking..."):
                    messages = [{"role": "system", "content": "You're an expert data and business assistant. Be helpful, concise, and insightful."}]
                    messages.append({"role": "user", "content": f"Here's the content:\n{content_for_gpt}"})
                    messages.extend(st.session_state.chat_history)

                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=messages,
                        temperature=0.2,
                        max_tokens=700
                    )
                    answer = response.choices[0].message.content
                    st.session_state.chat_history.append({"role": "assistant", "content": answer})
                    st.markdown(f"**💬 AI Response:** {answer}")

            # ✅ Summary option
            if st.checkbox("🧠 Summarize This File with AI"):
                with st.spinner("Summarizing your data..."):
                    summary_prompt = f"""
You're an AI business analyst. Give a short summary of this customer data.
Highlight:
- Key themes or trends in the notes
- Average opportunity score and spend
- Potential next steps based on what you see

Data:
{content_for_gpt}
                    """
                    summary_response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "user", "content": summary_prompt}],
                        temperature=0.3,
                        max_tokens=500
                    )
                    summary = summary_response.choices[0].message.content
                    st.markdown("### 🔍 AI Summary")
                    st.info(summary)

    except Exception as e:
        st.error(f"❌ Error reading file:\n\n{e}")
