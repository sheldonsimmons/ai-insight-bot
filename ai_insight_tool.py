import streamlit as st
import pandas as pd
import openai
import docx

# ✅ Initialize OpenAI client
client = openai.OpenAI(
    api_key=st.secrets["openai_api_key"],
    project="proj_f0kWLTOK35ZQ5JKBmb2Y5Su9"
)

# ✅ Session state setup
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data = ""

# ✅ UI: Title and Info
st.title("📊 Spreadsheet + Document Insight Bot")
st.markdown("Upload an **Excel (.xlsx)** or **Word (.docx)** file and ask natural language questions about its content.\nYour data is processed **in memory only** and never stored or transmitted.")

# ✅ File uploader
uploaded_file = st.file_uploader("Upload your file (.xlsx or .docx)", type=["xlsx", "docx"])

# Helper functions
def is_excel(filename):
    return filename.lower().endswith(".xlsx")

def is_word(filename):
    return filename.lower().endswith(".docx")

# ✅ File handling
if uploaded_file:
    try:
        file_name = uploaded_file.name
        if is_excel(file_name):
            df = pd.read_excel(uploaded_file)
            st.success("✅ Excel file uploaded successfully.")
            st.markdown("#### Preview (First 100 Rows)")
            st.dataframe(df.head(100))
            st.session_state.uploaded_data = df.head(100).to_string(index=False)

        elif is_word(file_name):
            doc = docx.Document(uploaded_file)
            full_text = "\n".join([para.text for para in doc.paragraphs])
            st.success("✅ Word document uploaded successfully.")
            st.text_area("📄 Document Preview", value=full_text[:1000], height=200)
            st.session_state.uploaded_data = full_text[:3000]

    except Exception as e:
        st.error(f"❌ Error reading file:\n\n{e}")

# ✅ Question interaction
if st.session_state.uploaded_data:
    question = st.text_input("Ask a question about the content:")
    if question:
        with st.spinner("💡 Thinking..."):
            messages = [{"role": "system", "content": "You are a helpful business analyst. Provide concise and relevant answers."}]
            if not st.session_state.chat_history:
                messages.append({"role": "user", "content": f"Here is the content to reference for all upcoming questions:\n{st.session_state.uploaded_data}"})
            messages.extend(st.session_state.chat_history)
            messages.append({"role": "user", "content": question})

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=messages,
                temperature=0.3,
                max_tokens=700
            )
            answer = response.choices[0].message.content
            st.session_state.chat_history.extend([
                {"role": "user", "content": question},
                {"role": "assistant", "content": answer}
            ])
            st.markdown(f"**💬 AI Response:** {answer}")

# ✅ Summary button
if st.session_state.uploaded_data:
    if st.button("🧠 Summarize This File with AI"):
        with st.spinner("Summarizing your data..."):
            summary_prompt = f"""
You're an AI business analyst. Give a short summary of this customer data.
Highlight:
- Key themes or trends in the notes
- Average opportunity score and spend
- Potential next steps based on what you see

Data:
{st.session_state.uploaded_data}
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
