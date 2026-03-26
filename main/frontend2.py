import streamlit as st
import requests
import json
import re
import uuid
from datetime import datetime

# --- CONFIG ---
ROUTER_URL = "http://127.0.0.1:8000/router"
AUTH_TOKEN = "Bearer hackathon-test-token"

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="AI Multi-Agent Studio",
    page_icon="🚀",
    layout="wide"
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
.file-card {
    padding: 14px;
    border-radius: 14px;
    background: #f8f9fa;
    border: 1px solid #eee;
    text-align: center;
    transition: 0.3s;
}
.file-card:hover {
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}
.small-text {
    font-size: 12px;
    color: gray;
}
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.title("🚀 AI Multi-Agent Studio")
st.caption("Generate documents, presentations & spreadsheets intelligently")

# --- SESSION ---
if "session_id" not in st.session_state:
    st.session_state.session_id = f"user_{uuid.uuid4().hex[:6]}"

if "messages" not in st.session_state:
    st.session_state.messages = []

if "history_export" not in st.session_state:
    st.session_state.history_export = ""

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Control Panel")

    output_format = st.selectbox(
        "Preferred Output",
        ["Auto", "DOCX", "PPTX", "Excel"]
    )

    st.markdown("### 📂 Upload JSON")
    uploaded_json = st.file_uploader("Upload structured input", type="json")

    st.markdown("---")

    if st.button("🧹 Clear Chat"):
        st.session_state.messages = []
        st.rerun()

    if st.button("📥 Export Chat"):
        st.download_button(
            "Download Chat",
            st.session_state.history_export,
            file_name="chat.txt"
        )

    st.markdown("---")
    st.write("Session ID")
    st.code(st.session_state.session_id)

# --- FILE RENDER ---


def render_files(text):
    urls = re.findall(
        r'(http://localhost:\d+/outputs/\S+\.(?:xlsx|pptx|docx))', text)

    if not urls:
        return

    st.markdown("### 📂 Generated Files")
    cols = st.columns(len(urls))

    for i, url in enumerate(urls):
        filename = url.split("/")[-1]
        ext = filename.split(".")[-1].upper()

        with cols[i]:
            st.markdown(f"""
            <div class="file-card">
                <b>{filename}</b><br>
                <span class="small-text">{ext} File</span>
            </div>
            """, unsafe_allow_html=True)

            st.link_button("⬇️ Download", url, use_container_width=True)

# --- AGENT STATUS ---


def show_pipeline():
    with st.expander("⚡ Agent Execution Pipeline", expanded=True):
        st.write("🧠 Router → selecting best agent...")
        st.write("📄 DOCX/PPTX/Excel Agent → generating file...")
        st.write("📦 Packaging output...")
        st.success("✅ Done")

# --- RESPONSE ---


def render_assistant(text):
    st.markdown(text)
    render_files(text)


# --- CHAT HISTORY ---
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            render_assistant(msg["content"])
        else:
            st.markdown(msg["content"])

# --- INPUT PRIORITY: JSON UPLOAD ---
if uploaded_json:
    prompt = uploaded_json.read().decode("utf-8")
else:
    prompt = st.chat_input("Try: Create a DOCX report on AI trends...")

# --- HANDLE INPUT ---
if prompt:

    st.session_state.messages.append({"role": "user", "content": prompt})

    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):

        placeholder = st.empty()
        full_response = ""

        show_pipeline()

        try:
            payload = {
                "session_id": st.session_state.session_id,
                "query": prompt,
                "preferred_format": output_format
            }

            headers = {"Authorization": AUTH_TOKEN}

            with requests.post(
                ROUTER_URL,
                data=payload,
                headers=headers,
                stream=True,
                timeout=120
            ) as r:

                if r.status_code == 200:
                    for line in r.iter_lines():
                        if line:
                            chunk = line.decode("utf-8")

                            try:
                                data = json.loads(chunk)
                                msg = data.get("response") or data.get(
                                    "result") or str(data)
                            except:
                                msg = chunk

                            full_response += msg
                            placeholder.markdown(full_response + "▌")

                    placeholder.markdown(full_response)

                    render_files(full_response)

                    # Save chat
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": full_response
                    })

                    # Export log
                    st.session_state.history_export += f"\nUser: {prompt}\nAI: {full_response}\n"

                else:
                    st.error(f"❌ Error {r.status_code}: {r.text}")

        except Exception as e:
            st.error(f"🚨 Connection Error: {e}")
