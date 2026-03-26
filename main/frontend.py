import streamlit as st
import requests
import json
import re
import uuid

# --- CONFIG ---
# Use the exact URL from your main.py
ROUTER_URL = "http://127.0.0.1:8000/router"
AUTH_TOKEN = "Bearer hackathon-test-token"

st.set_page_config(page_title="Multi-Agent Gateway",
                   page_icon="🏢", layout="wide")
st.title("🏢 Multi-Agent Gateway")

if "session_id" not in st.session_state:
    st.session_state.session_id = f"varun_{uuid.uuid4().hex[:4]}"
if "messages" not in st.session_state:
    st.session_state.messages = []


def render_assistant_response(text):
    # Detects links from any of your 3 agents
    url_pattern = r'(http://localhost:\d+/outputs/\S+\.(?:xlsx|pptx|docx))'
    urls = re.findall(url_pattern, text)
    st.markdown(text)
    if urls:
        st.divider()
        st.subheader("📥 Generated Files")
        cols = st.columns(len(urls))
        for idx, url in enumerate(urls):
            filename = url.split("/")[-1]
            ext = filename.split(".")[-1].upper()
            with cols[idx]:
                st.link_button(
                    f"Download {ext}", url, type="primary", use_container_width=True)


# Render Chat History
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            render_assistant_response(msg["content"])
        else:
            st.markdown(msg["content"])

if prompt := st.chat_input("Enter your mega-prompt..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        full_response_text = ""
        placeholder = st.empty()  # For streaming updates

        try:
            # THE FIX: Use 'data=' instead of 'json=' to send as Form Data
            payload = {
                "session_id": st.session_state.session_id,
                "query": prompt
            }

            headers = {"Authorization": AUTH_TOKEN}

            # We use stream=True because your Router returns a StreamingResponse
            with requests.post(ROUTER_URL, data=payload, headers=headers, stream=True, timeout=120) as r:
                if r.status_code == 200:
                    for line in r.iter_lines():
                        if line:
                            # Decodes the chunk from the stream
                            chunk = line.decode('utf-8')
                            try:
                                # If the chunk is a JSON string from the agent
                                data = json.loads(chunk)
                                # Extract content from common A2A/Orchestrator keys
                                msg_content = data.get("response") or data.get(
                                    "result") or str(data)
                                full_response_text += msg_content
                            except:
                                # If it's just raw text
                                full_response_text += chunk

                            placeholder.markdown(full_response_text + "▌")

                    placeholder.empty()  # Clear the cursor
                    render_assistant_response(full_response_text)
                    st.session_state.messages.append(
                        {"role": "assistant", "content": full_response_text})
                else:
                    st.error(f"Error {r.status_code}: {r.text}")

        except Exception as e:
            st.error(f"Connection Error: {e}")
