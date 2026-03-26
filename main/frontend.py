import streamlit as st
import requests
import json  # <-- Crucial import for the fix
import re
import uuid

# --- CONFIGURATION ---
ROUTER_URL = "http://localhost:8000/router"

# --- UI SETUP ---
st.set_page_config(page_title="Shipathon Multi-Agent Gateway",
                   page_icon="🚀", layout="centered")
st.title("🚀 Multi-Agent Gateway")
st.markdown(
    "Ask for an Excel file, PPTX, or DOCX, and the AI Router will handle the rest!")

# --- SESSION STATE (Chat Memory) ---
if "session_id" not in st.session_state:
    st.session_state.session_id = f"session_{uuid.uuid4().hex[:8]}"

if "messages" not in st.session_state:
    st.session_state.messages = []

# --- HELPER FUNCTION: RENDER THE DOWNLOAD BUTTON ---


def render_assistant_response(text):
    """Parses the backend response and renders a clean UI button if a file URL is found."""
    # Catch standard URLs for local files
    url_match = re.search(
        r'(http://localhost:\d+/\S+\.(?:xlsx|pptx|docx|csv))', text)

    if url_match:
        file_url = url_match.group(1)
        st.success("✨ Generation Complete!")
        st.link_button("📥 Download Your File", file_url, type="primary")
    else:
        st.markdown(text)


# --- DRAW PREVIOUS CHAT HISTORY ---
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            render_assistant_response(msg["content"])
        else:
            st.markdown(msg["content"])

# --- MAIN CHAT INPUT ---
if prompt := st.chat_input("E.g., Generate an excel file with names Varun, Shantanu, Deepika, Prapti..."):

    # 1. Add user message to history and display it
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 2. Assistant response block
    with st.chat_message("assistant"):
        final_text = ""

        # UI Magic: The collapsible status box
        with st.status("🤖 Routing query to the best agent...", expanded=True) as status:
            st.write("Sending payload to the Orchestrator...")

            try:
                # --- THE STRICT API CALL ---
                payload = {
                    "query": prompt,
                    "session_id": st.session_state.session_id
                }

                # Explicitly forcing JSON and adding the VIP badge
                headers = {
                    "Authorization": "Bearer hackathon_admin_token_123",
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                }

                # THE FIX: We use data=json.dumps() to force standard JSON serialization
                response = requests.post(
                    ROUTER_URL, data=json.dumps(payload), headers=headers)

                if response.status_code == 200:
                    try:
                        data = response.json()
                        final_text = data.get(
                            "response", data.get("result", str(data)))
                    except ValueError:
                        final_text = response.text

                    status.update(label="✅ Task Finished!",
                                  state="complete", expanded=False)
                else:
                    final_text = f"⚠️ Router Error: {response.status_code} - {response.text}"
                    status.update(label="❌ Task Failed",
                                  state="error", expanded=False)

            except requests.exceptions.ConnectionError:
                final_text = "🚨 Connection Error: Is the Router running on port 8000?"
                status.update(label="❌ Connection Failed",
                              state="error", expanded=False)
            except Exception as e:
                final_text = f"🚨 Unexpected Error: {str(e)}"
                status.update(label="❌ Error", state="error", expanded=False)

        # Render the result (either the Download Button or the Error Text)
        render_assistant_response(final_text)

        # Add assistant response to history so it stays on screen
        st.session_state.messages.append(
            {"role": "assistant", "content": final_text})
