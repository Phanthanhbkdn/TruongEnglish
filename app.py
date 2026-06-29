import html
import json
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

APP_DIR = Path(__file__).parent
DATA_PATH = APP_DIR / "data" / "English.xlsx"

st.set_page_config(
    page_title="English Teaching Dashboard",
    page_icon="📘",
    layout="wide",
)

# -----------------------------
# Data helpers
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data():
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Không tìm thấy file dữ liệu: {DATA_PATH}")
    sheets = pd.read_excel(DATA_PATH, sheet_name=None)
    for name, df in sheets.items():
        sheets[name] = df.fillna("")
    return sheets


def save_sheets(sheets: dict):
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(DATA_PATH, engine="openpyxl", mode="w") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    st.cache_data.clear()


def normalise_word(text: str) -> str:
    return " ".join(str(text).lower().strip().split())


def split_vocab(vocab_text: str):
    if not vocab_text:
        return []
    parts = []
    for item in str(vocab_text).replace(";", ",").split(","):
        item = item.strip()
        if item:
            parts.append(item)
    return parts


def df_date_series(series):
    return pd.to_datetime(series, errors="coerce").dt.date

# -----------------------------
# Audio helpers: browser SpeechSynthesis
# -----------------------------
def speak_button(label: str, text_to_speak: str, key: str = "", compact: bool = False):
    safe_label = html.escape(label)
    js_text = json.dumps(str(text_to_speak))
    button_class = "speak-btn compact" if compact else "speak-btn"
    component_key = f"speaker_{key}_{abs(hash(str(text_to_speak) + str(key))) % 10_000_000}"
    components.html(
        f"""
        <style>
        .speak-btn {{
            border: 1px solid #2E75B6;
            background: #EAF3F8;
            color: #1F4E78;
            border-radius: 8px;
            padding: 6px 10px;
            cursor: pointer;
            font-size: 14px;
            margin: 2px 0;
        }}
        .speak-btn:hover {{ background: #D9EAF7; }}
        .compact {{ padding: 3px 7px; font-size: 13px; }}
        </style>
        <button class="{button_class}" onclick='speakBritishMale_{component_key}({js_text})'>🔊 {safe_label}</button>
        <script>
        function pickVoice_{component_key}() {{
            const voices = window.speechSynthesis.getVoices();
            const gbVoices = voices.filter(v => v.lang && v.lang.toLowerCase().startsWith('en-gb'));
            const maleHints = ['daniel', 'george', 'arthur', 'oliver', 'male'];
            const maleVoice = gbVoices.find(v => maleHints.some(h => v.name.toLowerCase().includes(h)));
            return maleVoice || gbVoices[0] || voices.find(v => v.lang && v.lang.toLowerCase().startsWith('en')) || null;
        }}
        function speakBritishMale_{component_key}(text) {{
            window.speechSynthesis.cancel();
            const utter = new SpeechSynthesisUtterance(text);
            utter.lang = 'en-GB';
            utter.rate = 0.82;
            utter.pitch = 0.85;
            const voice = pickVoice_{component_key}();
            if (voice) utter.voice = voice;
            window.speechSynthesis.speak(utter);
        }}
        if (speechSynthesis.onvoiceschanged !== undefined) {{
            speechSynthesis.onvoiceschanged = pickVoice_{component_key};
        }}
        </script>
        """,
        height=44 if not compact else 34,
    )


def speak_text_inline(text: str, key: str):
    speak_button("UK", text, key=key, compact=True)

# -----------------------------
# UI helpers
# -----------------------------
def card(title, body, accent="#1F4E78"):
    st.markdown(
        f"""
        <div style="border:1px solid #ddd; border-left:6px solid {accent}; border-radius:10px; padding:14px 16px; background:#fff; margin-bottom:10px;">
            <div style="font-weight:700; color:{accent}; font-size:18px; margin-bottom:6px;">{html.escape(str(title))}</div>
            <div style="font-size:15px; line-height:1.5; color:#222; white-space:pre-wrap;">{html.escape(str(body))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_topic_tab(content_df: pd.DataFrame, topic_id: str, topic_name: str):
    st.subheader(f"{topic_id}. {topic_name}")
    topic_df = content_df[content_df["Topic ID"].astype(str) == topic_id].copy()
    if topic_df.empty:
        st.info("Chưa có dữ liệu cho topic này.")
        return

    for definition_id, def_df in topic_df.groupby("Definition ID", sort=False):
        definition = def_df["Definition"].iloc[0]
        st.markdown(
            f"""
            <div style="background:#FFF2CC; border:1px solid #E2C766; border-radius:10px; padding:12px 14px; margin-top:12px;">
                <span style="font-weight:800; color:#7F6000;">{html.escape(str(definition_id))}</span>
                <span style="font-weight:800; color:#7F6000; margin-left:8px;">{html.escape(str(definition))}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        for _, row in def_df.iterrows():
            detail_id = row.get("Detail ID", "")
            detail = row.get("Detail", "")
            example_id = row.get("Example ID", "")
            example = row.get("Example", "")
            audio_text = row.get("Pronunciation Text", example)

            col1, col2 = st.columns([0.84, 0.16])
            with col1:
                st.markdown(
                    f"""
                    <div style="padding:10px 12px; border-bottom:1px dashed #ddd;">
                        <div style="font-weight:700; color:#444;">{html.escape(str(detail_id))}</div>
                        <div style="margin:4px 0 8px 0; line-height:1.45;">{html.escape(str(detail))}</div>
                        <div style="font-style:italic; color:#008000;">{html.escape(str(example_id))} — {html.escape(str(example))}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with col2:
                speak_text_inline(audio_text, key=f"{topic_id}_{detail_id}_{example_id}")


def build_message(row: pd.Series):
    lesson_date = row.get("Lesson Date", "")
    if not isinstance(lesson_date, str):
        try:
            lesson_date = pd.to_datetime(lesson_date).strftime("%d/%m/%Y")
        except Exception:
            lesson_date = str(lesson_date)
    sentence = row.get("Sentence", "")
    vocab_list = split_vocab(row.get("Vocabulary", ""))
    grammar = row.get("Grammar Point", "")
    pron = row.get("Pronunciation Note", "")
    exercise = row.get("Exercise", "")
    topic = row.get("Topic", "")

    vocab_text = "\n".join([f"{i+1}. {w}" for i, w in enumerate(vocab_list)]) or "1. ..."
    return f"""English sentence today – {lesson_date}

Topic: {topic}

Today’s sentence:
{sentence}

Vocabulary:
{vocab_text}

Grammar point:
{grammar}

Pronunciation note:
{pron}

Practice:
{exercise}"""

# -----------------------------
# Main app
# -----------------------------
st.title("📘 English Teaching Dashboard")
st.caption("Nhập bài học, quản lý roadmap, tạo message gửi bạn học và phát âm Anh-Anh bằng trình duyệt.")

sheets = load_data()
content_df = sheets.get("Content", pd.DataFrame())
daily_df = sheets.get("Daily_Lessons", pd.DataFrame())
vocab_df = sheets.get("Vocabulary_Log", pd.DataFrame())

required_content_cols = ["Topic ID", "Topic", "Definition ID", "Definition", "Detail ID", "Detail", "Example ID", "Example", "Pronunciation Text", "Voice", "Note"]
for col in required_content_cols:
    if col not in content_df.columns:
        content_df[col] = ""

main_tabs = ["🏠 Dashboard"]
# Build topic tabs from Content data: A, B, C...
topic_order = content_df[["Topic ID", "Topic"]].drop_duplicates().values.tolist()
for tid, topic in topic_order:
    main_tabs.append(f"{tid}. {topic}")
main_tabs += ["📝 Nhập bài học", "🔁 Kiểm tra trùng từ", "✉️ Message gửi bạn học", "➕ Thêm nội dung Topic"]

tabs = st.tabs(main_tabs)

with tabs[0]:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Topic", content_df["Topic ID"].nunique() if not content_df.empty else 0)
    c2.metric("Definition", content_df["Definition ID"].nunique() if not content_df.empty else 0)
    c3.metric("Example", content_df["Example ID"].nunique() if not content_df.empty else 0)
    c4.metric("Daily lessons", len(daily_df))

    st.markdown("### Roadmap hiện tại")
    if not content_df.empty:
        roadmap = content_df.groupby(["Topic ID", "Topic"], as_index=False).agg(
            Definitions=("Definition ID", "nunique"),
            Details=("Detail ID", "nunique"),
            Examples=("Example ID", "nunique"),
        )
        st.dataframe(roadmap, use_container_width=True, hide_index=True)

    st.markdown("### Câu gần nhất")
    if not daily_df.empty:
        last = daily_df.tail(1).iloc[0]
        card(last.get("Topic", ""), last.get("Sentence", ""), accent="#2E75B6")
        speak_button("Đọc câu gần nhất - UK", last.get("Sentence", ""), key="latest_sentence")
    else:
        st.info("Chưa có bài học hằng ngày.")

# Topic tabs
for i, (tid, topic) in enumerate(topic_order, start=1):
    with tabs[i]:
        render_topic_tab(content_df, str(tid), str(topic))

# Input daily lesson tab
with tabs[len(topic_order)+1]:
    st.subheader("📝 Nhập bài học hôm nay")
    with st.form("daily_lesson_form", clear_on_submit=False):
        lesson_date = st.date_input("Ngày học", value=date.today())
        topic_ids = [str(x) for x in content_df["Topic ID"].dropna().unique().tolist()]
        topic_id = st.selectbox("Topic ID", options=topic_ids or ["A"])
        default_topic = content_df[content_df["Topic ID"].astype(str) == topic_id]["Topic"].iloc[0] if not content_df[content_df["Topic ID"].astype(str) == topic_id].empty else ""
        topic = st.text_input("Topic", value=default_topic)
        sentence = st.text_area("Câu mẫu hôm nay", height=90, placeholder="In our toolbox meeting today, ...")
        vocabulary = st.text_area("Từ vựng chính, cách nhau bằng dấu phẩy", height=70, placeholder="toolbox meeting, project manager, announced")
        grammar = st.text_input("Grammar point", placeholder="must + verb / passive voice / present simple ...")
        pronunciation_note = st.text_input("Ghi chú phát âm", placeholder="31st = thirty-first")
        exercise = st.text_area("Bài tập giao bạn học", height=70, placeholder="Make 2 similar sentences using ...")
        status = st.selectbox("Status", ["Draft", "Sent", "Reviewed"])
        submitted = st.form_submit_button("💾 Lưu bài học")

    if submitted:
        sheets = load_data()
        daily_df = sheets.get("Daily_Lessons", pd.DataFrame())
        vocab_df = sheets.get("Vocabulary_Log", pd.DataFrame())
        new_row = {
            "Lesson Date": lesson_date,
            "Topic ID": topic_id,
            "Topic": topic,
            "Sentence": sentence,
            "Vocabulary": vocabulary,
            "Grammar Point": grammar,
            "Pronunciation Note": pronunciation_note,
            "Exercise": exercise,
            "Status": status,
            "Created At": datetime.now(),
        }
        daily_df = pd.concat([daily_df, pd.DataFrame([new_row])], ignore_index=True)

        new_vocab_rows = []
        for w in split_vocab(vocabulary):
            new_vocab_rows.append({
                "Lesson Date": lesson_date,
                "Word/Phrase": w,
                "Meaning VI": "",
                "Topic": topic,
                "Source Sentence": sentence,
                "Duplicate Check Key": normalise_word(w),
            })
        if new_vocab_rows:
            vocab_df = pd.concat([vocab_df, pd.DataFrame(new_vocab_rows)], ignore_index=True)

        sheets["Daily_Lessons"] = daily_df
        sheets["Vocabulary_Log"] = vocab_df
        save_sheets(sheets)
        st.success("Đã lưu bài học và vocabulary log.")
        st.rerun()

# Duplicate check tab
with tabs[len(topic_order)+2]:
    st.subheader("🔁 Kiểm tra trùng từ vựng với ngày trước")
    if vocab_df.empty:
        st.info("Chưa có dữ liệu Vocabulary_Log.")
    else:
        temp = vocab_df.copy()
        temp["_date"] = df_date_series(temp["Lesson Date"])
        available_dates = sorted([d for d in temp["_date"].dropna().unique()])
        selected_date = st.date_input("Chọn ngày cần kiểm tra", value=available_dates[-1] if available_dates else date.today(), key="dup_date")
        previous_dates = [d for d in available_dates if d < selected_date]
        if not previous_dates:
            st.warning("Chưa có ngày trước để so sánh.")
        else:
            prev_date = max(previous_dates)
            today_vocab = temp[temp["_date"] == selected_date].copy()
            prev_vocab = temp[temp["_date"] == prev_date].copy()
            today_vocab["_key"] = today_vocab["Duplicate Check Key"].apply(normalise_word)
            prev_vocab["_key"] = prev_vocab["Duplicate Check Key"].apply(normalise_word)
            duplicates = today_vocab[today_vocab["_key"].isin(set(prev_vocab["_key"]))].copy()
            st.write(f"So sánh ngày **{selected_date.strftime('%d/%m/%Y')}** với ngày trước gần nhất **{prev_date.strftime('%d/%m/%Y')}**")
            if duplicates.empty:
                st.success("Không có từ/cụm từ trùng với ngày trước.")
            else:
                st.error("Có từ/cụm từ trùng. Nên đổi trước khi gửi bạn học.")
                show_cols = ["Word/Phrase", "Topic", "Source Sentence"]
                st.dataframe(duplicates[show_cols], use_container_width=True, hide_index=True)

                st.markdown("#### Gợi ý đổi nhanh")
                suggestions = {
                    "project manager": "site manager / construction manager",
                    "announced": "informed / stated / confirmed",
                    "toolbox meeting": "site briefing / safety briefing",
                    "structural package": "structural works / concrete frame package",
                }
                for w in duplicates["_key"].tolist():
                    st.write(f"- **{w}** → {suggestions.get(w, 'đổi sang từ đồng nghĩa hoặc cụm khác cùng chủ đề')}")

# Message tab
with tabs[len(topic_order)+3]:
    st.subheader("✉️ Tạo message gửi bạn học")
    if daily_df.empty:
        st.info("Chưa có bài học để tạo message.")
    else:
        daily_temp = daily_df.copy()
        daily_temp["_label"] = daily_temp.apply(lambda r: f"{pd.to_datetime(r['Lesson Date'], errors='coerce').strftime('%d/%m/%Y') if str(r.get('Lesson Date','')) else ''} - {r.get('Topic','')}", axis=1)
        selected_idx = st.selectbox("Chọn bài học", options=list(daily_temp.index), format_func=lambda i: daily_temp.loc[i, "_label"])
        row = daily_temp.loc[selected_idx]
        msg = build_message(row)
        st.text_area("Message", value=msg, height=330)
        c1, c2 = st.columns([0.25, 0.75])
        with c1:
            speak_button("Đọc câu mẫu - UK", row.get("Sentence", ""), key="message_sentence")
        with c2:
            st.caption("Copy nội dung message phía trên để gửi Zalo/Teams. Nút phát âm dùng giọng en-GB của trình duyệt, ưu tiên giọng nam nếu máy có hỗ trợ.")

# Add content topic tab
with tabs[len(topic_order)+4]:
    st.subheader("➕ Thêm nội dung vào Topic/Roadmap")
    with st.form("content_form", clear_on_submit=True):
        col_a, col_b = st.columns(2)
        with col_a:
            topic_id_new = st.text_input("Topic ID", value="C")
            topic_new = st.text_input("Topic", placeholder="Ví dụ: Bị động")
            definition_id = st.text_input("Definition ID", placeholder="C.1")
            definition = st.text_input("Definition", placeholder="Câu bị động là gì?")
        with col_b:
            detail_id = st.text_input("Detail ID", placeholder="C.1.1")
            example_id = st.text_input("Example ID", placeholder="C.1.1.a")
            voice = st.text_input("Voice", value="en-GB male")
            note = st.text_input("Note", value="")
        detail = st.text_area("Detail", height=90)
        example = st.text_area("Example", height=70)
        pronunciation_text = st.text_area("Pronunciation Text", value=example, height=70)
        add_content = st.form_submit_button("💾 Thêm vào Content")

    if add_content:
        sheets = load_data()
        content_df = sheets.get("Content", pd.DataFrame())
        new_content = {
            "Topic ID": topic_id_new.strip(),
            "Topic": topic_new.strip(),
            "Definition ID": definition_id.strip(),
            "Definition": definition.strip(),
            "Detail ID": detail_id.strip(),
            "Detail": detail.strip(),
            "Example ID": example_id.strip(),
            "Example": example.strip(),
            "Pronunciation Text": pronunciation_text.strip() or example.strip(),
            "Voice": voice.strip(),
            "Note": note.strip(),
        }
        content_df = pd.concat([content_df, pd.DataFrame([new_content])], ignore_index=True)
        sheets["Content"] = content_df
        save_sheets(sheets)
        st.success("Đã thêm nội dung. App sẽ tự tạo thêm tab Topic nếu Topic ID mới.")
        st.rerun()

st.markdown("---")
st.caption("Dữ liệu lưu tại data/English.xlsx. Có thể mở Excel để backup, nhưng nên nhập qua Streamlit để giữ cấu trúc.")
