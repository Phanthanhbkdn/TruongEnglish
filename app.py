import html
import json
from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# ============================================================
# CONFIG
# ============================================================

APP_DIR = Path(__file__).parent
DATA_PATH = APP_DIR / "data" / "English.xlsx"

st.set_page_config(
    page_title="English Teaching Dashboard",
    page_icon="📘",
    layout="wide",
)


# ============================================================
# STYLE
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 42px;
        font-weight: 800;
        margin-bottom: 0.2rem;
        color: #1f2937;
    }

    .sub-title {
        color: #6b7280;
        font-size: 15px;
        margin-bottom: 24px;
    }

    .definition-box {
        background: #fff3cd;
        border: 1px solid #f0c94d;
        color: #3b2f00;
        padding: 12px 16px;
        border-radius: 10px;
        font-weight: 800;
        margin-top: 18px;
        margin-bottom: 12px;
    }

    .detail-id {
        font-weight: 800;
        color: #1f2937;
        margin-top: 8px;
    }

    .detail-text {
        font-size: 16px;
        color: #111827;
        line-height: 1.6;
    }

    .example-text {
        color: #008000;
        font-style: italic;
        font-size: 16px;
        line-height: 1.5;
        margin-top: 8px;
    }

    .translation-text {
        color: #4b5563;
        font-size: 15px;
        margin-top: 2px;
        margin-bottom: 8px;
    }

    .vocab-chip {
        display: inline-block;
        background: #eef6ff;
        border: 1px solid #bfdbfe;
        color: #1e3a8a;
        border-radius: 999px;
        padding: 6px 12px;
        margin: 4px 6px 4px 0;
        font-size: 14px;
    }

    .ipa-text {
        color: #7c3aed;
        font-weight: 700;
    }

    .small-muted {
        color: #6b7280;
        font-size: 13px;
    }

    .section-card {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 12px;
        background: #ffffff;
    }

    hr {
        margin-top: 16px;
        margin-bottom: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:
    """Convert Excel cell to clean string."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Find first matching column, case-insensitive."""
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        key = c.strip().lower()
        if key in lower_map:
            return lower_map[key]
    return None


def normalize_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Normalize many possible Excel formats into one standard structure."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {
        "Topic ID": ["Topic ID", "TopicID", "Topic_Id", "ID Topic"],
        "Topic": ["Topic", "Tên Topic", "Chủ đề"],
        "Definition ID": ["Definition ID", "DefinitionID", "Def ID", "Def_ID", "ID"],
        "Definition": ["Definition", "Định nghĩa", "Noi dung chinh", "Nội dung chính"],
        "Detail ID": ["Detail ID", "DetailID", "Sub ID", "SubID", "sub ID"],
        "Detail": ["Detail", "Chi tiết", "Details"],
        "Example ID": ["Example ID", "ExampleID", "Ex ID", "Example_Id"],
        "Example": ["Example", "Ví dụ", "Example EN", "Sentence", "Câu ví dụ"],
        "Example VI": ["Example VI", "Example Vietnamese", "Translation", "Dịch nghĩa", "Dich nghia", "Vietnamese"],
        "Vocabulary": ["Vocabulary", "Word", "Từ vựng", "Tu vung", "Vocab"],
        "Meaning VI": ["Meaning VI", "Meaning", "Nghĩa", "Nghia", "Vietnamese Meaning"],
        "IPA": ["IPA", "Phonetic", "Phiên âm", "Phien am"],
    }

    out = pd.DataFrame()

    for std_col, candidates in col_map.items():
        found = first_existing_col(df, candidates)
        out[std_col] = df[found] if found else ""

    topic_from_sheet = sheet_name
    if " - " in topic_from_sheet:
        topic_from_sheet = topic_from_sheet.split(" - ", 1)[-1].strip()

    out["Topic"] = out["Topic"].apply(clean_text)
    out.loc[out["Topic"] == "", "Topic"] = topic_from_sheet

    inferred_topic_id = sheet_name.split(".", 1)[0].strip()
    if not inferred_topic_id:
        inferred_topic_id = "A"

    out["Topic ID"] = out["Topic ID"].apply(clean_text)
    out.loc[out["Topic ID"] == "", "Topic ID"] = inferred_topic_id

    for col in out.columns:
        out[col] = out[col].apply(clean_text)

    content_cols = ["Definition", "Detail", "Example", "Vocabulary", "IPA"]
    out = out[out[content_cols].apply(lambda r: any(clean_text(x) for x in r), axis=1)]

    return out.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_data() -> dict[str, pd.DataFrame]:
    if not DATA_PATH.exists():
        st.error(f"Không tìm thấy file dữ liệu: {DATA_PATH}")
        st.info("Hãy kiểm tra GitHub có đúng cấu trúc: data/English.xlsx chưa.")
        st.stop()

    sheets = pd.read_excel(DATA_PATH, sheet_name=None, engine="openpyxl")
    normalized = {}

    for sheet_name, df in sheets.items():
        if sheet_name.strip().lower() in ["index", "dailylessons", "messages", "sent_messages"]:
            continue
        if df.empty:
            continue

        norm = normalize_sheet(df, sheet_name)
        if not norm.empty:
            normalized[sheet_name] = norm

    if not normalized:
        st.warning("File English.xlsx chưa có dữ liệu bài học để hiển thị.")
        st.stop()

    return normalized


def render_speak_button(text: str, key: str, label: str = "🔊 UK", height: int = 42):
    """
    Button phát âm dùng Web Speech API.
    Ưu tiên giọng en-GB male nếu trình duyệt có hỗ trợ.
    """
    text = clean_text(text)
    if not text:
        st.caption("")
        return

    payload = json.dumps(text)

    components.html(
        f"""
        <button id="btn_{key}" onclick="speak_{key}()"
            style="
                border:1px solid #1f77b4;
                border-radius:8px;
                padding:4px 10px;
                background:white;
                color:#0b5394;
                cursor:pointer;
                font-size:14px;
                white-space:nowrap;
            ">
            {label}
        </button>

        <script>
        function speak_{key}() {{
            const text = {payload};
            const utterance = new SpeechSynthesisUtterance(text);
            utterance.lang = "en-GB";
            utterance.rate = 0.82;
            utterance.pitch = 0.95;

            function setVoiceAndSpeak() {{
                const voices = window.speechSynthesis.getVoices();

                const maleGbVoice = voices.find(v =>
                    v.lang && v.lang.toLowerCase().startsWith("en-gb") &&
                    v.name && v.name.toLowerCase().includes("male")
                );

                const gbVoice = voices.find(v =>
                    v.lang && v.lang.toLowerCase().startsWith("en-gb")
                );

                const englishVoice = voices.find(v =>
                    v.lang && v.lang.toLowerCase().startsWith("en")
                );

                utterance.voice = maleGbVoice || gbVoice || englishVoice || null;

                window.speechSynthesis.cancel();
                window.speechSynthesis.speak(utterance);
            }}

            if (window.speechSynthesis.getVoices().length === 0) {{
                window.speechSynthesis.onvoiceschanged = setVoiceAndSpeak;
            }} else {{
                setVoiceAndSpeak();
            }}
        }}
        </script>
        """,
        height=height,
    )


def render_vocab_row(vocab: str, ipa: str, meaning: str, key: str):
    vocab = clean_text(vocab)
    ipa = clean_text(ipa)
    meaning = clean_text(meaning)

    if not vocab and not ipa and not meaning:
        return

    col1, col2 = st.columns([8, 1])

    with col1:
        parts = []
        if vocab:
            parts.append(f"<b>{html.escape(vocab)}</b>")
        if ipa:
            parts.append(f"<span class='ipa-text'>/{html.escape(ipa.strip('/'))}/</span>")
        if meaning:
            parts.append(f"<span style='color:#4b5563;'>= {html.escape(meaning)}</span>")

        st.markdown(
            f"<div class='vocab-chip'>{' &nbsp; '.join(parts)}</div>",
            unsafe_allow_html=True,
        )

    with col2:
        render_speak_button(vocab, key=f"vocab_{key}", label="🔊")


def render_topic(topic_key: str, df: pd.DataFrame):
    topic_id = clean_text(df["Topic ID"].iloc[0]) if "Topic ID" in df.columns else topic_key
    topic_name = clean_text(df["Topic"].iloc[0]) if "Topic" in df.columns else topic_key

    st.subheader(f"{topic_id}. {topic_name}")

    definitions = []
    seen = set()
    last_def_id = ""
    last_def = ""

    for _, row in df.iterrows():
        def_id = clean_text(row.get("Definition ID", ""))
        definition = clean_text(row.get("Definition", ""))

        if definition:
            last_def_id = def_id
            last_def = definition

        group_key = f"{last_def_id}__{last_def}"
        if last_def and group_key not in seen:
            seen.add(group_key)
            definitions.append((group_key, last_def_id, last_def))

    for group_key, def_id, definition in definitions:
        st.markdown(
            f"""
            <div class="definition-box">
                {html.escape(def_id)} &nbsp; {html.escape(definition)}
            </div>
            """,
            unsafe_allow_html=True,
        )

        current_rows = []
        active_def_id = ""
        active_def = ""

        for _, row in df.iterrows():
            row_def_id = clean_text(row.get("Definition ID", ""))
            row_def = clean_text(row.get("Definition", ""))

            if row_def:
                active_def_id = row_def_id
                active_def = row_def

            if f"{active_def_id}__{active_def}" == group_key:
                current_rows.append(row)

        for i, row in enumerate(current_rows):
            detail_id = clean_text(row.get("Detail ID", ""))
            detail = clean_text(row.get("Detail", ""))
            example_id = clean_text(row.get("Example ID", ""))
            example = clean_text(row.get("Example", ""))
            example_vi = clean_text(row.get("Example VI", ""))
            vocab = clean_text(row.get("Vocabulary", ""))
            meaning = clean_text(row.get("Meaning VI", ""))
            ipa = clean_text(row.get("IPA", ""))

            if detail_id or detail:
                st.markdown(
                    f"<div class='detail-id'>{html.escape(detail_id)}</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"<div class='detail-text'>{html.escape(detail)}</div>",
                    unsafe_allow_html=True,
                )

            if example:
                col_ex, col_audio = st.columns([8, 1])

                with col_ex:
                    id_prefix = f"{html.escape(example_id)} — " if example_id else ""
                    st.markdown(
                        f"<div class='example-text'>{id_prefix}{html.escape(example)}</div>",
                        unsafe_allow_html=True,
                    )

                    if example_vi:
                        st.markdown(
                            f"<div class='translation-text'>🇻🇳 {html.escape(example_vi)}</div>",
                            unsafe_allow_html=True,
                        )

                with col_audio:
                    render_speak_button(example, key=f"ex_{topic_id}_{def_id}_{i}")

            if vocab or ipa or meaning:
                render_vocab_row(
                    vocab=vocab,
                    ipa=ipa,
                    meaning=meaning,
                    key=f"{topic_id}_{def_id}_{i}",
                )

            st.markdown("<hr>", unsafe_allow_html=True)


def build_message_from_inputs(topic, sentence, vocab, grammar, note, exercise):
    lines = [
        f"📘 English sentence today – {date.today().strftime('%d/%m/%Y')}",
        "",
        f"Topic: {topic}",
        "",
        "Today’s sentence:",
        sentence,
        "",
        "Vocabulary:",
    ]

    vocab_items = [v.strip() for v in vocab.split(",") if v.strip()]
    for idx, item in enumerate(vocab_items, 1):
        lines.append(f"{idx}. {item}")

    if grammar:
        lines += ["", f"Grammar point: {grammar}"]
    if note:
        lines += ["", f"Pronunciation note: {note}"]
    if exercise:
        lines += ["", f"Practice: {exercise}"]

    return "\n".join(lines)


# ============================================================
# MAIN APP
# ============================================================

st.markdown("<div class='main-title'>📘 English Teaching Dashboard</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='sub-title'>Nhập bài học, quản lý roadmap, tạo message gửi bạn học và phát âm Anh-Anh bằng trình duyệt.</div>",
    unsafe_allow_html=True,
)

sheets = load_data()

topic_tab_names = []
topic_sheet_items = list(sheets.items())

for sheet_name, df in topic_sheet_items:
    topic_id = clean_text(df["Topic ID"].iloc[0]) if "Topic ID" in df.columns else sheet_name
    topic_name = clean_text(df["Topic"].iloc[0]) if "Topic" in df.columns else sheet_name
    topic_tab_names.append(f"{topic_id}. {topic_name}")

extra_tabs = [
    "🏠 Dashboard",
    "📝 Nhập bài học",
    "🔁 Kiểm tra trùng từ",
    "✉️ Message gửi bạn học",
]

tabs = st.tabs(extra_tabs + topic_tab_names)

with tabs[0]:
    st.header("Tổng quan")

    total_topics = len(sheets)
    total_definitions = 0
    total_examples = 0
    total_vocab = 0

    for df in sheets.values():
        total_definitions += df["Definition"].replace("", pd.NA).dropna().nunique()
        total_examples += (df["Example"].astype(str).str.strip() != "").sum()
        total_vocab += (df["Vocabulary"].astype(str).str.strip() != "").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Topic", total_topics)
    c2.metric("Definition", total_definitions)
    c3.metric("Example", total_examples)
    c4.metric("Vocabulary", total_vocab)

    st.info(
        "Nội dung được đọc từ file `data/English.xlsx`. "
        "Muốn cập nhật bền vững trên Streamlit Cloud, hãy upload file Excel mới lên GitHub rồi reboot app."
    )

with tabs[1]:
    st.header("Nhập bài học hôm nay")

    with st.form("lesson_form"):
        topic = st.text_input("Topic", value="Construction English")
        sentence = st.text_area(
            "Mẫu câu hôm nay",
            value="In our toolbox meeting today, the project manager announced that we must complete construction of the structural package by August 31st.",
            height=100,
        )
        vocab = st.text_input(
            "Từ vựng chính, cách nhau bằng dấu phẩy",
            value="toolbox meeting, project manager, announced, structural package",
        )
        grammar = st.text_input("Grammar point", value="must + verb / reported announcement")
        note = st.text_input("Ghi chú phát âm", value="31st = thirty-first")
        exercise = st.text_input("Bài tập", value='Make 2 similar sentences using "announced that".')
        submitted = st.form_submit_button("Tạo message")

    if submitted:
        msg = build_message_from_inputs(topic, sentence, vocab, grammar, note, exercise)
        st.text_area("Message để copy gửi bạn học", value=msg, height=260)
        st.caption("Tab này hiện tạo message để copy. Muốn lưu bền vững cần ghi vào Google Sheet hoặc upload lại Excel lên GitHub.")

with tabs[2]:
    st.header("Kiểm tra trùng từ vựng")

    all_vocab = []
    for sheet_name, df in sheets.items():
        for _, row in df.iterrows():
            word = clean_text(row.get("Vocabulary", ""))
            if word:
                all_vocab.append(
                    {
                        "Sheet": sheet_name,
                        "Topic": clean_text(row.get("Topic", "")),
                        "Vocabulary": word.lower(),
                        "Display": word,
                        "IPA": clean_text(row.get("IPA", "")),
                        "Meaning": clean_text(row.get("Meaning VI", "")),
                    }
                )

    vocab_df = pd.DataFrame(all_vocab)

    today_vocab = st.text_input(
        "Nhập từ vựng hôm nay, cách nhau bằng dấu phẩy",
        value="toolbox meeting, project manager, announced, structural package",
    )

    words = [w.strip().lower() for w in today_vocab.split(",") if w.strip()]

    if st.button("Kiểm tra"):
        if vocab_df.empty:
            st.warning("Chưa có dữ liệu từ vựng trong Excel.")
        else:
            dup = vocab_df[vocab_df["Vocabulary"].isin(words)]

            if dup.empty:
                st.success("Không phát hiện từ vựng trùng trong dữ liệu hiện có.")
            else:
                st.warning("Có từ vựng đã xuất hiện trong dữ liệu:")
                st.dataframe(dup[["Sheet", "Topic", "Display", "IPA", "Meaning"]], use_container_width=True)

with tabs[3]:
    st.header("Message gửi bạn học")

    st.write("Có thể copy mẫu này để gửi Zalo/Teams:")

    sample = build_message_from_inputs(
        topic="Construction deadline",
        sentence="In our toolbox meeting today, the project manager announced that we must complete construction of the structural package by August 31st.",
        vocab="toolbox meeting = họp đầu giờ, project manager = quản lý dự án, announced = đã thông báo, structural package = gói kết cấu",
        grammar="must + verb",
        note="31st is read as thirty-first.",
        exercise='Make 2 similar sentences using "announced that".',
    )

    st.text_area("Sample message", value=sample, height=280)

for tab, (sheet_name, df) in zip(tabs[4:], topic_sheet_items):
    with tab:
        render_topic(sheet_name, df)
