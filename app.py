import os
import re
import tempfile
import warnings
import streamlit as st
import whisper
from docx import Document
from opencc import OpenCC

warnings.filterwarnings(
    "ignore", message="FP16 is not supported on CPU; using FP32 instead"
)

st.set_page_config(
    page_title="MP3 語音轉 Word 逐字稿",
    page_icon="🎙️",
    layout="centered",
)

st.title("🎙️ MP3 語音轉 Word 逐字稿工具")
st.caption("上傳本機 MP3 語音檔，辨識中文與方言內容、整理中文標點，並匯出 .docx 下載。")


@st.cache_resource(show_spinner="正在載入語音辨識模型...")
def load_whisper_model(model_size: str):
    return whisper.load_model(model_size)


@st.cache_resource(show_spinner=False)
def get_opencc_converter(variant: str):
    if variant == "hk":
        return OpenCC("s2hk")
    if variant == "tw":
        return OpenCC("s2twp")
    return None


def format_chinese_punctuation(text: str, variant: str = "hk") -> str:
    """清理多餘空白、轉換繁體（依設定），並按中文語法整理與補齊標點符號。"""
    if not text:
        return ""

    s = text.strip()
    converter = get_opencc_converter(variant)
    if converter is not None:
        s = converter.convert(s)

    # 移除中文字元之間多餘的半形空白
    s = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", s)

    # 半形標點轉全形中文標點
    replacements = {
        ",": "，",
        "。": "。",
        "?": "？",
        "!": "！",
        ":": "：",
        ";": "；",
    }
    for half_p, full_p in replacements.items():
        s = s.replace(half_p, full_p)

    # 合併重複標點
    s = re.sub(r"，{2,}", "，", s)
    s = re.sub(r"。{2,}", "。", s)
    s = re.sub(r"？{2,}", "？", s)
    s = re.sub(r"！{2,}", "！", s)

    # 句尾補齊適當標點
    question_endings = (
        "嗎", "呢", "吧", "嘛", "啊", "呀", "咩", "囉",
        "唔係", "係咪", "點解", "邊度", "幾時", "乜嘢",
    )
    if not s.endswith(("。", "！", "？", "…", "」", "』", "）")):
        if any(s.endswith(q) for q in question_endings):
            s = s + "？"
        else:
            s = s + "。"

    # 常見連接詞與轉折詞前補逗號
    conjunctions = (
        "但是", "不過", "然而", "所以", "因此", "另外",
        "而且", "然後", "接著", "其實", "舉例來說", "總結來說",
    )
    for conj in conjunctions:
        s = re.sub(
            rf"(?<!^)(?<![，。！？；：、]){re.escape(conj)}",
            f"，{conj}",
            s,
        )

    s = s.lstrip("，")
    s = re.sub(r"，+", "，", s)
    return s.strip()


def build_docx_bytes(
    audio_filename: str,
    full_text: str,
    formatted_segments: list[dict],
) -> bytes:
    """生成 Word (.docx) 並回傳二進位資料供 Streamlit 下載。"""
    doc = Document()
    doc.add_heading(f"語音轉文字逐字稿 - {audio_filename}", level=1)

    doc.add_heading("完整內容（已套用中文標點）", level=2)
    doc.add_paragraph(full_text if full_text else "（未辨識出有效語音內容）")

    if formatted_segments:
        doc.add_heading("分段時間軸記錄", level=2)
        for item in formatted_segments:
            s_min, s_sec = divmod(item["start"], 60)
            e_min, e_sec = divmod(item["end"], 60)
            time_range = f"[{s_min:02d}:{s_sec:02d} - {e_min:02d}:{e_sec:02d}]"
            doc.add_paragraph(f"{time_range} {item['text']}")

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
        tmp_docx_path = tmp_docx.name

    try:
        doc.save(tmp_docx_path)
        with open(tmp_docx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_docx_path):
            os.remove(tmp_docx_path)


# ==========================================
# 側邊欄參數設定
# ==========================================
with st.sidebar:
    st.header("辨識與輸出設定")
    model_size = st.selectbox(
        "Whisper 模型大小",
        options=["tiny", "base", "small", "medium", "large"],
        index=1,  # 預設 base，在雲端免費資源下較穩健；準確率需求高可改 small/medium
        help="模型越大準確率越高，但記憶體占用與運算時間也越長。",
    )
    lang_mode = st.selectbox(
        "語音語言模式",
        options=["auto", "zh"],
        format_func=lambda x: "自動偵測（適合方言／中英混雜）" if x == "auto" else "固定中文（zh）",
        index=0,
    )
    chinese_variant = st.selectbox(
        "繁體中文轉換",
        options=["hk", "tw", "none"],
        format_func=lambda x: {
            "hk": "繁體中文（港／澳用語）",
            "tw": "繁體中文（台灣用語）",
            "none": "維持原始輸出",
        }[x],
        index=0,
    )
    custom_prompt = st.text_area(
        "語境提示詞（選填）",
        value="以下是中文及方言語音內容，請使用正確的中文標點符號（，。？！：；）與語法輸出。",
        height=80,
    )

# ==========================================
# 主畫面：上傳 MP3、轉錄與下載 Word
# ==========================================
uploaded_file = st.file_uploader(
    "請選擇本機的 MP3 語音檔",
    type=["mp3"],
    accept_multiple_files=False,
)

if uploaded_file is not None:
    st.audio(uploaded_file, format="audio/mp3")

    if st.button("開始轉換為 Word", type="primary"):
        tmp_audio_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp_audio:
                tmp_audio.write(uploaded_file.getbuffer())
                tmp_audio_path = tmp_audio.name

            model = load_whisper_model(model_size)
            language_code = None if lang_mode == "auto" else "zh"

            with st.spinner("正在進行語音辨識與中文標點整理，請稍候..."):
                result = model.transcribe(
                    tmp_audio_path,
                    language=language_code,
                    initial_prompt=custom_prompt.strip() or None,
                )

            raw_segments = result.get("segments", [])
            formatted_segments = []
            for seg in raw_segments:
                seg_clean = format_chinese_punctuation(
                    seg.get("text", ""), variant=chinese_variant
                )
                if seg_clean:
                    formatted_segments.append(
                        {
                            "start": int(seg["start"]),
                            "end": int(seg["end"]),
                            "text": seg_clean,
                        }
                    )

            if formatted_segments:
                full_text = "".join(item["text"] for item in formatted_segments)
            else:
                full_text = format_chinese_punctuation(
                    result.get("text", ""), variant=chinese_variant
                )

            docx_bytes = build_docx_bytes(
                audio_filename=uploaded_file.name,
                full_text=full_text,
                formatted_segments=formatted_segments,
            )

            base_name = os.path.splitext(uploaded_file.name)[0]
            output_filename = f"{base_name}_transcript.docx"

            st.success("語音轉文字完成！請點擊下方按鈕下載 Word 檔。")

            st.download_button(
                label=f"⬇️ 下載 {output_filename}",
                data=docx_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )

            with st.expander("預覽辨識結果", expanded=True):
                st.write("**完整內容：**")
                st.write(full_text if full_text else "（未辨識出有效語音內容）")
                if formatted_segments:
                    st.write("**分段時間軸：**")
                    for item in formatted_segments:
                        s_min, s_sec = divmod(item["start"], 60)
                        e_min, e_sec = divmod(item["end"], 60)
                        st.text(
                            f"[{s_min:02d}:{s_sec:02d} - {e_min:02d}:{e_sec:02d}] {item['text']}"
                        )

        except Exception as e:
            st.error(f"轉換過程中發生錯誤：{e}")
        finally:
            if tmp_audio_path and os.path.exists(tmp_audio_path):
                os.remove(tmp_audio_path)
