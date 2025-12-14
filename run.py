import streamlit as st
import tempfile
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import os
import re  # 新增：用于文本预处理

# Page configuration
st.set_page_config(page_title="DOCX Translator", page_icon="📄", layout="wide")
st.title("🚀 DOCX Document Translator (Multilingual Translation)")

# Usage tip
st.info("""
💡 **Usage Tip**: If your file is in PDF format, convert it to DOCX first via [ILovePDF](https://www.ilovepdf.com/). 
After translation, you can convert the DOCX back to PDF using ILovePDF if needed. 
We tried integrating PDF-DOCX conversion directly into Streamlit, but it drastically slowed down the entire application.
""")
st.markdown("---")

# Define supported languages (Name: deep-translator code)
SUPPORT_LANGUAGES = {
    "Chinese": "zh-CN",
    "English": "en",
    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Arabic": "ar",
    "Japanese": "ja",
    "Korean": "ko"
}
LANG_NAMES = list(SUPPORT_LANGUAGES.keys())

# 新增：文本预处理函数（重点解决阿拉伯语格式问题）
def clean_special_text(text, source_lang):
    if source_lang == "ar":  # 仅对阿拉伯语做预处理
        # 1. 保留阿拉伯语核心字符（字母、数字、基本标点），清除隐藏格式/特殊符号
        # \u0600-\u06FF：阿拉伯语字母范围；\u0660-\u0669：阿拉伯数字；\u06F0-\u06F9：扩展阿拉伯数字
        cleaned = re.sub(r'[^\u0600-\u06FF\u0660-\u0669\u06F0-\u06F9\s.,!?;:()\-]', '', text)
        # 2. 清除多余空格和换行
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned
    # 其他语言仅清除多余空格
    return re.sub(r'\s+', ' ', text).strip()

# File upload
uf = st.file_uploader("Select Word Document (.docx)", type=["docx"])

if uf:
    # File information
    c1, c2 = st.columns([2,1])
    with c1: st.success(f"📁 **File:** {uf.name}")
    with c2: st.metric("Size", f"{uf.size/(1024*1024):.2f} MB")
    st.markdown("---")

    # Translation settings
    c1, c2, c3 = st.columns(3)
    with c1:
        source_lang_name = st.selectbox(
            "**Source Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("French")
        )
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox(
            "**Target Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("English")
        )
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        wk = st.slider("**Thread Count**", 1, 3, 1)

    if st.button("🚀 Start Translation", type="primary", use_container_width=True):
        log_area = st.empty()
        log = []

        # Temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uf.getvalue())
            fp = tmp.name

        try:
            stt = time.time()
            doc = Document(fp)
            ti, at = [], []
            pc, cc = 0, 0

            # 提取文本 + 预处理（重点优化阿拉伯语）
            for p in doc.paragraphs:
                if txt := p.text.strip():
                    cleaned_txt = clean_special_text(txt, source_lang)  # 预处理
                    if cleaned_txt:  # 确保预处理后有有效文本
                        ti.append((p, cleaned_txt))
                        at.append(cleaned_txt)
                        pc += 1
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs:
                            if txt := p.text.strip():
                                cleaned_txt = clean_special_text(txt, source_lang)  # 预处理
                                if cleaned_txt:
                                    ti.append((p, cleaned_txt))
                                    at.append(cleaned_txt)
                                    cc += 1

            total = len(at)
            if total == 0:
                st.error("❌ No valid text in document")
                st.stop()

            log.append(f"✅ Extracted {total} text segments for translation")
            log.append(f"🔤 Translation direction: {source_lang_name} → {target_lang_name}")
            log_area.markdown("\n".join(log))

            # 重写翻译函数：单条翻译+重试，解决阿拉伯语批量翻译丢失问题
            ta = [None]*total
            BS = 50  # 减小批次大小（阿拉伯语建议更小）
            def tb(txts):
                translator = GoogleTranslator(source=source_lang, target=target_lang)
                results = []
                for txt in txts:
                    max_retries = 3
                    success = False
                    for retry in range(max_retries):
                        try:
                            # 单条翻译，避免批量问题
                            translated = translator.translate(txt)
                            results.append(translated)
                            time.sleep(0.3)  # 小延迟，降低风控概率
                            success = True
                            break
                        except Exception as e:
                            time.sleep(2 ** retry)  # 指数退避重试
                    if not success:
                        # 重试失败则保留原文并标记
                        results.append(f"[Untranslated] {txt}")
                return results

            # 多线程执行
            with ThreadPoolExecutor(max_workers=wk) as exe:
                futs = {}
                for i in range(0, total, BS):
                    batch = at[i:i+BS]
                    fut = exe.submit(tb, batch)
                    futs[fut] = i

                for fut in as_completed(futs):
                    start_idx = futs[fut]
                    res = fut.result()
                    for idx in range(len(res)):
                        if start_idx+idx < total:
                            ta[start_idx+idx] = res[idx]
                    done = sum(1 for x in ta if x is not None)
                    if done % 5 == 0:  # 更频繁的日志（阿拉伯语翻译慢，让用户看到进度）
                        log.append(f"🔄 Translating: {done}/{total}")
                        log_area.markdown("\n".join(log))

            log.append(f"✅ Translation completed: {total}/{total}")
            log_area.markdown("\n".join(log))

            # 更新文档
            log.append("📝 Updating document...")
            log_area.markdown("\n".join(log))
            for idx, (p_obj, _) in enumerate(ti):
                if ta[idx] and not ta[idx].startswith("[Untranslated]"):
                    p_obj.text = ta[idx]
                elif ta[idx]:
                    # 未翻译的文本，保留原文并提示
                    p_obj.text = ta[idx]

            # 保存下载
            op = fp.replace(".docx", "_translated.docx")
            doc.save(op)
            tot_t = time.time()-stt

            st.balloons()
            st.success(f"### ✅ Translation Completed!（{source_lang_name} → {target_lang_name}）")
            c1,c2 = st.columns(2)
            c1.metric("Total Time", f"{tot_t:.1f}s")
            c2.metric("Average Speed", f"{total/tot_t:.1f} segments/s")

            # 统计未翻译数量（方便排查）
            untranslated = sum(1 for x in ta if x and x.startswith("[Untranslated]"))
            if untranslated > 0:
                st.warning(f"⚠️ {untranslated} segments were untranslated (check special characters)")

            st.markdown("---")
            with open(op, "rb") as f:
                st.download_button(
                    "📥 Download Translated Document", f,
                    file_name=f"{source_lang_name}2{target_lang_name}_{uf.name}",
                    use_container_width=True, type="primary"
                )

        except Exception as e:
            st.error("### ❌ Translation Failed")
            st.exception(e)
        finally:
            try:
                os.unlink(fp)
                if 'op' in locals() and os.path.exists(op):
                    os.unlink(op)
            except Exception as cleanup_e:
                st.warning(f"⚠️ Temporary file cleanup failed: {cleanup_e}")
