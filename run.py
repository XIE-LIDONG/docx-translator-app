import streamlit as st
import tempfile
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import os
import re  # 通用文本预处理

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

# Define supported languages
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

# 通用文本预处理函数（适配所有语言，解决格式/特殊字符导致的跳过问题）
def clean_text(text):
    """
    清理所有语言文本中的隐藏格式/无效字符，避免被误判为空或翻译接口无法识别
    保留：各国语言核心字符 + 数字 + 基本标点
    """
    # 1. 移除隐藏控制字符（如换行符、制表符、双向文本标记等）
    text = re.sub(r'[\x00-\x1F\x7F-\x9F\u200B-\u200F\u202A-\u202E]', ' ', text)
    # 2. 移除多余空格/重复标点，保留单个空格分隔
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'([.,!?;:()\-])\1+', r'\1', text)
    # 3. 首尾去空格
    return text.strip()

# File upload
uf = st.file_uploader("Select Word Document (.docx)", type=["docx"])

if uf:
    # File info
    c1, c2 = st.columns([2,1])
    with c1: st.success(f"📁 **File:** {uf.name}")
    with c2: st.metric("Size", f"{uf.size/(1024*1024):.2f} MB")
    st.markdown("---")

    # Translation settings
    c1, c2, c3 = st.columns(3)
    with c1:
        source_lang_name = st.selectbox(
            "**Source Language**", LANG_NAMES, index=LANG_NAMES.index("French")
        )
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox(
            "**Target Language**", LANG_NAMES, index=LANG_NAMES.index("English")
        )
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        wk = st.slider("**Thread Count**", 1, 3, 1)  # 保留1-3线程

    if st.button("🚀 Start Translation", type="primary", use_container_width=True):
        log_area = st.empty()
        log = []
        filtered_texts = []  # 记录被过滤的文本（方便排查）

        # 临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uf.getvalue())
            fp = tmp.name

        try:
            stt = time.time()
            doc = Document(fp)
            ti, at = [], []  # (段落对象, 清理后文本), 待翻译文本列表
            pc, cc = 0, 0

            # 提取文本 + 通用预处理（核心优化：避免假空文本被过滤）
            def extract_and_clean(p):
                """提取并清理段落文本，返回(是否有效, 清理后文本)"""
                raw_txt = p.text
                cleaned_txt = clean_text(raw_txt)
                # 判定有效文本：清理后长度≥1，且不是纯标点/空格
                is_valid = len(cleaned_txt) > 0 and not re.match(r'^[.,!?;:()\- ]+$', cleaned_txt)
                if not is_valid and raw_txt.strip():
                    filtered_texts.append(f"[过滤] 原文本：{raw_txt[:50]}...（清理后无有效内容）")
                return is_valid, cleaned_txt

            # 提取段落文本
            for p in doc.paragraphs:
                is_valid, cleaned_txt = extract_and_clean(p)
                if is_valid:
                    ti.append((p, cleaned_txt))
                    at.append(cleaned_txt)
                    pc += 1

            # 提取表格文本（重点解决表格整段跳过问题）
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs:
                            is_valid, cleaned_txt = extract_and_clean(p)
                            if is_valid:
                                ti.append((p, cleaned_txt))
                                at.append(cleaned_txt)
                                cc += 1

            total = len(at)
            if total == 0:
                st.error("❌ No valid text in document")
                # 显示被过滤的文本，方便排查
                if filtered_texts:
                    st.expander("🔍 Filtered Texts (Click to View)", expanded=True).write("\n".join(filtered_texts[:20]))
                st.stop()

            # 初始日志
            log.append(f"✅ Extracted {total} valid text segments (Paragraphs: {pc}, Table Cells: {cc})")
            log.append(f"🔤 Translation direction: {source_lang_name} → {target_lang_name}")
            if filtered_texts:
                log.append(f"⚠️ Filtered {len(filtered_texts)} invalid text segments (check special characters)")
            log_area.markdown("\n".join(log))

            # 多线程批量翻译（完全保留你的逻辑）
            ta = [None]*total
            BS = 100  # 保留100条批次
            def tb(txts):
                """批量翻译 + 空结果兜底"""
                try:
                    res = GoogleTranslator(source=source_lang, target=target_lang).translate_batch(txts)
                    # 兜底：空翻译结果替换为原文
                    return [r if r and r.strip() else txt for r, txt in zip(res, txts)]
                except Exception as e:
                    # 翻译失败时返回原文
                    log.append(f"⚠️ Batch translation error: {str(e)[:50]}")
                    return txts

            # 提交多线程任务
            with ThreadPoolExecutor(max_workers=wk) as exe:
                futs = {}
                for i in range(0, total, BS):
                    batch = at[i:i+BS]
                    fut = exe.submit(tb, batch)
                    futs[fut] = i

                # 处理结果
                for fut in as_completed(futs):
                    start_idx = futs[fut]
                    res = fut.result()
                    for idx in range(len(res)):
                        if start_idx+idx < total:
                            ta[start_idx+idx] = res[idx]
                    # 进度日志
                    done = sum(1 for x in ta if x is not None)
                    if done % 10 == 0:
                        log.append(f"🔄 Translating: {done}/{total}")
                        log_area.markdown("\n".join(log))

            # 更新文档（确保无空白替换）
            log.append(f"✅ Translation completed: {total}/{total}")
            log.append("📝 Updating document...")
            log_area.markdown("\n".join(log))
            for idx, (p_obj, original_txt) in enumerate(ti):
                translated_txt = ta[idx] or original_txt  # 最终兜底
                p_obj.text = translated_txt

            # 保存下载
            op = fp.replace(".docx", "_translated.docx")
            doc.save(op)
            tot_t = time.time()-stt

            # 结果展示
            st.balloons()
            st.success(f"### ✅ Translation Completed!（{source_lang_name} → {target_lang_name}）")
            c1,c2 = st.columns(2)
            c1.metric("Total Time", f"{tot_t:.1f}s")
            c2.metric("Average Speed", f"{total/tot_t:.1f} segments/s")

            # 显示过滤日志（方便排查）
            if filtered_texts:
                st.expander("🔍 Filtered Texts (Click to View)", expanded=False).write("\n".join(filtered_texts[:20]))

            # 下载按钮
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
            # 清理临时文件
            try:
                os.unlink(fp)
                if 'op' in locals() and os.path.exists(op):
                    os.unlink(op)
            except Exception as cleanup_e:
                st.warning(f"⚠️ Temporary file cleanup failed: {cleanup_e}")
