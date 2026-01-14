import streamlit as st
import tempfile
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import os

# Page configuration
st.set_page_config(page_title="DOCX Document language Translator", page_icon="📄", layout="wide")

# 标题+署名：同一行布局（标题左，署名右）
st.markdown(
    """
    <div style='display: flex; justify-content: space-between; align-items: center;'>
        <h1 style='margin: 0;'>📄 DOCX Document Language Translator</h1>
        <p style='color: #666666; font-size: 14px; margin: 0;'>By XIE LI DONG</p>
    </div>
    """,
    unsafe_allow_html=True
)

# Added usage tip below title
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
# Extract language name list (for dropdown)
LANG_NAMES = list(SUPPORT_LANGUAGES.keys())

# File upload
uf = st.file_uploader("Select Word Document (.docx)", type=["docx"])

if uf:
    # File information
    c1, c2 = st.columns([2,1])
    with c1: st.success(f"📁 **File:** {uf.name}")
    with c2: st.metric("Size", f"{uf.size/(1024*1024):.2f} MB")
    st.markdown("---")

    # Translation settings (multilingual dropdown + 线程/批次配置)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        source_lang_name = st.selectbox(
            "**Source Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("English")  # 优化默认值：英文
        )
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox(
            "**Target Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("Chinese")  # 优化默认值：中文
        )
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        wk = st.slider(
            "**Thread Count**",
            min_value=1,
            max_value=3,
            value=2,
            help="Number of parallel translation threads (1-3 for stability)"
        )
    with c4:
        BS = st.slider(
            "**Batch Size**",
            min_value=20,
            max_value=100,
            value=100,
            step=10,
            help="Number of text segments per translation batch (20-100)"
        )

    if st.button("🚀 Start Translation", type="primary", use_container_width=True):
        # Progress log area
        log_area = st.empty()
        log = []

        # Temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uf.getvalue())
            fp = tmp.name

        try:
            stt = time.time()
            # Parse document
            doc = Document(fp)
            # ✅ 修复：修改存储结构，存(对象, 文本, 类型)，区分段落/单元格，解决表格回写问题
            text_items = []  # 格式: [(obj, text, type), ...] type: 'paragraph'/'cell'
            all_texts = []   # 纯文本列表，用于翻译
            para_count = 0
            cell_count = 0

            # ✅ 提取【普通段落】文本 - 原有逻辑保留（无问题）
            for para in doc.paragraphs:
                if text := para.text.strip():
                    text_items.append((para, text, 'paragraph'))
                    all_texts.append(text)
                    para_count += 1

            # ✅ 核心修复：提取【表格单元格】文本，修正逻辑+兼容合并单元格
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text.strip()  # ✅ 正确写法：直接取cell.text
                        if cell_text:
                            text_items.append((cell, cell_text, 'cell'))
                            all_texts.append(cell_text)
                            cell_count += 1

            total = len(all_texts)
            if total == 0:
                st.error("❌ No valid text in document")
                st.stop()

            # Initial log
            log.append(f"✅ Extracted {total} text segments for translation")
            log.append(f"📄 {para_count} Paragraphs | {cell_count} Table Cells")
            log.append(f"🔤 Translation direction: {source_lang_name} → {target_lang_name}")
            log.append(f"⚙️ Configuration: {wk} threads | {BS} segments per batch")
            log_area.markdown("\n".join(log))

            # 翻译结果存储
            translations = [None]*total

            # 批次翻译函数
            def translate_batch(txt_list):
                try:
                    res = GoogleTranslator(source=source_lang, target=target_lang).translate_batch(txt_list)
                    # ✅ 修复：翻译结果非空兜底，避免返回None导致文本丢失
                    return [r if r is not None else txt for r, txt in zip(res, txt_list)]
                except Exception:
                    return txt_list  # 翻译失败时返回原文

            # 多线程执行翻译
            with ThreadPoolExecutor(max_workers=wk) as exe:
                futures = {}
                # 切割批次，无越界风险
                for start_idx in range(0, total, BS):
                    end_idx = min(start_idx + BS, total)  # ✅ 修复：增加边界判断，防止索引越界
                    batch_texts = all_texts[start_idx:end_idx]
                    future = exe.submit(translate_batch, batch_texts)
                    futures[future] = (start_idx, end_idx)

                # 处理结果+实时日志
                for future in as_completed(futures):
                    s_idx, e_idx = futures[future]
                    batch_res = future.result()
                    # 赋值翻译结果
                    for idx in range(len(batch_res)):
                        translations[s_idx + idx] = batch_res[idx]
                    # 进度日志
                    done = sum(1 for x in translations if x is not None)
                    if done % 10 == 0:
                        log.append(f"🔄 Translating: {done}/{total} segments completed")
                        log_area.markdown("\n".join(log))

            # 完成翻译日志
            log.append(f"✅ Translation completed: {total}/{total} segments")
            log_area.markdown("\n".join(log))

            # ✅ 核心修复：文本回写，区分【段落】和【表格单元格】两种类型，分别赋值
            log.append("📝 Updating document content (including tables)...")
            log_area.markdown("\n".join(log))
            for idx, (obj, original_text, obj_type) in enumerate(text_items):
                trans_text = translations[idx]
                if trans_text and trans_text != original_text:
                    if obj_type == 'paragraph':
                        obj.text = trans_text  # 段落直接赋值
                    elif obj_type == 'cell':
                        obj.text = trans_text   # 单元格直接赋值

            # 保存翻译后的文档
            output_path = fp.replace(".docx", "_translated.docx")
            doc.save(output_path)
            total_time = time.time() - stt

            # 成功提示
            st.balloons()
            st.success(f"### ✅ Translation Completed!（{source_lang_name} → {target_lang_name}）")
            c1,c2 = st.columns(2)
            c1.metric("Total Time", f"{total_time:.1f}s")
            c2.metric("Average Speed", f"{total/total_time:.1f} segments/s")

            # 下载按钮
            st.markdown("---")
            with open(output_path, "rb") as f:
                st.download_button(
                    "📥 Download Translated Document", f,
                    file_name=f"{source_lang_name}2{target_lang_name}_{uf.name}",
                    use_container_width=True, type="primary"
                )

        except Exception as e:
            st.error("### ❌ Translation Failed")
            st.exception(e)
        finally:
            # ✅ 优化：彻底清理临时文件，增加异常捕获，避免残留
            try:
                if os.path.exists(fp):
                    os.unlink(fp)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.unlink(output_path)
            except Exception as cleanup_e:
                st.warning(f"⚠️ Temporary file cleanup warning: {str(cleanup_e)[:50]}...")
