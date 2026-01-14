import streamlit as st
import tempfile
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import os
import io  # 新增：用于完整读取上传文件

# 1. 页面配置
st.set_page_config(page_title="DOCX Document Language Translator", page_icon="📄", layout="wide")

# 标题+署名
st.markdown(
    """
    <div style='display: flex; justify-content: space-between; align-items: center;'>
        <h1 style='margin: 0;'>📄 DOCX Document Language Translator</h1>
        <p style='color: #666666; font-size: 14px; margin: 0;'>By XIE LI DONG</p>
    </div>
    """,
    unsafe_allow_html=True
)

st.info("""
💡 **Usage Tip**: For PDF files, convert to DOCX via [ILovePDF](https://www.ilovepdf.com/). 
Supports multi-page documents (including cross-page tables).
""")
st.markdown("---")

# 2. 支持语言
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

# 3. 文件上传
uf = st.file_uploader("Select Word Document (.docx)", type=["docx"])

if uf:
    # 显示文件信息（增加文件大小对比，验证是否完整）
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        st.success(f"📁 **File**: {uf.name}")
    with c2:
        upload_size = uf.size / (1024 * 1024)
        st.metric("Upload Size", f"{upload_size:.2f} MB")
    with c3:
        # 提前读取文件内容，用于后续验证
        uf_content = uf.getvalue()
        st.metric("Content Size", f"{len(uf_content)/(1024*1024):.2f} MB")
    st.markdown("---")

    # 4. 翻译参数配置
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        source_lang_name = st.selectbox("**Source Language**", LANG_NAMES, index=LANG_NAMES.index("English"))
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox("**Target Language**", LANG_NAMES, index=LANG_NAMES.index("Chinese"))
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        wk = st.slider("**Thread Count**", 1, 3, 2, help="1-3 for stability")
    with c4:
        BS = st.slider("**Batch Size**", 20, 100, 50, step=10, help="Segments per batch")

    if st.button("🚀 Start Translation", type="primary", use_container_width=True):
        log_area = st.empty()
        log = []
        doc = None
        fp = None
        output_path = None
        text_items = []  # (对象, 原文, 类型: paragraph/cell)
        all_texts = []
        translations = []
        para_count = 0
        cell_count = 0

        try:
            stt = time.time()

            # ========== 修复1：完整保存临时文件（关键！解决多页读取问题） ==========
            # 方案：用 io.BytesIO 先缓存内容，再写入临时文件，确保完整
            with io.BytesIO(uf_content) as bio:
                # 创建临时文件（mode='wb+' 支持读写）
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", mode='wb+') as tmp:
                    tmp.write(bio.getvalue())
                    tmp.flush()  # 强制刷新缓存，确保所有内容写入磁盘
                    tmp.seek(0)  # 重置文件指针，方便后续读取
                    fp = tmp.name

            # 验证临时文件是否完整
            tmp_size = os.path.getsize(fp) / (1024 * 1024)
            if tmp_size < upload_size * 0.9:  # 若临时文件小于上传文件的90%，判定为保存失败
                raise Exception(f"Temporary file incomplete (size: {tmp_size:.2f} MB < {upload_size:.2f} MB)")
            log.append(f"✅ Temporary file saved (size: {tmp_size:.2f} MB)")
            log_area.markdown("\n".join(log))

            # ========== 修复2：完整读取DOCX，验证页数和表格数 ==========
            doc = Document(fp)
            # 新增：打印文档基本信息，确认是否读取到多页内容
            total_tables = len(doc.tables)
            log.append(f"✅ DOCX parsed: {len(doc.paragraphs)} paragraphs | {total_tables} tables")
            log_area.markdown("\n".join(log))

            # ========== 修复3：完整提取段落（含跨页段落） ==========
            # 逻辑：遍历所有段落，包括跨页的连续段落
            for para_idx, para in enumerate(doc.paragraphs):
                text = para.text.strip()
                if text and len(text) >= 1:
                    text_items.append((para, text, 'paragraph'))
                    all_texts.append(text)
                    para_count += 1
                    # 每提取100个段落打印日志，确认是否覆盖多页
                    if para_count % 100 == 0:
                        log.append(f"🔍 Extracted {para_count} paragraphs (current para index: {para_idx})")
                        log_area.markdown("\n".join(log))

            # ========== 修复4：完整提取表格（含跨页表格） ==========
            # 逻辑：遍历所有表格，逐行逐 cell 提取，不遗漏跨页单元格
            for table_idx, table in enumerate(doc.tables):
                log.append(f"🔍 Processing table {table_idx + 1}/{total_tables} (rows: {len(table.rows)})")
                log_area.markdown("\n".join(log))
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        # 关键：用 cell.text 提取完整文本，包括跨页单元格
                        cell_text = cell.text.strip()
                        if cell_text:
                            text_items.append((cell, cell_text, 'cell'))
                            all_texts.append(cell_text)
                            cell_count += 1
                            # 每提取50个单元格打印日志，确认是否覆盖跨页表格
                            if cell_count % 50 == 0:
                                log.append(f"🔍 Extracted {cell_count} table cells (table {table_idx+1}, row {row_idx+1})")
                                log_area.markdown("\n".join(log))

            # 检查提取结果
            total = len(all_texts)
            if total == 0:
                st.error("❌ No valid text found")
                st.stop()
            log.append(f"✅ Text extraction completed: {para_count} paragraphs + {cell_count} cells = {total} segments")
            log.append(f"🔤 Direction: {source_lang_name} → {target_lang_name} | ⚙️ {wk} threads | {BS} segments/batch")
            log_area.markdown("\n".join(log))

            # ========== 修复5：多线程翻译，避免批次错位 ==========
            translations = [None] * total
            def translate_batch(txt_list):
                try:
                    # 卡车术语预处理（适配你的文档）
                    term_map = {
                        "GVW": "Gross Vehicle Weight",
                        "Curb Weight": "Curb Weight of Chassis",
                        "Axle Load": "Axle Load Distribution",
                        "Wheel Base": "Wheel Base Length",
                        "Max Torque": "Maximum Torque",
                        "Fuel Consumption": "Fuel Consumption per 100km",
                        "LHD": "Left-Hand Drive",
                        "ABS": "Anti-lock Braking System"
                    }
                    # 替换术语
                    processed = [txt for txt in txt_list]
                    for i, txt in enumerate(processed):
                        for term, full in term_map.items():
                            if term in txt:
                                processed[i] = txt.replace(term, full)
                    # 翻译
                    res = GoogleTranslator(source=source_lang, target=target_lang).translate_batch(processed)
                    # 兜底：确保结果长度与输入一致
                    if len(res) != len(processed):
                        st.warning(f"⚠️ Batch result length mismatch (input: {len(processed)}, output: {len(res)})")
                        return processed
                    # 还原术语
                    final = []
                    for r in res:
                        if r is None or r.strip() == "":
                            final.append(processed[len(final)])  # 用原文兜底
                        else:
                            for term, full in term_map.items():
                                if full in r:
                                    r = r.replace(full, term)
                            final.append(r)
                    return final
                except Exception as e:
                    st.warning(f"⚠️ Batch error: {str(e)[:30]}...")
                    return txt_list

            # 提交多线程任务（修复：确保批次索引不越界）
            with ThreadPoolExecutor(max_workers=wk) as exe:
                futures = {}
                for start_idx in range(0, total, BS):
                    end_idx = min(start_idx + BS, total)
                    batch = all_texts[start_idx:end_idx]
                    fut = exe.submit(translate_batch, batch)
                    futures[fut] = (start_idx, end_idx)  # 记录批次的起始和结束索引

                # 处理结果（修复：按批次范围赋值，避免错位）
                for fut in as_completed(futures):
                    start_idx, end_idx = futures[fut]
                    batch_res = fut.result()
                    # 确保批次结果长度与批次范围一致
                    batch_len = end_idx - start_idx
                    if len(batch_res) != batch_len:
                        batch_res = batch_res[:batch_len] + [all_texts[start_idx + i] for i in range(len(batch_res), batch_len)]
                    # 赋值到翻译结果列表
                    for i in range(batch_len):
                        translations[start_idx + i] = batch_res[i]
                    # 进度日志
                    done = sum(1 for x in translations if x is not None)
                    if done % 20 == 0 or done == total:
                        log.append(f"🔄 Translated: {done}/{total} segments")
                        log_area.markdown("\n".join(log))

            log.append(f"✅ Translation completed: {total}/{total} segments")
            log_area.markdown("\n".join(log))

            # ========== 修复6：完整回写，确保跨页内容更新 ==========
            log.append("📝 Updating document (including cross-page content)...")
            log_area.markdown("\n".join(log))
            update_count = 0
            for idx, (obj, original, obj_type) in enumerate(text_items):
                if idx >= len(translations):
                    continue  # 避免索引越界
                trans = translations[idx]
                if trans and trans != original:
                    if obj_type == 'paragraph':
                        obj.text = trans
                    elif obj_type == 'cell':
                        obj.text = trans
                    update_count += 1

            log.append(f"✅ Document updated: {update_count}/{len(text_items)} segments modified")
            log_area.markdown("\n".join(log))

            # 保存结果
            output_path = fp.replace(".docx", f"_{source_lang_name}2{target_lang_name}.docx")
            doc.save(output_path)
            tot_t = time.time() - stt

            # 成功提示
            st.balloons()
            st.success(f"### ✅ Translation Completed!（{source_lang_name} → {target_lang_name}）")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Time", f"{tot_t:.1f}s")
            c2.metric("Total Segments", f"{total}")
            c3.metric("Updated Segments", f"{update_count}")

            # 下载
            st.markdown("---")
            with open(output_path, "rb") as f:
                st.download_button(
                    "📥 Download Translated Document",
                    f,
                    file_name=f"{source_lang_name}2{target_lang_name}_{uf.name}",
                    use_container_width=True,
                    type="primary"
                )

        except Exception as e:
            st.error("### ❌ Translation Failed")
            st.exception(e)
            log.append(f"❌ Error: {str(e)[:60]}...")
            log_area.markdown("\n".join(log))

        finally:
            # 清理临时文件
            try:
                if fp and os.path.exists(fp):
                    os.unlink(fp)
                if output_path and os.path.exists(output_path):
                    os.unlink(output_path)
                log.append("✅ Temporary files cleaned up")
                log_area.markdown("\n".join(log))
            except Exception as cleanup_e:
                st.warning(f"⚠️ Cleanup warning: {str(cleanup_e)[:50]}...")
