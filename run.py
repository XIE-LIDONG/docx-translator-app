import streamlit as st
import tempfile
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import os

# 1. 页面基础配置
st.set_page_config(page_title="DOCX Document Language Translator", page_icon="📄", layout="wide")

# 标题+署名：左右布局
st.markdown(
    """
    <div style='display: flex; justify-content: space-between; align-items: center;'>
        <h1 style='margin: 0;'>📄 DOCX Document Language Translator</h1>
        <p style='color: #666666; font-size: 14px; margin: 0;'>By XIE LI DONG</p>
    </div>
    """,
    unsafe_allow_html=True
)

# 使用提示
st.info("""
💡 **Usage Tip**: If your file is in PDF format, convert it to DOCX first via [ILovePDF](https://www.ilovepdf.com/). 
After translation, you can convert the DOCX back to PDF using ILovePDF if needed. 
PDF-DOCX direct integration is skipped to avoid app slowdown.
""")
st.markdown("---")

# 2. 支持语言配置（适配卡车文档的中英文翻译）
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

# 3. 文件上传模块
uf = st.file_uploader("Select Word Document (.docx)", type=["docx"])

if uf:
    # 显示文件信息
    c1, c2 = st.columns([2, 1])
    with c1:
        st.success(f"📁 **File**: {uf.name}")
    with c2:
        st.metric("Size", f"{uf.size/(1024*1024):.2f} MB")
    st.markdown("---")

    # 4. 翻译参数配置（线程、批次、语言选择）
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        source_lang_name = st.selectbox(
            "**Source Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("English")  # 默认源语言：英文（适配卡车文档）
        )
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox(
            "**Target Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("Chinese")  # 默认目标语言：中文
        )
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        wk = st.slider(
            "**Thread Count**",
            min_value=1,
            max_value=3,
            value=2,
            help="Number of parallel threads (1-3 for stability)"
        )
    with c4:
        BS = st.slider(
            "**Batch Size**",
            min_value=20,
            max_value=100,
            value=50,  # 适配参数文档：50段/批次，平衡速度与稳定性
            step=10,
            help="Text segments per batch (20-100)"
        )

    # 5. 核心翻译逻辑（点击按钮触发）
    if st.button("🚀 Start Translation", type="primary", use_container_width=True):
        # 进度日志区域
        log_area = st.empty()
        log = []
        # 初始化关键变量（避免未定义错误）
        doc = None
        fp = None
        output_path = None
        text_items = []  # 存储 (对象, 原文, 类型)
        all_texts = []   # 存储所有待翻译文本
        translations = []  # 存储翻译结果
        para_count = 0
        cell_count = 0

        try:
            stt = time.time()  # 计时开始

            # 5.1 创建临时文件（存储上传的DOCX）
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(uf.getvalue())
                fp = tmp.name  # 临时文件路径
            log.append(f"✅ Temporary file created: {os.path.basename(fp)}")
            log_area.markdown("\n".join(log))

            # 5.2 加载DOCX文档
            doc = Document(fp)
            log.append("✅ DOCX file parsed successfully")
            log_area.markdown("\n".join(log))

            # ========== 5.3 文本提取模块（完整提取段落+表格） ==========
            # 提取【普通段落】文本
            for para in doc.paragraphs:
                text = para.text.strip()
                if text and len(text) >= 1:  # 仅过滤空文本（保留参数短文本）
                    text_items.append((para, text, 'paragraph'))
                    all_texts.append(text)
                    para_count += 1

            # 提取【表格单元格】文本（适配卡车参数表）
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_full_text = cell.text.strip()
                        if cell_full_text:
                            # 保留单元格完整文本（如"GVW (kg): 41T"），避免碎片化
                            text_items.append((cell, cell_full_text, 'cell'))
                            all_texts.append(cell_full_text)
                            cell_count += 1

            # 检查是否有可翻译文本
            total = len(all_texts)
            if total == 0:
                st.error("❌ No valid text found in document")
                st.stop()
            log.append(f"✅ Text extraction completed: {para_count} paragraphs + {cell_count} table cells = {total} segments")
            log.append(f"🔤 Translation direction: {source_lang_name} → {target_lang_name}")
            log.append(f"⚙️ Configuration: {wk} threads | {BS} segments/batch")
            log_area.markdown("\n".join(log))

            # ========== 5.4 翻译函数（增强专业术语容错） ==========
            def translate_batch(txt_list):
                try:
                    # 术语预处理：替换卡车领域缩写（提升翻译准确性）
                    term_map = {
                        "GVW": "Gross Vehicle Weight",
                        "Curb Weight": "Curb Weight of Chassis",
                        "Axle Load": "Axle Load Distribution",
                        "Wheel Base": "Wheel Base Length",
                        "Max Torque": "Maximum Torque",
                        "Fuel Consumption": "Fuel Consumption per 100km",
                        "Nm": "Newton-meters",
                        "LHD": "Left-Hand Drive",
                        "ABS": "Anti-lock Braking System"
                    }
                    # 替换缩写为完整术语
                    processed_txt = []
                    for txt in txt_list:
                        processed = txt
                        for term, full_term in term_map.items():
                            if term in processed:
                                processed = processed.replace(term, full_term)
                        processed_txt.append(processed)

                    # 调用Google翻译（批量）
                    res = GoogleTranslator(source=source_lang, target=target_lang).translate_batch(processed_txt)

                    # 非空兜底：翻译失败时返回原文
                    translated = []
                    for r, txt in zip(res, processed_txt):
                        if r is not None and r.strip() != "":
                            translated.append(r)
                        else:
                            translated.append(txt)

                    # 还原术语缩写（如"Gross Vehicle Weight"→"GVW"）
                    final_res = []
                    for txt in translated:
                        final = txt
                        for term, full_term in term_map.items():
                            if full_term in final:
                                final = final.replace(full_term, term)
                        final_res.append(final)

                    return final_res
                except Exception as e:
                    # 单个批次报错时返回原文，不中断整体流程
                    st.warning(f"⚠️ Batch translation warning: {str(e)[:40]}...")
                    return txt_list

            # ========== 5.5 多线程批量翻译 ==========
            translations = [None] * total  # 初始化翻译结果列表
            with ThreadPoolExecutor(max_workers=wk) as exe:
                # 提交批次任务
                futures = {}
                for start_idx in range(0, total, BS):
                    end_idx = min(start_idx + BS, total)  # 避免索引越界
                    batch_txt = all_texts[start_idx:end_idx]
                    fut = exe.submit(translate_batch, batch_txt)
                    futures[fut] = start_idx  # 记录批次起始位置

                # 处理翻译结果
                for fut in as_completed(futures):
                    start_idx = futures[fut]
                    batch_res = fut.result()
                    # 赋值到结果列表
                    for idx in range(len(batch_res)):
                        if start_idx + idx < total:
                            translations[start_idx + idx] = batch_res[idx]
                    # 日志更新（每完成20段显示一次）
                    done = sum(1 for x in translations if x is not None)
                    if done % 20 == 0 or done == total:
                        log.append(f"🔄 Translating: {done}/{total} segments completed")
                        log_area.markdown("\n".join(log))

            log.append(f"✅ All translation completed: {total}/{total} segments")
            log_area.markdown("\n".join(log))

            # ========== 5.6 文本回写模块（段落+表格分别回写） ==========
            log.append("📝 Updating document content (including tables)...")
            log_area.markdown("\n".join(log))
            for idx, (obj, original_text, obj_type) in enumerate(text_items):
                trans_text = translations[idx]
                # 仅当翻译结果有效且与原文不同时回写
                if trans_text and trans_text != original_text:
                    if obj_type == 'paragraph':
                        obj.text = trans_text  # 段落回写
                    elif obj_type == 'cell':
                        obj.text = trans_text  # 表格单元格回写（关键修复）

            # 5.7 保存翻译后的文档
            output_path = fp.replace(".docx", f"_{source_lang_name}2{target_lang_name}.docx")
            doc.save(output_path)
            tot_t = time.time() - stt  # 计时结束

            # 5.8 显示翻译完成信息
            st.balloons()
            st.success(f"### ✅ Translation Completed!（{source_lang_name} → {target_lang_name}）")
            # 统计信息
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Time", f"{tot_t:.1f}s")
            c1.metric("Total Segments", f"{total}")
            c2.metric("Paragraphs Translated", f"{para_count}")
            c2.metric("Table Cells Translated", f"{cell_count}")
            c3.metric("Average Speed", f"{total/tot_t:.1f} segments/s")

            # 5.9 下载按钮
            st.markdown("---")
            with open(output_path, "rb") as f:
                st.download_button(
                    "📥 Download Translated Document",
                    f,
                    file_name=f"{source_lang_name}2{target_lang_name}_{uf.name}",
                    use_container_width=True,
                    type="primary"
                )

        # 5.10 异常捕获（避免程序崩溃）
        except Exception as e:
            st.error("### ❌ Translation Failed")
            st.exception(e)
            log.append(f"❌ Error: {str(e)[:50]}...")
            log_area.markdown("\n".join(log))

        # 5.11 清理临时文件（避免残留）
        finally:
            try:
                # 删除临时文件
                if fp and os.path.exists(fp):
                    os.unlink(fp)
                if output_path and os.path.exists(output_path):
                    os.unlink(output_path)
                log.append("✅ Temporary files cleaned up")
                log_area.markdown("\n".join(log))
            except Exception as cleanup_e:
                st.warning(f"⚠️ Temporary file cleanup warning: {str(cleanup_e)[:50]}...")
