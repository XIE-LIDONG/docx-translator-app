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
    c1, c2, c3, c4 = st.columns(4)  # 新增一列放批次选择
    with c1:
        source_lang_name = st.selectbox(
            "**Source Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("French")  # Default source: French
        )
        # Convert to deep-translator code
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox(
            "**Target Language**",
            LANG_NAMES,
            index=LANG_NAMES.index("English")  # Default target: English
        )
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        # Thread count limited to 1-3, 默认值改为2
        wk = st.slider(
            "**Thread Count**",
            min_value=1,
            max_value=3,
            value=2,  # 默认线程数2
            help="Number of parallel translation threads (1-3 for stability)"
        )
    with c4:
        # 新增批次选择器：20-100，默认100
        BS = st.slider(
            "**Batch Size**",
            min_value=20,
            max_value=100,
            value=100,  # 默认批次大小100
            step=10,  # 步长10，方便调整
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
            ti, at = [], []  # text_items, all_texts
            pc, cc = 0, 0    # paragraph/table count

            # Extract text
            for p in doc.paragraphs:
                if txt := p.text.strip():
                    ti.append((p, txt))
                    at.append(txt)
                    pc += 1
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs:
                            if txt := p.text.strip():
                                ti.append((p, txt))
                                at.append(txt)
                                cc += 1

            total = len(at)
            if total == 0:
                st.error("❌ No valid text in document")
                st.stop()

            # Initial log (language + 线程/批次信息)
            log.append(f"✅ Extracted {total} text segments for translation")
            log.append(f"🔤 Translation direction: {source_lang_name} → {target_lang_name}")
            log.append(f"⚙️ Configuration: {wk} threads | {BS} segments per batch")  # 新增配置日志
            log_area.markdown("\n".join(log))

            # Multi-thread translation
            ta = [None]*total  # translation results
            def tb(txts):  # batch translation function
                return GoogleTranslator(source=source_lang, target=target_lang).translate_batch(txts)

            with ThreadPoolExecutor(max_workers=wk) as exe:
                # Submit batch tasks (使用用户选择的BS值)
                futs = {}
                for i in range(0, total, BS):
                    batch = at[i:i+BS]
                    fut = exe.submit(tb, batch)
                    futs[fut] = i  # record batch start index

                # Process results + real-time log
                for fut in as_completed(futs):
                    start_idx = futs[fut]
                    res = fut.result()
                    # Save results
                    for idx in range(len(res)):
                        if start_idx+idx < total:
                            ta[start_idx+idx] = res[idx]
                    # Calculate completed count
                    done = sum(1 for x in ta if x is not None)
                    # Log every 10 segments (保留原有日志逻辑)
                    if done % 10 == 0:
                        log.append(f"🔄 Translating: {done}/{total}")
                        log_area.markdown("\n".join(log))

            # Final translation completion log
            log.append(f"✅ Translation completed: {total}/{total}")
            log_area.markdown("\n".join(log))

            # Update document
            log.append("📝 Updating document...")
            log_area.markdown("\n".join(log))
            for idx, (p_obj, _) in enumerate(ti):
                if ta[idx]:
                    p_obj.text = ta[idx]

            # Save for download
            op = fp.replace(".docx", "_translated.docx")
            doc.save(op)
            tot_t = time.time()-stt

            st.balloons()
            st.success(f"### ✅ Translation Completed!（{source_lang_name} → {target_lang_name}）")
            # Statistics
            c1,c2 = st.columns(2)  # Compact layout
            c1.metric("Total Time", f"{tot_t:.1f}s")
            c2.metric("Average Speed", f"{total/tot_t:.1f} segments/s")

            # Download button
            st.markdown("---")
            with open(op, "rb") as f:
                st.download_button(
                    "📥 Download Translated Document", f,
                    file_name=f"{source_lang_name}2{target_lang_name}_{uf.name}",  # Filename with translation direction
                    use_container_width=True, type="primary"
                )

        except Exception as e:
            st.error("### ❌ Translation Failed")
            st.exception(e)
        finally:
            # Clean up temporary files
            try:
                os.unlink(fp)
                if 'op' in locals() and os.path.exists(op):
                    os.unlink(op)
            except Exception as cleanup_e:
                st.warning(f"⚠️ Temporary file cleanup failed: {cleanup_e}")



