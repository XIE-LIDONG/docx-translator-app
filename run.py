
import streamlit as st
import tempfile
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import os

# 页面配置
st.set_page_config(page_title="DOCX翻译器", page_icon="📄", layout="wide")
st.title("🚀 DOCX文档翻译器（多语言互译）")
st.markdown("---")

# 定义支持的语言（名称: deep-translator对应代码）
SUPPORT_LANGUAGES = {
    "中文": "zh-CN",
    "英语": "en",
    "法语": "fr",
    "德语": "de",
    "西班牙语": "es",
    "阿拉伯语": "ar",
    "日语": "ja",
    "韩语": "ko"
}
# 提取语言名称列表（用于下拉框）
LANG_NAMES = list(SUPPORT_LANGUAGES.keys())

# 上传文件
uf = st.file_uploader("选择Word文档 (.docx)", type=["docx"])

if uf:
    # 文件信息
    c1, c2 = st.columns([2,1])
    with c1: st.success(f"📁 **文件:** {uf.name}")
    with c2: st.metric("大小", f"{uf.size/(1024*1024):.2f} MB")
    st.markdown("---")

    # 翻译设置（多语言互译下拉框）
    c1, c2, c3 = st.columns(3)
    with c1:
        source_lang_name = st.selectbox(
            "**源语言**",
            LANG_NAMES,
            index=LANG_NAMES.index("法语")  # 默认源语言为法语
        )
        # 转换为deep-translator识别的代码
        source_lang = SUPPORT_LANGUAGES[source_lang_name]
    with c2:
        target_lang_name = st.selectbox(
            "**目标语言**",
            LANG_NAMES,
            index=LANG_NAMES.index("英语")  # 默认目标语言为英语
        )
        target_lang = SUPPORT_LANGUAGES[target_lang_name]
    with c3:
        wk = st.slider("**线程数**", 1, 8, 3)

    if st.button("🚀 开始翻译", type="primary", use_container_width=True):
        # 用文本区域显示进度日志
        log_area = st.empty()
        log = []

        # 临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uf.getvalue())
            fp = tmp.name

        try:
            stt = time.time()
            # 分析文档
            doc = Document(fp)
            ti, at = [], []  # text_items, all_texts
            pc, cc = 0, 0    # 段落/表格计数

            # 提取文本
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
                st.error("❌ 文档无有效文本")
                st.stop()

            # 初始日志（显示语言信息）
            log.append(f"✅ 共提取 {total} 段待翻译文字")
            log.append(f"🔤 翻译方向: {source_lang_name} → {target_lang_name}")
            log_area.markdown("\n".join(log))

            # 多线程翻译
            ta = [None]*total  # 翻译结果
            BS = 100  # 批次大小
            def tb(txts):  # 批次翻译函数
                return GoogleTranslator(source=source_lang, target=target_lang).translate_batch(txts)

            with ThreadPoolExecutor(max_workers=wk) as exe:
                # 提交批次任务
                futs = {}
                for i in range(0, total, BS):
                    batch = at[i:i+BS]
                    fut = exe.submit(tb, batch)
                    futs[fut] = i  # 记录批次起始索引

                # 处理结果+实时打日志
                for fut in as_completed(futs):
                    start_idx = futs[fut]
                    res = fut.result()
                    # 保存结果
                    for idx in range(len(res)):
                        if start_idx+idx < total:
                            ta[start_idx+idx] = res[idx]
                    # 计算已翻译数量
                    done = sum(1 for x in ta if x is not None)
                    # 每翻译10段打一次日志
                    if done % 10 == 0:
                        log.append(f"🔄 翻译中: {done}/{total}")
                        log_area.markdown("\n".join(log))

            # 最终翻译完成日志
            log.append(f"✅ 翻译完成: {total}/{total}")
            log_area.markdown("\n".join(log))

            # 更新文档
            log.append("📝 更新文档中...")
            log_area.markdown("\n".join(log))
            for idx, (p_obj, _) in enumerate(ti):
                if ta[idx]:
                    p_obj.text = ta[idx]

            # 保存下载
            op = fp.replace(".docx", "_translated.docx")
            doc.save(op)
            tot_t = time.time()-stt

            st.balloons()
            st.success(f"### ✅ 翻译完成！（{source_lang_name} → {target_lang_name}）")
            # 统计信息
            c1,c2 = st.columns(2)  # 调整列数更紧凑
            c1.metric("总耗时", f"{tot_t:.1f}秒")
            c2.metric("平均速度", f"{total/tot_t:.1f}条/秒")

            # 下载按钮
            st.markdown("---")
            with open(op, "rb") as f:
                st.download_button(
                    "📥 下载翻译文档", f,
                    file_name=f"{source_lang_name}2{target_lang_name}_{uf.name}",  # 文件名带翻译方向
                    use_container_width=True, type="primary"
                )

        except Exception as e:
            st.error("### ❌ 翻译失败")
            st.exception(e)
        finally:
            # 清理临时文件
            try:
                os.unlink(fp)
                if 'op' in locals() and os.path.exists(op):
                    os.unlink(op)
            except Exception as cleanup_e:
                st.warning(f"⚠️ 临时文件清理失败: {cleanup_e}")