# ========== 1. 文本提取模块：完整提取表格文本 ==========
# 提取【普通段落】文本（保留原逻辑）
for para in doc.paragraphs:
    text = para.text.strip()
    # 修复：取消短文本过滤，确保所有非空文本都被提取
    if text and len(text) >= 1:  # 仅过滤纯空文本
        text_items.append((para, text, 'paragraph'))
        all_texts.append(text)
        para_count += 1

# 提取【表格单元格】文本：改用 cell.text 提取完整文本，再处理多段落
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            # 关键修复：用 cell.text 获取单元格完整文本（含所有段落）
            cell_full_text = cell.text.strip()
            # 处理单元格内多段落（用换行符分隔）
            if cell_full_text:
                # 将单元格完整文本拆分为单个文本段（避免多段落导致翻译碎片化）
                text_items.append((cell, cell_full_text, 'cell'))
                all_texts.append(cell_full_text)
                cell_count += 1

# ========== 2. 翻译函数优化：增强专业术语翻译容错 ==========
def translate_batch(txt_list):
    try:
        # 术语预处理：将专业缩写替换为完整英文（提升翻译准确性）
        processed_txt = []
        term_map = {
            "GVW": "Gross Vehicle Weight",
            "Curb Weight": "Curb Weight of Chassis",
            "Axle Load": "Axle Load Distribution",
            "Wheel Base": "Wheel Base Length",
            "Max Torque": "Maximum Torque",
            "Fuel Consumption": "Fuel Consumption per 100km"
        }
        for txt in txt_list:
            # 替换专业术语缩写
            for term, full_term in term_map.items():
                if term in txt:
                    txt = txt.replace(term, full_term)
            processed_txt.append(txt)
        
        # 调用翻译接口
        res = GoogleTranslator(source=source_lang, target=target_lang).translate_batch(processed_txt)
        # 非空兜底：若翻译结果为空，返回处理后的原文（而非原始原文）
        translated = [r if (r is not None and r.strip() != "") else txt for r, txt in zip(res, processed_txt)]
        # 还原术语缩写（避免翻译后术语不一致）
        final_res = []
        for txt in translated:
            for term, full_term in term_map.items():
                if full_term in txt:
                    txt = txt.replace(full_term, term)
            final_res.append(txt)
        return final_res
    except Exception as e:
        st.warning(f"⚠️ Batch translation warning: {str(e)[:30]}...")
        return txt_list  # 报错时返回原文

# ========== 3. 文本回写模块：适配单元格完整文本 ==========
log.append("📝 Updating document content (including tables)...")
log_area.markdown("\n".join(log))
for idx, (obj, original_text, obj_type) in enumerate(text_items):
    trans_text = translations[idx]
    if trans_text and trans_text != original_text:
        if obj_type == 'paragraph':
            obj.text = trans_text  # 段落回写
        elif obj_type == 'cell':
            # 关键修复：直接修改单元格完整文本（覆盖所有段落）
            obj.text = trans_text  # 单元格回写：cell.text 是可写属性（之前的认知错误！）
