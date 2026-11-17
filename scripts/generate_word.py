import sys
import os
import re
import json
import io
import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("generate_word")

# Import citeproc-py components
try:
    from citeproc import CitationStylesStyle, CitationStylesBibliography
    from citeproc import Citation, CitationItem
    from citeproc import formatter
    from citeproc.source.json import CiteProcJSON
    CITEPROC_AVAILABLE = True
except ImportError:
    CITEPROC_AVAILABLE = False


def apply_hardcoded_formatting(paragraph, style_name='Normal'):
    paragraph.paragraph_format.line_spacing = 1.25
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    for run in paragraph.runs:
        run.font.name = 'Times New Roman'
        if style_name in ['Normal', 'Bình thường', 'Normal (Web)', 'Heading 1', 'Heading 2', 'Heading 3']:
            if style_name in ['Normal', 'Bình thường']:
                run.font.size = Pt(14)
            
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'))
            rPr.append(rFonts)
            
        rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')


def add_paragraph_with_inline_bold(doc, text, style_name):
    """Tạo đoạn văn và phân tích cú pháp **text** thành run in đậm."""
    p = doc.add_paragraph(style=style_name)
    # Tách văn bản theo cú pháp **...**
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = p.add_run(part[2:-2])  # Bỏ ** ở đầu và cuối
            run.bold = True
        else:
            p.add_run(part)
    apply_hardcoded_formatting(p, style_name)
    return p



def set_cell_shading(cell, color_hex):
    """Đổ bóng nền cho ô bảng."""
    try:
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shd)
    except Exception as e:
        logger.warning(f"Không thể đổ bóng ô: {e}")


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Đặt khoảng đệm (padding) cho ô bảng bằng dxa (1 pt = 20 dxa)."""
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m_name, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
            node = OxmlElement(m_name)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)
    except Exception as e:
        logger.warning(f"Không thể đặt cell margins: {e}")


def set_table_borders(table):
    """Thiết lập viền lưới mỏng cho bảng bằng XML theo phong cách học thuật hiện đại."""
    try:
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="888888"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)
    except Exception as e:
        logger.warning(f"Không thể áp dụng viền bảng: {e}")


def add_styled_table(doc, rows_data):
    """Tạo bảng và áp dụng định dạng Times New Roman cỡ 11, shading tiêu đề."""
    if not rows_data:
        return
    num_cols = len(rows_data[0])
    table = doc.add_table(rows=0, cols=num_cols)
    set_table_borders(table)
    
    for row_idx, row_cells in enumerate(rows_data):
        row = table.add_row()
        for col_idx, cell_text in enumerate(row_cells):
            # Nếu số ô trong dòng ít hơn số cột, ta dùng chỉ mục an toàn
            if col_idx >= len(row.cells):
                break
            cell = row.cells[col_idx]
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            
            # Phân tích cú pháp inline bold **...**
            parts = re.split(r'(\*\*.*?\*\*)', cell_text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    run = p.add_run(part)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = rPr.makeelement(qn('w:rFonts'))
                    rPr.append(rFonts)
                rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                rFonts.set(qn('w:cs'), 'Times New Roman')
            
            # Định dạng tiêu đề bảng
            if row_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.bold = True
                set_cell_shading(cell, 'F2F2F2')
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Đệm khoảng cách ô (5pt top/bottom, 7.5pt left/right)
            set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
            
    # Thêm dòng trống sau bảng
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(6)


def generate_manuscript_docx(input_text_file, output_docx_file, template_file=None, references_db_file=None):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        skill_root = os.path.dirname(script_dir)
        
        if not template_file:
            template_file = os.path.join(skill_root, "references", "manuscript_template.docx")
        csl_style_file = os.path.join(skill_root, "references", "vietnam-ministry-of-education-and-training-en.csl")

        workspace_dir = os.getcwd()
        if not os.path.isabs(output_docx_file):
            output_docx_file = os.path.join(workspace_dir, output_docx_file)

        if os.path.exists(template_file):
            doc = Document(template_file)
        else:
            doc = Document()

        # 4. Khởi tạo CiteProc
        bibliography = None
        bib_source = {}
        if not references_db_file:
            references_db_file = os.path.join(workspace_dir, "references_db.json")
            
        if CITEPROC_AVAILABLE and os.path.exists(references_db_file) and os.path.exists(csl_style_file):
            with open(references_db_file, "r", encoding="utf-8") as f:
                db_data = json.load(f)
            if db_data:
                bib_source = CiteProcJSON(db_data)
                bib_style = CitationStylesStyle(csl_style_file, validate=False)
                bibliography = CitationStylesBibliography(bib_style, bib_source, formatter.plain)

        # 5. Đọc dữ liệu đầu vào
        with open(input_text_file, 'r', encoding='utf-8') as f:
            content = f.readlines()

        anchor_pattern = re.compile(r'\[@PMID:\s*([^\]]+)\]', re.IGNORECASE)
        
        # --- ĐĂNG KÝ TRÍCH DẪN VÀ TẠO BẢN ĐỒ ---
        pmid_to_num = {}
        registered_keys = []
        
        if bibliography:
            full_text = "".join(content)
            all_anchors = anchor_pattern.findall(full_text)
            for anchor_content in all_anchors:
                # Tách các PMID
                pmids = [p.strip().lower() for p in anchor_content.split(',')]
                for pmid_val in pmids:
                    # Chuyển thành lowercase pmid_... vì citeproc-py tự động lowercase keys
                    citation_key = pmid_val
                    if citation_key not in bib_source and f"pmid_{pmid_val}" in bib_source:
                        citation_key = f"pmid_{pmid_val}"
                    if citation_key in bib_source and citation_key not in registered_keys:
                        citation = Citation([CitationItem(citation_key)])
                        bibliography.register(citation)
                        registered_keys.append(citation_key)
            
            # Sau khi đăng ký xong, tạo bản đồ số thứ tự
            for i, key in enumerate(registered_keys):
                pmid_only = key.replace("pmid_", "")
                pmid_to_num[pmid_only] = f"[{i+1}]"

        # 6. Xử lý văn bản và xuất Word
        # Heading patterns
        h1_pattern = re.compile(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+\S.*')
        h2_pattern = re.compile(r'^[1-9]\d{0,1}\.\s+\S.*')
        h3_pattern = re.compile(r'^[a-z]\.\s+\S.*')

        is_first_line = True
        available_styles = [s.name for s in doc.styles if s.type == 1]

        def get_best_style(target_names):
            for name in target_names:
                if name in available_styles:
                    return name
            return 'Normal'

        # Thay thế mỏ neo trên tất cả các dòng trước khi phân chia khối
        processed_lines = []
        for line in content:
            line = line.strip()
            if not line:
                processed_lines.append("")
                continue
            
            # Thay thế mỏ neo
            matches = list(anchor_pattern.finditer(line))
            for match in reversed(matches):
                pmid_raw = match.group(1).strip()
                pmid_parts = [p.strip().lower() for p in pmid_raw.split(',')]
                
                nums = []
                for p in pmid_parts:
                    if p in pmid_to_num:
                        nums.append(pmid_to_num[p].strip('[]'))
                
                if nums:
                    replacement = "[" + ", ".join(nums) + "]"
                    line = line[:match.start()] + replacement + line[match.end():]
            
            # Loại bỏ các chú thích HTML (như <!--ref:...--> và <!--anchor:...-->)
            line = re.sub(r'<!--.*?-->', '', line)
            
            processed_lines.append(line)

        # Phân nhóm thành các khối text hoặc bảng
        blocks = []
        current_table = []
        
        for line in processed_lines:
            if not line:
                if current_table:
                    blocks.append(("table", current_table))
                    current_table = []
                continue
                
            if line.startswith('|') and line.endswith('|'):
                # Dòng của bảng
                cells = [c.strip() for c in line.split('|')[1:-1]]
                # Kiểm tra xem có phải dòng gạch ngang phân cách |---|---| không
                is_separator = all(re.match(r'^:?-+:?$', c) for c in cells) if cells else False
                if is_separator:
                    continue
                current_table.append(cells)
            else:
                if current_table:
                    blocks.append(("table", current_table))
                    current_table = []
                blocks.append(("text", line))
                
        if current_table:
            blocks.append(("table", current_table))

        # Xuất các khối ra Word
        for b_type, b_val in blocks:
            if b_type == "text":
                if is_first_line:
                    style_name = get_best_style(['tiêu đề', 'TIÊU ĐỀ', 'Title'])
                elif h1_pattern.match(b_val):
                    style_name = get_best_style(['Heading 1', 'H1'])
                elif h2_pattern.match(b_val):
                    style_name = get_best_style(['Heading 2', 'H2'])
                elif h3_pattern.match(b_val):
                    style_name = get_best_style(['Heading 3', 'H3'])
                else:
                    style_name = get_best_style(['Bình thường', 'Normal'])
                
                is_first_line = False
                add_paragraph_with_inline_bold(doc, b_val, style_name)
            elif b_type == "table":
                add_styled_table(doc, b_val)

        # 7. Xuất Danh mục Tài liệu tham khảo
        if bibliography and registered_keys:
            h1_style = get_best_style(['Heading 1', 'H1'])
            p_ref_title = doc.add_paragraph("TÀI LIỆU THAM KHẢO", style=h1_style)
            apply_hardcoded_formatting(p_ref_title, h1_style)
            
            norm_style = get_best_style(['Bình thường', 'Normal'])
            bib_items = bibliography.bibliography()
            
            def add_custom_numbering(doc):
                numbering_part = doc.part.numbering_part
                if numbering_part is None:
                    return None
                
                numbering = numbering_part.numbering_definitions._numbering
                existing_abs_ids = set()
                existing_num_ids = set()
                for elem in numbering.findall(qn('w:abstractNum')):
                    aid = elem.get(qn('w:abstractNumId'))
                    if aid:
                        existing_abs_ids.add(int(aid))
                for elem in numbering.findall(qn('w:num')):
                    nid = elem.get(qn('w:numId'))
                    if nid:
                        existing_num_ids.add(int(nid))
                
                abs_num_id = max(existing_abs_ids, default=0) + 1
                num_id = max(existing_num_ids, default=0) + 1
                
                abstractNum = OxmlElement('w:abstractNum')
                abstractNum.set(qn('w:abstractNumId'), str(abs_num_id))
                multiLevelType = OxmlElement('w:multiLevelType')
                multiLevelType.set(qn('w:val'), 'hybridMultilevel')
                abstractNum.append(multiLevelType)
                lvl = OxmlElement('w:lvl')
                lvl.set(qn('w:ilvl'), '0')
                start = OxmlElement('w:start')
                start.set(qn('w:val'), '1')
                lvl.append(start)
                numFmt = OxmlElement('w:numFmt')
                numFmt.set(qn('w:val'), 'decimal')
                lvl.append(numFmt)
                lvlText = OxmlElement('w:lvlText')
                lvlText.set(qn('w:val'), '%1.')
                lvl.append(lvlText)
                lvlJc = OxmlElement('w:lvlJc')
                lvlJc.set(qn('w:val'), 'right')
                lvl.append(lvlJc)
                pPr = OxmlElement('w:pPr')
                ind = OxmlElement('w:ind')
                ind.set(qn('w:left'), '720')
                ind.set(qn('w:hanging'), '360')
                pPr.append(ind)
                lvl.append(pPr)
                abstractNum.append(lvl)
                numbering.insert(0, abstractNum)
                num = OxmlElement('w:num')
                num.set(qn('w:numId'), str(num_id))
                abstractNumId = OxmlElement('w:abstractNumId')
                abstractNumId.set(qn('w:val'), str(abs_num_id))
                num.append(abstractNumId)
                numbering.append(num)
                return num_id

            custom_num_id = None
            try:
                custom_num_id = add_custom_numbering(doc)
            except Exception as e:
                logger.warning(f"Không thể tạo custom numbering: {e}")

            for item in bib_items:
                full_text = str(item)
                full_text = re.sub(r'^\d+\.\s*', '', full_text)
                p_ref = doc.add_paragraph("")
                p_ref.style = norm_style
                if custom_num_id:
                    pPr = p_ref._element.get_or_add_pPr()
                    numPr = OxmlElement('w:numPr')
                    ilvl = OxmlElement('w:ilvl')
                    ilvl.set(qn('w:val'), '0')
                    numId = OxmlElement('w:numId')
                    numId.set(qn('w:val'), str(custom_num_id))
                    numPr.append(ilvl)
                    numPr.append(numId)
                    pPr.append(numPr)
                match_year = re.search(r'\(\d{4}\)', full_text)
                if match_year:
                    authors_part = full_text[:match_year.start()].strip()
                    rest_part = " " + full_text[match_year.start():]
                    run_authors = p_ref.add_run(authors_part)
                    run_authors.font.bold = True
                    run_rest = p_ref.add_run(rest_part)
                else:
                    parts = full_text.split(". ", 1)
                    if len(parts) > 1:
                        authors_part = parts[0] + "."
                        rest_part = " " + parts[1]
                        run_authors = p_ref.add_run(authors_part)
                        run_authors.font.bold = True
                        run_rest = p_ref.add_run(rest_part)
                    else:
                        run_all = p_ref.add_run(full_text)
                        run_all.font.bold = True
                apply_hardcoded_formatting(p_ref, 'Normal')

        doc.save(output_docx_file)
        logger.info(f"Xuất file thành công: {output_docx_file}")
        print(f'{{"status": "success", "file_path": "{output_docx_file}", "message": "Hoàn tất."}}')

    except Exception as e:
        logger.error(f"Lỗi tạo file Word: {e}")
        print(f'{{"status": "error", "message": "Lỗi: {str(e)}"}}')

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    if len(sys.argv) > 2:
        generate_manuscript_docx(sys.argv[1], sys.argv[2])
    else:
        print(json.dumps({"status": "error", "message": "Thiếu tham số."}))