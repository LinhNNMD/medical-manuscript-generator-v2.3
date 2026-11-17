"""
parse_clinical_data.py — Quét dữ liệu lâm sàng (ca lâm sàng, kết quả nghiên cứu) từ các thư mục tương ứng.
"""
import io
import os
import sys
import json
import logging
import argparse
from datetime import datetime

# Tránh lỗi ModuleNotFoundError khi chạy thư viện ngoài
try:
    import pypdf
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
    import pypdf

try:
    import docx
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("parse_clinical")


def extract_text_from_pdf(pdf_path):
    """Trích xuất văn bản từ tệp PDF."""
    try:
        reader = pypdf.PdfReader(pdf_path)
        text_parts = []
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
        return "\n".join(text_parts)
    except Exception as e:
        logger.error(f"Lỗi đọc PDF {pdf_path}: {e}")
        return ""


def extract_text_from_docx(docx_path):
    """Trích xuất văn bản từ tệp DOCX."""
    try:
        doc = docx.Document(docx_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text]
        return "\n".join(paragraphs)
    except Exception as e:
        logger.error(f"Lỗi đọc DOCX {docx_path}: {e}")
        return ""


def extract_text_from_text_file(file_path):
    """Trích xuất văn bản từ tệp txt, md, json."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception as e:
        logger.error(f"Lỗi đọc tệp {file_path}: {e}")
        return ""


def process_folder(folder_path):
    """Quét thư mục lâm sàng và trích xuất nội dung của tất cả các file hợp lệ."""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path, exist_ok=True)
        logger.info(f"Đã tạo thư mục trống {folder_path}. Chưa có tệp nào để quét.")
        return []

    allowed_exts = ('.pdf', '.docx', '.txt', '.md', '.json')
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(allowed_exts) and f.lower() != "readme.txt"]
    if not files:
        logger.info(f"Thư mục '{folder_path}' trống hoặc không chứa file dữ liệu lâm sàng hợp lệ.")
        return []

    logger.info(f"Tìm thấy {len(files)} tệp dữ liệu lâm sàng trong thư mục '{folder_path}'.")
    results = []

    for file in files:
        file_path = os.path.join(folder_path, file)
        logger.info(f"📄 Đang xử lý file dữ liệu: {file}")

        content = ""
        if file.lower().endswith('.pdf'):
            content = extract_text_from_pdf(file_path)
        elif file.lower().endswith('.docx'):
            content = extract_text_from_docx(file_path)
        elif file.lower().endswith(('.txt', '.md', '.json')):
            content = extract_text_from_text_file(file_path)

        if content.strip():
            results.append({
                "filename": file,
                "content": content.strip()
            })
        else:
            logger.warning(f"File {file} không trích xuất được nội dung.")

    return results


def main():
    parser = argparse.ArgumentParser(description="Scan and extract clinical cases & research results.")
    parser.add_argument("--cases-dir", type=str, default="clinical_cases", help="Path to clinical cases folder")
    parser.add_argument("--results-dir", type=str, default="research_results", help="Path to research results folder")
    parser.add_argument("--output", type=str, default="clinical_data_log.md", help="Path to output markdown file")
    parser.add_argument("--mode", type=str, choices=["cases", "results", "all"], default="all", help="Scan mode: cases, results, or all")
    args = parser.parse_args()

    # Nhận diện đường dẫn tuyệt đối cho output
    workspace_dir = os.getcwd()
    output_path = args.output
    if not os.path.isabs(output_path):
        output_path = os.path.join(workspace_dir, output_path)

    cases_data = []
    results_data = []

    if args.mode in ["cases", "all"]:
        cases_data = process_folder(args.cases_dir)
    else:
        logger.info("Chế độ quét bỏ qua thư mục ca lâm sàng (clinical_cases).")

    if args.mode in ["results", "all"]:
        results_data = process_folder(args.results_dir)
    else:
        logger.info("Chế độ quét bỏ qua thư mục kết quả nghiên cứu (research_results).")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("# NHẬT KÝ THÔNG TIN LÂM SÀNG VÀ KẾT QUẢ NGHIÊN CỨU CỦA NGƯỜI DÙNG\n\n")
        f.write(f"Thời gian quét: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Tổng số tệp thông tin ca lâm sàng: {len(cases_data)}\n")
        f.write(f"Tổng số tệp kết quả nghiên cứu: {len(results_data)}\n\n")

        f.write("## 1. Dữ liệu Thông tin Ca Lâm sàng (clinical_cases)\n")
        f.write("Dữ liệu dưới đây sẽ được sử dụng làm thông tin thực tế cốt lõi để viết Báo cáo Ca lâm sàng (Nhóm D).\n\n")
        if not cases_data:
            f.write("*Không tìm thấy dữ liệu ca lâm sàng nào được cung cấp.*\n\n")
        else:
            for item in cases_data:
                f.write(f"### 📄 Tệp ca bệnh: `{item['filename']}`\n")
                f.write("```text\n")
                f.write(item['content'])
                f.write("\n```\n\n")
                f.write("---\n\n")

        f.write("## 2. Dữ liệu Kết quả Nghiên cứu (research_results)\n")
        f.write("Dữ liệu dưới đây sẽ được sử dụng làm số liệu thực tế cốt lõi để viết Công bố Kết quả Nghiên cứu (Nhóm E).\n\n")
        if not results_data:
            f.write("*Không tìm thấy dữ liệu kết quả nghiên cứu nào được cung cấp.*\n\n")
        else:
            for item in results_data:
                f.write(f"### 📄 Tệp kết quả: `{item['filename']}`\n")
                f.write("```text\n")
                f.write(item['content'])
                f.write("\n```\n\n")
                f.write("---\n\n")

    logger.info(f"Đã xuất nhật ký dữ liệu lâm sàng thành công tại {output_path}")
    print(json.dumps({
        "status": "success",
        "clinical_cases_count": len(cases_data),
        "research_results_count": len(results_data),
        "message": "Quét và phân tích dữ liệu lâm sàng thành công."
    }))


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8") if hasattr(sys, 'stdout') and hasattr(sys.stdout, 'buffer') else sys.stdout
    main()
