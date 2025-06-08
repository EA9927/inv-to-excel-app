import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io

st.set_page_config(page_title="PDF 发票 ➡️ Excel 转换器", layout="centered")
st.title("📄 PDF Invoice ➡️ Excel 转换器")
st.markdown("上传您的发票 PDF 文件，系统将自动提取资料并生成 Excel 文件。\n\nUpload your invoice PDF below to generate an Excel report.")

uploaded_file = st.file_uploader("上传发票 PDF | Upload Invoice PDF", type=["pdf"])

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    invoices = []

    for page_num, page in enumerate(doc, start=1):
        text = page.get_text()

        # 尝试提取各个字段
        invoice_no = re.search(r"No\.\s+(IV-\d+)", text)
        date = re.search(r"Date\s+(\d{2}/\d{2}/\d{4})", text)
        desc_block = re.search(r"Description\s+Qty\s+U/Price\s+Amt\s+Tax\s+Net Amt\n(.+?)\s+\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}", text, re.DOTALL)
        qty = re.search(r"Description\s+Qty\s+U/Price\s+Amt\s+Tax\s+Net Amt\n.+?\s+(\d+)\s+", text, re.DOTALL)
        uprice = re.search(r"\s(\d+\.\d{2})\s+\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}", text)
        amount = re.search(r"\s\d+\.\d{2}\s+(\d+\.\d{2})\s+\d+\.\d{2}", text)
        tax = re.search(r"Service Tax \(8%\)\s+RM(\d+\.\d{2})", text)
        total = re.search(r"Total\s+RM(\d+\.\d{2})", text)

        data = {
            "Invoice No": invoice_no.group(1) if invoice_no else "",
            "Date": date.group(1) if date else "",
            "Description": desc_block.group(1).strip() if desc_block else "",
            "Qty": int(qty.group(1)) if qty else "",
            "Unit Price (RM)": float(uprice.group(1)) if uprice else "",
            "Amount (RM)": float(amount.group(1)) if amount else "",
            "Tax (RM)": float(tax.group(1)) if tax else "",
            "Subtotal (RM)": (float(total.group(1)) - float(tax.group(1))) if total and tax else "",
            "Total (RM)": float(total.group(1)) if total else "",
        }

        # 检查是否缺字段
        missing_fields = [key for key, value in data.items() if value == ""]
        if missing_fields:
            data["Status"] = f"⚠️ Missing: {', '.join(missing_fields)}"
        else:
            data["Status"] = "✅ Complete"

        invoices.append(data)

    # 生成 Excel 表格
    if invoices:
        df = pd.DataFrame(invoices)
        st.success(f"✅ 成功提取 {len(df)} 张发票（部分可能缺字段）")
        st.dataframe(df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Invoices")

        st.download_button(
            label="📥 下载 Excel 文件 | Download Excel",
            data=output.getvalue(),
            file_name="invoices_all_pages.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ 没有找到任何发票内容，请确认 PDF 格式。")
