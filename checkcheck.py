from PyPDF2 import PdfReader

PDF_PATH = "KV Hessen.pdf"

# ================= 提取文本和表单字段 =================
def get_text_and_fields(pdf_path):
    reader = PdfReader(pdf_path)

    # 提取文本
    text = "\n".join(
        page.extract_text() or "" for page in reader.pages
    )

    # 提取表单字段名
    fields = reader.get_fields() or {}
    field_names = set(fields.keys())

    return text, field_names

# ================= 生成 AK 字段列表 =================
def generate_ak_list():
    ak = []

    ak += ["AK 1.1"]
    ak += [f"AK 2.{i}" for i in range(1, 7)]
    ak += [f"AK 3.{i}" for i in range(1, 5)]
    ak += [f"AK 4.{i}" for i in range(1, 7)]
    ak += [f"AK 5.{i}" for i in range(1, 3)]
    ak += ["AK 6.1"]
    ak += [f"AK 7.{i}" for i in range(1, 4)]
    ak += [f"AK 8.{i}" for i in range(1, 6)]
    ak += [f"AK 9.{i}" for i in range(1, 3)]
    ak += [f"AK 10.{i}" for i in range(1, 3)]
    ak += ["AK 11.1"]
    ak += [f"AK 12.{i}" for i in range(1, 3)]
    ak += [f"AK 20.{i}" for i in range(1, 13)]
    ak += [f"AK 21.{i}" for i in range(1, 9)]
    ak += [f"AK 22.{i}" for i in range(1, 3)]
    ak += ["AK 23.1"]

    return ak

# ================= 字段标准化 =================
def normalize(ak):
    # "AK 20.11" -> "ak_20_11"
    return ak.lower().replace(" ", "_").replace(".", "_")

# ================= 主函数 =================
def main():
    text, fields = get_text_and_fields(PDF_PATH)
    ak_list = generate_ak_list()

    results = []

    for ak in ak_list:
        # exists = (
        #     re.search(rf"\b{re.escape(ak)}\b", text) is not None
        #     or normalize(ak) in {f.lower() for f in fields}  # 这里改
        # )
        exists = (normalize(ak) in {f.lower() for f in fields})
        results.append((ak, exists))

    # 输出
    print("\n结果：")
    for ak, ok in results:
        print(f"{ak:<8} {'✓' if ok else '✗'}")

    # 统计
    total = len(results)
    ok = sum(r[1] for r in results)
    print(f"\n完成度: {ok}/{total} ({ok/total:.1%})")

# ================= 运行 =================
if __name__ == "__main__":
    main()