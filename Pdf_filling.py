import re
from pdfrw import PdfReader, PdfWriter, PdfName, PdfDict

# ================= 文本字段填写 =================
def set_text_field(pdf, field_name, value):
    for page in pdf.pages:
        if PdfName.Annots in page:
            for annot in page[PdfName.Annots]:
                if annot.get(PdfName.T):
                    name = annot[PdfName.T].to_unicode()
                    print(name)
                    if name == field_name:
                        annot.update(PdfDict(V=value, AP=''))
#
# def set_text_field(pdf, field_name, value):
#     found = False
#
#     for page in pdf.pages:
#         if PdfName.Annots in page:
#             for annot in page[PdfName.Annots]:
#                 if annot.get(PdfName.T):
#                     name = annot[PdfName.T].to_unicode().strip()
#
#                     if name == field_name:
#                         annot.update(PdfDict(V=value))
#                         found = True
#                         break
#
#         if found:
#             break

# ================= 频率解析 =================

frequency_map = {
    "L743-2": "4-3 MHz",
    "L761-2": "7-6 MHz",
    "L742UB": "7-4 MHz",
    "L1042UB": "10-4 MHz",
    "L552UB": "5-5 MHz",
    "C361-2": "6-3 MHz",
    "C321-2": "3-2 MHz",
    "C352UB": "5-3 MHz",
    "C611-2": "6-1 MHz",
    "C422UB": "4-2 MHz",
    "C612UB": "6-1 MHz",
    "C6152UB": "6-1 MHz",
    "E611-2": "6-1 MHz",
    "E612UB": "6-1 MHz",
}

def get_frequency(probe):

    # ✅ 1. 先查映射表
    if probe in frequency_map:
        return frequency_map[probe]

    # ✅ 2. 再用正则自动解析
    match = re.search(r'(\d+)-(\d+)', probe)
    if match:
        return f"{match.group(2)}-{match.group(1)} MHz"

    # ✅ 3. 都没有就返回空
    return ""

# ================= 填探头 =================
def fill_probes(pdf, probes):
    for i, probe in enumerate(probes[:5], start=1):
        set_text_field(pdf, f"Schallkopf {i}", probe)
        set_text_field(pdf, f"Schallkopf {i}_1", probe)
        freq = get_frequency(probe)
        set_text_field(pdf, f"Frequenz {i}", freq)
        set_text_field(pdf, f"Frequenz {i}_1", freq)

# ================= 标准化 AK 名称（🔥核心修复） =================
def normalize_name(name):
    name = name.lower()

    # AK2.3 → ak_2_3
    name = re.sub(r'ak\s*(\d+)\.(\d+)', r'ak_\1_\2', name)

    # 通用处理
    name = name.replace(" ", "_").replace(".", "_")
    name = name.replace("(", "").replace(")", "")

    return name

# ================= 勾选 checkbox（🔥核心修复） =================
def check_items_by_name(pdf, items_to_check):
    checked_count = 0

    normalized_items = [normalize_name(i) for i in items_to_check]
    print("目标字段:", normalized_items)

    for page_num, page in enumerate(pdf.pages, 1):
        if PdfName.Annots in page:
            for annot in page[PdfName.Annots]:

                if not annot.get(PdfName.T):
                    continue

                raw_name = annot[PdfName.T].to_unicode()
                field_name = normalize_name(raw_name)

                if field_name in normalized_items:

                    # print(f"\n🎯 命中字段: {raw_name} (页 {page_num})")

                    # ===== 获取真实状态 =====
                    on_value = None

                    if annot.get(PdfName.AP):
                        ap = annot[PdfName.AP]
                        if PdfName.N in ap:
                            for key in ap[PdfName.N].keys():
                                if key != PdfName.Off:
                                    on_value = key
                                    break

                    # print("👉 使用状态:", on_value)

                    # fallback
                    if not on_value:
                        on_value = PdfName.Yes

                    # ===== 勾选 =====
                    annot.update(PdfDict(
                        V=on_value,
                        AS=on_value
                    ))

                    checked_count += 1
                    print(f"✅ 已勾选: {raw_name}")

    print(f"\n完成勾选 {checked_count} 个checkbox")
    return checked_count

# ================= 主函数 =================
def fill_pdf_fields(
    pdf_path,
    probes=None,
    items_to_check=None,
    output_path=None,
    adresse="",
    name="",
    geraet="",
    baujahr="",
    datum=""
):
    probes = probes or []
    items_to_check = items_to_check or []

    pdf = PdfReader(pdf_path)

    # 文本字段
    set_text_field(pdf, "Adresse", adresse)
    set_text_field(pdf, "Bezeichnung", name)
    set_text_field(pdf, "GeräteNummer", geraet)
    set_text_field(pdf, "Baujahr", baujahr)
    set_text_field(pdf, "Auslieferungsdatum", datum)

    set_text_field(pdf, "datum", datum)

    ort_datum_value = f"Langen, {datum}"
    set_text_field(pdf, "Ort Datum", ort_datum_value)
    set_text_field(pdf, "Ort Datum_1", ort_datum_value)

    set_text_field(pdf, "Adresse_1", adresse)
    set_text_field(pdf, "Bezeichnung_1", name)
    set_text_field(pdf, "GeräteNummer_1", geraet)
    set_text_field(pdf, "Baujahr_1", baujahr)
    set_text_field(pdf, "Auslieferungsdatum_1", datum)

    # 探头
    fill_probes(pdf, probes)

    # checkbox
    check_items_by_name(pdf, items_to_check)

    # 强制刷新显示（很关键）
    if pdf.Root.AcroForm:
        pdf.Root.AcroForm.update(
            PdfDict(NeedAppearances=PdfName('True'))
        )

    # 保存
    PdfWriter(output_path, trailer=pdf).write()
    print(f"\n✅ PDF已保存到: {output_path}")

# ================= 示例 =================
if __name__ == "__main__":
    pdf_file = "KV Baden Württemberg.pdf"
    output_file = "Anlage_filled.pdf"

    probes = ["C5-2Q", "L12-5Q"]

    items_to_check = [
        "AK2.1",
        "AK2.3",
        "AK3.4",
        "AK10.2"
    ]

    fill_pdf_fields(
        pdf_path=pdf_file,
        probes=probes,
        items_to_check=items_to_check,
        output_path=output_file,
        adresse="Musterstraße 1, 12345 Stadt",
        name="Ultraschallgeraet",
        geraet="123456",
        baujahr="2023",
        datum="2026-04-09"
    )