from pdfrw import PdfReader, PdfWriter, PdfName, PdfDict
from Excel import find_items

def fill_checkboxes_by_index(pdf_path, items_to_check, output_path):
    """
    通过预设索引勾选复选框并保存PDF
    """
    # 读取PDF
    pdf = PdfReader(pdf_path)

    index_mapping = {
        # 第2页 (根据你的PDF结构调整具体索引)
        'AK1.1': (2, 0),
        'AK2.1': (2, 1),
        'AK2.2': (2, 2),
        'AK2.3': (2, 3),
        'AK2.4': (2, 4),
        'AK2.5': (2, 5),
        'AK2.6': (2, 6),
        'AK3.1': (2, 7),
        'AK3.2': (2, 8),
        'AK3.3': (2, 9),
        'AK3.4': (2, 10),
        'AK4.1': (2, 11),
        'AK4.2': (2, 12),
        'AK4.3': (2, 13),
        'AK4.4': (2, 14),
        'AK4.5': (2, 15),
        'AK4.6': (2, 16),
        'AK5.1': (2, 17),
        'AK5.2': (2, 18),
        'AK6.1': (2, 19),

        # 第3页
        'AK7.1': (3, 0),
        'AK7.2': (3, 1),
        'AK7.3': (3, 2),
        'AK8.1': (3, 3),
        'AK8.2': (3, 4),
        'AK8.3': (3, 5),
        'AK8.4': (3, 6),
        'AK8.5': (3, 7),
        'AK9.1': (3, 8),
        'AK9.2': (3, 9),
        'AK10.1': (3, 10),
        'AK10.2': (3, 11),
        'AK11.1': (3, 12),
        'AK12.1': (3, 13),
        'AK12.2': (3, 14),
        'AK20.1': (3, 15),
        'AK20.2': (3, 16),
        'AK20.3': (3, 17),
        'AK20.4': (3, 18),
        'AK20.5': (3, 19),
        'AK20.6': (3, 20),
        'AK20.7': (3, 21),
        'AK20.8': (3, 22),
        'AK20.9': (3, 23),
        'AK20.10': (3, 24),

        # 第4页
        'AK21.1': (4, 0),
        'AK21.2': (4, 1),
        'AK21.3': (4, 2),
        'AK21.4': (4, 3),
        'AK21.5': (4, 4),
        'AK21.6': (4, 5),
        'AK21.7': (4, 6),
        'AK21.8': (4, 7),
        'AK22.1': (4, 8),
        'AK22.2': (4, 9),
        'AK23.1': (4, 10),
    }

    # 标准化需要勾选的项（处理空格问题）
    normalized_items = []
    for item in items_to_check:
        # 移除空格并转为大写
        clean_item = item.replace(' ', '').upper()
        normalized_items.append(clean_item)

    print(f"需要勾选的项: {items_to_check}")

    # 统计勾选数量
    checked_count = 0

    # 遍历每个需要勾选的项
    for item in items_to_check:
        # 尝试多种匹配方式
        matched = False
        for key in [item, item.replace(' ', ''), item.upper(), item.replace(' ', '').upper()]:
            if key in index_mapping:
                page_num, field_idx = index_mapping[key]

                # 获取指定页面（页码从1开始，所以要减1）
                if page_num <= len(pdf.pages):
                    page = pdf.pages[page_num - 1]

                    # 检查页面是否有表单字段
                    if PdfName.Annots in page:
                        annotations = page[PdfName.Annots]

                        # 检查字段索引是否有效
                        if field_idx < len(annotations):
                            annotation = annotations[field_idx]
                            # 勾选复选框 - 多种方式尝试
                            try:
                                # 方法1: 设置值和外观状态
                                annotation.V = '/Yes'
                                annotation.AS = '/Yes'

                                # 方法2: 使用update
                                annotation.update(PdfDict(
                                    V='/Yes',
                                    AS='/Yes'
                                ))

                                # 方法3: 有些PDF使用/On而不是/Yes
                                if hasattr(annotation, 'AP'):
                                    if PdfName.N in annotation.AP:
                                        # 如果有外观字典，尝试使用第一个外观状态
                                        appearances = annotation.AP[PdfName.N]
                                        if appearances:
                                            first_state = list(appearances.keys())[0]
                                            if first_state != '/Off':
                                                annotation.AS = first_state
                            except Exception as e:
                                print(f"  勾选时出错: {e}")

                            print(f"✅ 已勾选: {item} (页面 {page_num}, 索引 {field_idx})")
                            checked_count += 1
                            matched = True
                            break
                        else:
                            print(f"⚠️ 页面 {page_num} 没有索引 {field_idx} 的字段 (最大索引: {len(annotations) - 1})")
                    else:
                        print(f"⚠️ 页面 {page_num} 没有表单字段")

        if not matched:
            print(f"❌ 未找到 {item} 的映射")

    # 保存PDF
    if checked_count > 0:
        PdfWriter(output_path, trailer=pdf).write()
        print(f"\n{'=' * 50}")
        print(f"✅ 成功勾选 {checked_count} 个复选框")
        print(f"📁 PDF文件已保存到: {output_path}")
        print(f"{'=' * 50}")
    else:
        print("\n❌ 没有勾选任何复选框")

    return checked_count > 0


def verify_pdf_structure(pdf_path):
    """
    验证PDF结构，帮助确定正确的索引
    """
    pdf = PdfReader(pdf_path)

    print(f"PDF文件: {pdf_path}")
    print(f"总页数: {len(pdf.pages)}")
    print("\n各页表单字段数量:")

    for page_num, page in enumerate(pdf.pages, 1):
        if hasattr(page, 'Annots') or (PdfName.Annots in page):
            # 兼容不同的PDF结构
            if PdfName.Annots in page:
                annotations = page[PdfName.Annots]
            else:
                annotations = page.Annots

            field_count = len(annotations)
            print(f"页面 {page_num}: {field_count} 个表单字段")

            # 显示前几个字段的名称和类型
            for i, annot in enumerate(annotations[:5]):  # 只显示前5个
                field_name = "未命名"
                field_type = "未知"

                if PdfName.T in annot:
                    field_name = annot[PdfName.T].to_unicode()
                if PdfName.FT in annot:
                    field_type = annot[PdfName.FT]

                print(f"  字段 {i}: {field_name} (类型: {field_type})")

            if field_count > 5:
                print(f"  ... 还有 {field_count - 5} 个字段")
        else:
            print(f"页面 {page_num}: 没有表单字段")

    print("\n" + "=" * 50)


# 主程序
if __name__ == "__main__":

    # 步骤2: 从Excel获取需要勾选的项目
    print("\n步骤1: 从Excel读取数据")
    # 请将 'your_column_name' 替换为Excel中实际的列名
    # items_to_check = find_items('your_column_name')
    col = input("请输入要查询的列名: ")
    items_to_check = find_items(col)

    # 或者手动指定测试数据
    # items_to_check = ['AK3.2', 'AK3.3', 'AK3.4', 'AK5.1', 'AK 6.1', 'AK8.1']
    # print(f"需要勾选的项目: {items_to_check}")

    # 步骤3: 勾选复选框并保存
    print("\n步骤2: 开始勾选复选框")
    output_file = "Anlage_checked.pdf"  # 输出文件名

    success = fill_checkboxes_by_index("Anlage.pdf", items_to_check, output_file)

    if success:
        print("\n✨ 处理完成！")
    else:
        print("\n⚠️ 处理可能不完整，请检查索引映射")

        # 如果失败，提供调试建议
        print("\n调试建议:")
        print("1. 检查 verify_pdf_structure 输出的字段数量")
        print("2. 根据实际字段数量调整 index_mapping 中的索引值")
        print("3. 确认PDF文件路径是否正确")