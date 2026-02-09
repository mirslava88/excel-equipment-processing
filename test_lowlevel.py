#!/usr/bin/env python3
"""
Низкоуровневое чтение структуры Excel файла
"""
import zipfile
import xml.etree.ElementTree as ET

test_file = "/Users/mirslava/Desktop/myapps/ФОП_приоритетная поддержка 30-01.xlsx"

print("=" * 60)
print("НИЗКОУРОВНЕВЫЙ АНАЛИЗ XLSX ФАЙЛА")
print("=" * 60)

try:
    with zipfile.ZipFile(test_file, 'r') as z:
        print("\nСодержимое ZIP архива:")
        for name in z.namelist()[:20]:
            print(f"  {name}")
        
        # Пробуем прочитать workbook.xml
        if 'xl/workbook.xml' in z.namelist():
            print("\n✓ Найден xl/workbook.xml")
            content = z.read('xl/workbook.xml')
            root = ET.fromstring(content)
            
            print(f"\nXML root tag: {root.tag}")
            print(f"XML root attribs: {root.attrib}")
            
            # Пробуем найти все элементы без namespace
            all_elements = root.findall('.//')
            print(f"\nВсего элементов в XML: {len(all_elements)}")
            
            # Ищем любые теги, содержащие 'sheet'
            sheet_elements = [el for el in all_elements if 'sheet' in el.tag.lower()]
            print(f"Элементов со словом 'sheet': {len(sheet_elements)}")
            
            for el in sheet_elements[:10]:
                print(f"  Tag: {el.tag}, Attribs: {el.attrib}")
            
            # Попробуем несколько вариантов namespace
            namespaces = [
                {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'},
                {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/9/main'},
                {},
            ]
            
            for ns in namespaces:
                if ns:
                    sheets = root.findall('.//s:sheet', ns)
                else:
                    # Без namespace
                    sheets = [el for el in all_elements if el.tag.endswith('sheet') or el.tag.endswith('Sheet')]
                
                if sheets:
                    print(f"\n✓ Найдено листов (namespace {ns}): {len(sheets)}")
                    for sheet in sheets:
                        name = sheet.get('name') or sheet.get('Name')
                        print(f"  - {name} | attrs: {sheet.attrib}")
                    break
        else:
            print("\n❌ xl/workbook.xml не найден")
            
except Exception as e:
    print(f"\n❌ Ошибка: {e}")
    import traceback
    traceback.print_exc()
