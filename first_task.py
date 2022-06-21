import os
import xml.etree.ElementTree as Xet
import pandas as pd

current_directory = os.path.abspath(os.path.dirname('Data_Repositories_Test.docx'))
path = os.path.join(current_directory, 'Explainers')
files = os.listdir(path)
rows = []
for file in files:
    xml_parse = Xet.parse(os.path.join(path, file))
    name = file
    root = xml_parse.getroot()
    explainer_id = int(root.get('id'))
    seo_met = root.find("seo_meta_description")
    if seo_met is not None:
        seo_meta = root.find("seo_meta_description").text
        if seo_meta is None:
            seo_meta_description = root.find("seo_meta_description").text
        else:
            seo_meta_description = " ".join(
                root.find("seo_meta_description").text.split())
    else:
        seo_meta_description = "None"
    sourceid = root.find("source_id")
    if sourceid is not None:
        source = root.find("source_id").text
        if source is None:
            source_id = root.find("source_id").text
        else:
            source_id = int(root.find("source_id").text)
    else:
        source_id = "None"
    developer = root.find("developer_name")
    if developer is not None:
        developer_name = root.find("developer_name").text
    else:
        developer_name = "None"
    data = {'1': ' ',
            '2': ' '}
    rows.append({'1': name,
                 '2': explainer_id})
    rows.append({'1': ' ',
                 '2': seo_meta_description})
    rows.append({'1': ' ',
                 '2': source_id})
    rows.append({'1': ' ',
                 '2': developer_name})
    df = pd.DataFrame(rows, columns=data)
    writer = pd.ExcelWriter(os.path.join(current_directory, 'Output', 'first_task_Explainers', 'data.xlsx'), engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Explainers',
                header=False, na_rep='None', index=None)
    workbook = writer.book
    num_fmt = workbook.add_format({'num_format': '0', 'align': 'left'})
    worksheet = writer.sheets['Explainers']
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 130, num_fmt)
    writer.save()
