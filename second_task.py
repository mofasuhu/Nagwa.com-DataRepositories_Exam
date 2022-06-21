import os
import xml.etree.ElementTree as Xet
import pandas as pd


current_directory = os.path.abspath(os.path.dirname('Data_Repositories_Test.docx'))
path = os.path.join(current_directory, 'Video_Transcripts')
files = os.listdir(path)
for file in files:
    xml_parse = Xet.parse(os.path.join(path, file))
    name = file
    root = xml_parse.getroot()
    transcript = root.find("transcript")
    questions = transcript.findall("question")
    rows = []
    for question in questions:
        questionid = question.get('id')
        if questionid is not None:
            question_idi = question.get('id')
            if question_idi == "":
                question_id = "None"
            else:
                question_id = int(question.get('id'))
        else:
            question_id = question.get('id')
        questionmediaidentifier = question.get('media_identifier')
        if questionmediaidentifier is not None:
            question_mediaidentifier = question.get('media_identifier')
            if question_mediaidentifier == "":
                question_media_identifier = "None"
            else:
                question_media_identifier = int(
                    question.get('media_identifier'))
        else:
            question_media_identifier = question.get('media_identifier')
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

        questiontitle = question.find("question_title")
        if questiontitle is not None:
            questiontitlespaces = question.find("question_title").text
            if questiontitlespaces is None:
                question_title = question.find("question_title").text
            else:
                question_title = " ".join(
                    question.find("question_title").text.split())
        else:
            question_title = "None"
        questionps = question.findall("p")
        question_p = questionps[0]
        question_start = question_p.findall("s")
        question_start_time = question_start[0]
        question_start_time_attribute = question_start_time.get("start_time")
        question_end_p = questionps[-1]
        question_end = question_end_p.findall("s")
        question_end_time = question_end[-1]
        question_end_time_attribute = question_end_time.get("end_time")
        rows.append({'question_id': question_id,
                     'q_media_identifier': question_media_identifier,
                     'question_title': question_title,
                     'q_start_time': question_start_time_attribute,
                     'q_end_time': question_end_time_attribute})
    df = pd.DataFrame(rows)
    df = df.fillna('None')
    writer = pd.ExcelWriter(os.path.join(current_directory, 'Output', 'second_task_Video_Transcripts', f'{name[:-4]}.xlsx'), engine='xlsxwriter')
    df.to_excel(writer, sheet_name='video_transcript_questions', index=False, encoding='utf-8')
    workbook = writer.book
    num_fmt = workbook.add_format({'num_format': '0', 'align': 'left'})
    worksheet = writer.sheets['video_transcript_questions']
    worksheet.set_column('A:A', 12, num_fmt)
    worksheet.set_column('B:B', 20, num_fmt)
    worksheet.set_column('C:C', 80)
    worksheet.set_column('D:D', 13, num_fmt)
    worksheet.set_column('E:E', 13, num_fmt)
    writer.save()
