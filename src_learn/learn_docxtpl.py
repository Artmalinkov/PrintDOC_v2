from docxtpl import DocxTemplate

context = {
    'company_name': 'Название компании',
    'number': 'Произвольный номер',
    'number_2': 'Номер на нижнем колонтитуле'}
doc = DocxTemplate(r"src_learn/word_tpl.docx")

doc.render(context)
doc.save(r"src_learn/generated_docx.docx")