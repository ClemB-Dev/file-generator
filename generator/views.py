from django.shortcuts import render
from django.http import StreamingHttpResponse, HttpResponse, FileResponse
from .forms import FileFormatForm
import io
import aspose.words as aw
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import locale
locale.setlocale(locale.LC_ALL, 'fr_FR')


def index(request):
    context = {
        'form': FileFormatForm(),
    }
    return render(request, './generator/index.html', context)


def build_document(data):
    document = Document()
    print(type(document))

    client_data = document.add_paragraph()
    run = client_data.add_run()
    run.bold = True
    run.add_text(str(data["client_field"]))
    run.add_break()
    run = client_data.add_run()
    run.bold = False
    run.add_text('54 rue Eugène Dupuis')
    run.add_break()
    run.add_text('94000 Créteil')

    urssaf_data = document.add_paragraph()
    paragraph_format = urssaf_data.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph_format.space_after = Inches(0.5)
    run = urssaf_data.add_run()
    run.bold = True
    run.add_text(str(data["urssaf_name_field"]))
    run.add_break()
    run = urssaf_data.add_run()
    run.bold = False
    run.add_text('54 rue Eugène Dupuis')
    run.add_break()
    run.add_text('94000 Créteil')

    city_and_date = document.add_paragraph()
    paragraph_format = city_and_date.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph_format.space_after = Inches(1)
    run = city_and_date.add_run()
    run.add_text('{0}, le {1}'.format(
        str(data['city_field']).capitalize(),
        data['send_date_field'].strftime("%d %B %Y")))

    document.add_paragraph('Madame, Monsieur,')

    content = document.add_paragraph()
    run = content.add_run()
    run.add_text(str(data['reason_field']))

    document.add_paragraph('Restant à votre disposition pour '
                           'tout renseignement complémentaire, '
                           'nous vous prions d\'agréer, '
                           'Madame, Monsieur, nos sincères salutations.')

    return document


def file_generator(request):
    form = FileFormatForm(request.POST)
    document = build_document(form.cleaned_data)
    if request.method == 'POST' and form.is_valid():
        if 'Word' in request.POST:
            buffer_bytes = io.BytesIO()
            document.save(buffer_bytes)
            buffer_bytes.seek(0)

            response = StreamingHttpResponse(
                streaming_content=buffer_bytes,
                content_type='application/vnd.openxmlformats-officedocument'
                             '.wordprocessingml.document'
            )

            response['Content-Disposition'] = 'attachment;filename=courrier.docx'
            response["Content-Encoding"] = 'UTF-8'
            return response
        elif 'PDF' in request.POST:
            response = StreamingHttpResponse(
                streaming_content=buffer_bytes,
                content_type='application/pdf'
            )

            response['Content-Disposition'] = 'attachment;filename=courrier.pdf'
            response["Content-Encoding"] = 'UTF-8'
            return response
        else:
            return HttpResponse('Something went wrong')
