from django.shortcuts import render
from django.http import StreamingHttpResponse, HttpResponse
from .forms import FileFormatForm
from docxtpl import DocxTemplate
from docx2pdf import convert
from pathlib import Path
import io
import os
import locale
from django.conf import settings
import random
import string


def index(request):
    context = {
        'form': FileFormatForm(),
    }
    return render(request, './generator/index.html', context)


def build_docx_document(data):
    template_dir = 'generator/templates/courrier_word'

    courrier_liste = [
        'attestation_CA',
        'attestation_compte_courant',
        'attestation_dividendes',
        'attestation_emploi',
        'attestation_non_emploi',
        'attestation_bnc',
        'changement_regime_IS',
        'changement_regime_BIC',
        'changement_regime_TVA',
        'centre_impots',
        'centre_retraite',
        'centre_urssaf',
        'lettre_mission',
        'mandat_prelevement_SEPA',
        ]

    courrier = courrier_liste[-1]

    if courrier == 'attestation_CA':
        data_folder = Path('./{0}/attestation_CA.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_representant': 'représentant',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'company_siren' : 'siren',
            'period1': True,
            'period1_benefits': 'bénéfices1 100',
            'period1_start': 'start 1',
            'period1_end': 'end 1',
            'period2': True,
            'period2_benefits': 'bénéfices2 200',
            'period2_start': 'start 2',
            'period2_end': 'end 2',
            'period3': False,
            'period3_benefits': 'bénéfices3 300',
            'period3_start': 'start 3',
            'period3_end': 'end 3',
            'send_date': 'date envoi',
            }
    elif courrier == 'attestation_compte_courant':
        data_folder = Path('./{0}/attestation_compte_courant.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'company_siren' : 'siren',
            'send_date': 'date envoi',
            'financial_contribution': 'financial_contribution',
            'expenditure_reasons': 'expenditure_reasons' 
            }
    elif courrier == 'attestation_dividendes':
        data_folder = Path('./{0}/attestation_dividendes.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'company_siren' : 'siren',
            'company_representant': 'représentant',
            'period1': True,
            'period1_amount': 'bénéfices1 100',
            'period1_start': 'start 1',
            'period1_end': 'end 1',
            'period2': True,
            'period2_amount': 'bénéfices2 200',
            'period2_start': 'start 2',
            'period2_end': 'end 2',
            'period3': False,
            'period3_amount': 'bénéfices3 300',
            'period3_start': 'start 3',
            'period3_end': 'end 3',
            'send_date': 'date envoi',
            }
    elif courrier == 'attestation_emploi':
        data_folder = Path('./{0}/attestation_emploi.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_representant': 'représentant',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'company_siren' : 'siren',
            'send_date': 'date envoi',
            'nb_employees': 'nb_employees' 
            }
    elif courrier == 'attestation_non_emploi':
        data_folder = Path('./{0}/attestation_non_emploi.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_representant': 'représentant',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'company_siren' : 'siren',
            'send_date': 'date envoi',
            'non_employment_date': 'non_employment_date' 
            }
    elif courrier == 'attestation_bnc':
        data_folder = Path('./{0}/attestation_bnc.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_representant': 'représentant',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'company_siren' : 'siren',
            'period1': True,
            'period1_benefits': 'bénéfices1 100',
            'period1_start': 'start 1',
            'period1_end': 'end 1',
            'period2': True,
            'period2_benefits': 'bénéfices2 200',
            'period2_start': 'start 2',
            'period2_end': 'end 2',
            'period3': False,
            'period3_benefits': 'bénéfices3 300',
            'period3_start': 'start 3',
            'period3_end': 'end 3',
            'send_date': 'date envoi',
            }
    elif courrier == 'changement_regime_IS':
        data_folder = Path('./{0}/changement_regime_IS.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'sie_city' : 'ville',
            'sie_postcode' : 'ville',
            'sie_address' : 'ville',
            'sie_complement' : 'ville',
            'send_date': 'date envoi',
            'company_siren' : 'siren',
            'financial_opening': 'financial_opening',
            'company_representant': 'représentant',
            }
    elif courrier == 'changement_regime_TVA':
        data_folder = Path('./{0}/changement_regime_TVA.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'sie_city' : 'ville',
            'sie_postcode' : 'ville',
            'sie_address' : 'ville',
            'sie_complement' : 'ville',
            'send_date': 'date envoi',
            'company_siren' : 'siren',
            'regime' : 'regime',
            'financial_opening': 'financial_opening',
            'company_representant': 'représentant',
            }
    elif courrier == 'changement_regime_BIC':
        data_folder = Path('./{0}/changement_regime_BIC.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'sie_city' : 'ville',
            'sie_postcode' : 'ville',
            'sie_address' : 'ville',
            'sie_complement' : 'ville',
            'send_date': 'date envoi',
            'company_siren' : 'siren',
            'financial_opening': 'financial_opening',
            'company_representant': 'représentant',
            }
    elif courrier == 'centre_impots':
        data_folder = Path('./{0}/centre_impots.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'taxes_center_address' : 'adresse',
            'taxes_center_complement' : 'complement',
            'taxes_center_postcode' : 'postcode',
            'taxes_center_city' : 'city',
            'send_date': 'date envoi',
            'taxes_type': 'taxes_type',
            'reason': 'reason',
            'schedule': True,
            'nb_months': 'NB MOIS',
            'due_taxes': 'DUE_TAXES',
            'monthly_amount': 'monthylyAMOUNT',
            'due_date': 'DUE_DATE',
            'remission_request': True,
            'increases_amount': 'MAJORATIONS',
            'increases_penalities': 'PENALITES',
            }
    elif courrier == 'centre_retraite':
        data_folder = Path('./{0}/centre_retraite.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'retirement_city' : 'ville',
            'retirement_postcode' : 'code postal',
            'retirement_address' : 'Adresse société',
            'retirement_complement' : 'Adresse société',
            'send_date': 'date envoi',
            'object_period': 'object_period',
            'reason': 'reason',
            'averaging_contributions': False,
            'total_contributions': 'total contributions',
            'period1': True,
            'period1_contributions': 'bénéfices1 100',
            'period1_start': 'start 1',
            'period1_end': 'end 1',
            'period2': True,
            'period2_contributions': 'bénéfices2 200',
            'period2_start': 'start 2',
            'period2_end': 'end 2',
            'period3': False,
            'period3_contributions': 'bénéfices3 300',
            'period3_start': 'start 3',
            'period3_end': 'end 3',
            'nb_months': '*nb month*',
            'payment_mode': '*payment_mode*',
            'due_date': '*due_date*',
            'discount': True,
            'surcharges_amount': '**majorations**',
            'penalities_amount': '**penalités**',
            }
    elif courrier == 'centre_urssaf':
        data_folder = Path('./{0}/centre_urssaf.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_city' : 'ville',
            'retirement_city' : 'ville',
            'retirement_postcode' : 'code postal',
            'retirement_address' : 'Adresse société',
            'retirement_complement' : 'Adresse société',
            'send_date': 'date envoi',
            'object_period': 'object_period',
            'reason': 'reason',
            'averaging_contributions': True,
            'total_contributions': 'total contributions',
            'period1': True,
            'period1_contributions': 'bénéfices1 100',
            'period1_start': 'start 1',
            'period1_end': 'end 1',
            'period2': True,
            'period2_contributions': 'bénéfices2 200',
            'period2_start': 'start 2',
            'period2_end': 'end 2',
            'period3': False,
            'period3_contributions': 'bénéfices3 300',
            'period3_start': 'start 3',
            'period3_end': 'end 3',
            'nb_months': '*nb month*',
            'payment_mode': '*payment_mode*',
            'due_date': '*due_date*',
            'discount': False,
            'surcharges_amount': '**majorations**',
            'penalities_amount': '**penalités**',
            }
    elif courrier == 'lettre_mission':
        data_folder = Path('./{0}/lettre_mission.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_city' : 'ville',
            'send_date': 'date envoi',
            'periodicity': '*** periodicity  ***',
            'company_creation': '*date de création*',
            'capital': '*capital*',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'company_sector' : 'code postal',
            'start_exercice' : 'ouverture exercice',
            'end_exercice' : 'cloture exercice',
            }
    elif courrier == 'mandat_prelevement_SEPA':
        data_folder = Path('./{0}/mandat_prelevement_SEPA.docx'.format(template_dir))
        context = {
            'company_name' : 'Company name',
            'company_city' : 'ville',
            'company_address' : 'Adresse société',
            'company_postcode' : 'code postal',
            'bank': 'bank',
            'code_iban': 'IBAN',
            'send_date': 'date envoi',
            }
    doc = DocxTemplate(data_folder)
    doc.render(context)
    return doc


def file_generator(request):
    form = FileFormatForm(request.POST)
    if request.method == 'POST' and form.is_valid():
        document = build_docx_document(form.cleaned_data)
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
            id = ''.join(random.SystemRandom().choice(string.ascii_letters + string.digits) for _ in range(10))

            docx_file = Path('{0}/test{1}.docx'.format(settings.MEDIA_ROOT, id))
            pdf_file = Path('{0}/test{1}.pdf'.format(settings.MEDIA_ROOT, id))
            document.save(docx_file)
            convert(docx_file, pdf_file)
            buffer_bytes = io.open(pdf_file, "rb", buffering = 0)
            response = StreamingHttpResponse(
                streaming_content=buffer_bytes,
                content_type='application/pdf'
            )
            response['Content-Disposition'] = 'attachment;filename=courrier.pdf'
            response["Content-Encoding"] = 'UTF-8'
            os.remove(docx_file)
            # os.remove(pdf_file)
            return response
        else:
            return HttpResponse('Something went wrong')
    else:
        return HttpResponse('Something went wrong not')
