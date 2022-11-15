from django import forms
from .models import Client, Reason


class FileFormatForm(forms.Form):
    client_field = forms.ModelChoiceField(
        queryset=Client.objects.all(),
        empty_label="-- Choisissez un client --",
        label='Client'
    )
    urssaf_name_field = forms.CharField(max_length=50, label='URSSAF')
    city_field = forms.CharField(max_length=50)
    send_date_field = forms.DateField(
        input_formats=['%d-%m-%Y',
                       '%d/%m/%Y',
                       '%d/%m/%y'],
        label='Date d\'envoi')
    period = forms.CharField(max_length=100, label='Période sur l\'objet')
    reason_field = forms.ModelChoiceField(
        queryset=Reason.objects.all(),
        empty_label="-- Choisissez un justificatif d'envoi --",
        label='Justificatif'
    )
