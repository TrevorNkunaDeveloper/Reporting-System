from django import forms

class UploadFileForm(forms.Form):
    start_date = forms.DateField(widget=forms.SelectDateWidget())
    end_date = forms.DateField(widget=forms.SelectDateWidget())
    file = forms.FileField()
