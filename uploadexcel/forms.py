from django import forms
from .models import UploadedExcel

class ExcelUploadForm(forms.ModelForm):
    class Meta:
        model = UploadedExcel
        fields = ['file']
        
    def clean_file(self):
        file = self.cleaned_data['file']
        if file:
            if not file.name.endswith(('.xlsx', '.xls')):
                raise forms.ValidationError('請上傳 Excel 檔案 (.xlsx 或 .xls)')
        return file