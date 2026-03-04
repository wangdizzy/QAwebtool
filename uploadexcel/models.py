from django.db import models

# Create your models here.
class UploadedExcel(models.Model):
    file = models.FileField(upload_to='excel_files/')
    uploaded_time = models.DateTimeField(auto_now_add=True)
    processed = models.BooleanField(default=True)

class ExcelData(models.Model):
    
    game_name = models.CharField(max_length=100)  # 遊戲名稱 (如 Aztec Smash)
    currency = models.CharField(max_length=10)    # 貨幣代碼 (如 CNY, USD)
    min_bet = models.DecimalField(max_digits=10, decimal_places=2)  # 最小下注
    max_bet = models.DecimalField(max_digits=10, decimal_places=2)  # 最大下注
    default_bet = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)  # 預設下注
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'excel_data'
    
    def __str__(self):
        return f"{self.game_name} - {self.currency}"