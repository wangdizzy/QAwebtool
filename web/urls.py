"""web URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from home.views import homepage, showpost, upload, oddsConversion
from gamecheck.views import ppsgFunctionSelection, report_function
from home.views import ppsg, check_paths
from uploadexcel.views import upload_excel
from checkWebFunction.views import test

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', homepage),
    path('post/<slug:slug>/', showpost),
    path('game/', ppsg),
    path('upload/', upload),
    path('ppsgFunctionSelection/', ppsgFunctionSelection),
    path('adminfunction/', report_function),
    path('oddsconversion/', oddsConversion),
    path('uploadExcel/', upload_excel),
    path('check-paths/', check_paths),
    path('checkweb/', test)
]
