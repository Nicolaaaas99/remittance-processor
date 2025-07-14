# remittance_processor_project/remittance_processor/urls.py
from django.urls import path
from . import views

app_name = 'remittance_processor'
urlpatterns = [
    path('', views.process_remittance_view, name='process_remittance'),
    path('download/total_import/', views.download_total_import, name='download_total_import'),
    path('download/detail_import/', views.download_detail_import, name='download_detail_import'),
]