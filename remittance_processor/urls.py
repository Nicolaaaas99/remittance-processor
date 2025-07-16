# remittance_processor_project/remittance_processor/urls.py
from django.urls import path
from . import views

app_name = 'remittance_processor'  # Namespace for the app

urlpatterns = [
    path('', views.process_remittance_view, name='index'),  # Name the root URL as 'index'
    path('download_total_import/', views.download_total_import, name='download_total_import'),
    path('download_detail_import/', views.download_detail_import, name='download_detail_import'),
]