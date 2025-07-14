# remittance_processor_project/remittance_processor_project/urls.py
from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('remittance_processor.urls')),
]