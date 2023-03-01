from django.urls import path
from django.contrib import admin
from .views import ajuan


urlpatterns = [
    path('', admin.site.urls),
    # path('generate-pdf/', generate_pdf, name='generate-pdf'),
    # path('ajuantk', views.ajuan_tk, name='ajuantk'),
    # path('invoice', views.invoice, name='invoice'),
]