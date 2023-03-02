from django.urls import path
from django.contrib import admin
from . import views


urlpatterns = [
    path('', admin.site.urls),
    path('generate_pdf/', views.generate_pdf, name='generate_pdf'),
    # path('generate-pdf/', generate_pdf, name='generate-pdf'),
    # path('ajuantk', views.ajuan_tk, name='ajuantk'),
    # path('invoice', views.invoice, name='invoice'),
]