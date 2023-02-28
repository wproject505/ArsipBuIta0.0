from django.urls import path
from .views import ajuan


urlpatterns = [
    path('', ajuan, name='home'),
    # path('generate-pdf/', generate_pdf, name='generate-pdf'),
    # path('ajuantk', views.ajuan_tk, name='ajuantk'),
    # path('invoice', views.invoice, name='invoice'),
]