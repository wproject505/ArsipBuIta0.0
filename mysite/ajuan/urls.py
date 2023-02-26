from django.urls import path
from . import views

urlpatterns = [
    path('', views.ajuan, name='home'),
    # path('ajuantk', views.ajuan_tk, name='ajuantk'),
    # path('invoice', views.invoice, name='invoice'),
]