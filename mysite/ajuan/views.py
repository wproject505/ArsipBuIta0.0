from django.shortcuts import render
import pandas as pd
from django.http import HttpResponse
from reportlab.pdfgen import canvas
from .models import *

def ajuan(request):
    item = Ajuan.objects.all().values()
    df = pd.DataFrame(item)
    mydict = {
        'df':df.to_html()
    }
    return render(request, 'index.html',context=mydict)






