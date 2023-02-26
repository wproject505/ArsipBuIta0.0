from django.shortcuts import render
from .models import *
import pandas as pd

def ajuan(request):
    item = Ajuan.objects.all().values()
    df = pd.DataFrame(item)
    mydict = {
        'df':df.to_html()
    }
    return render(request, 'index.html',context=mydict)




