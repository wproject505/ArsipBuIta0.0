from django.shortcuts import render
from django.http import HttpResponse
from reportlab.pdfgen import canvas
from .models import *
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from django.http import HttpResponse
from .models import BuktiKasKeluar

# def ajuan(request):
#     pass




def generate_pdf(request):
    # Get data from database
    bkk = BuktiKasKeluar.objects.first()

    # Create response object
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="kwitansi.pdf"'

    # Create PDF document
    pdf = canvas.Canvas(response, pagesize=letter)

    # Set font size and type
    pdf.setFont('Helvetica', 12)

    # Draw header
    pdf.drawCentredString(4.25*inch, 10*inch, 'KWITANSI PEMBAYARAN')
    pdf.line(1*inch, 9.5*inch, 7.5*inch, 9.5*inch)

    # Draw content
    pdf.drawString(1*inch, 9*inch, 'No. BKK : ' + bkk.no_BKK)
    pdf.drawString(1*inch, 8.75*inch, 'Tanggal : ' + bkk.tanggal_BKK.strftime('%d %B %Y'))
    pdf.drawString(1*inch, 8.5*inch, 'Dibayarkan kepada : ' + bkk.dibayarkan_kepada)
    pdf.drawString(1*inch, 8.25*inch, 'Uraian : ' + bkk.uraian)
    pdf.drawString(1*inch, 8*inch, 'Kode Bank : ' + bkk.kode_bank)
    pdf.drawString(1*inch, 7.75*inch, 'Nomor Cek : ' + bkk.nomer_cek)

    # Draw footer
    pdf.line(1*inch, 1*inch, 7.5*inch, 1*inch)
    pdf.drawString(1*inch, 0.75*inch, 'Terima kasih telah menggunakan layanan kami.')

    # Save PDF document
    pdf.save()

    return response





