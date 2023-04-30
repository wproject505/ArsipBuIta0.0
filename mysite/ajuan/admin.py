import babel
from .models import *
from django.http import HttpResponse
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from .models import BuktiKasKeluar
from babel.numbers import format_currency
from decimal import Decimal
from reportlab.lib.pagesizes import letter, A4
from openpyxl import Workbook
from reportlab.lib.pagesizes import landscape, portrait, A5
from django.contrib import admin
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

styles = getSampleStyleSheet()


class DanaMasukAdmin(admin.ModelAdmin):
    search_fields = ('nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana')
    list_display = ('nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'get_total_dana')
    actions = ["export_as_pdf", "export_to_excel"]

    def export_to_excel(self, request, queryset):
        # Query data dari model Dana Masuk
        data = queryset.values()

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana'])

        # Tambahkan data ke file Excel
        for item in data:
            row = [item['nama_dana_masuk'], item['waktu_masuk'], item['penanggung_jawab'], item['total_dana']]
            ws.append(row)

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=Dana_Masuk.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'

    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="Dana Masuk.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(A4))
        elements = []
        styles = getSampleStyleSheet()
        title_style = styles['Heading1']
        title_style.alignment = TA_CENTER
        title = Paragraph('Dana Masuk', style=title_style)
        elements.append(title)
        data = []
        data.append(['No.','Nama Dana Masuk', 'Waktu Masuk', 'Penanggung Jawab', 'Total Dana'])
        total = Decimal(0)
        row_num = 1
        for dana_masuk in queryset:
            total_ajuan = Decimal(dana_masuk.total_dana)
            row = [
                row_num,
                dana_masuk.nama_dana_masuk,
                dana_masuk.waktu_masuk,
                dana_masuk.penanggung_jawab,
                format_currency(dana_masuk.total_dana, 'IDR', locale='id_ID'),
            ]
            data.append(row)
            row_num += 1

            # Add total_ajuan to total
            total += total_ajuan

            # Add row for total_ajuan
        data.append(['Total', '', '','', format_currency(total, 'IDR', locale='id_ID')])

        table = Table(data, colWidths=[70, 200, 150, 150,])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # Set "NAMA KEGIATAN" in the second row to align left
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (1, 0), (-1, -1), 100),
        ]))

        elements.append(table)
        doc.build(elements)

        return response

    export_as_pdf.short_description = "Export selected as PDF"

    def get_total_dana(self, obj):
        return babel.numbers.format_currency(obj.total_dana, 'IDR', locale='id_ID')
    get_total_dana.short_description = 'Total Dana'


class BuktiKasKeluarAdmin(admin.ModelAdmin):
    search_fields = ('no_BKK', 'tanggal_BKK','ajuan__nomor_pengajuan', 'dibayarkan_kepada', 'uraian', 'kode_bank', 'nomer_cek')
    list_display = ('no_BKK', 'tanggal_BKK', 'ajuan','get_total_ajuan', 'dibayarkan_kepada', 'uraian', 'kode_bank', 'nomer_cek')
    actions = ["export_as_pdf_global","export_to_excel"]
    raw_id_fields = ['ajuan', 'nomer_cek']

    def get_total_ajuan(self, obj):
        total_ajuan = obj.ajuan.total_ajuan
        return babel.numbers.format_currency(total_ajuan, 'IDR', locale='id_ID')
    get_total_ajuan.short_description = 'Total Ajuan'

    def export_to_excel(self, request, queryset):
        # Query data dari model BuktiKasKeluar
        data = queryset.values()

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['No BKK', 'Tanggal BKK', 'Ajuan', 'Dibayarkan Kepada', 'Uraian', 'Kode Bank', 'Nomor Cek'])

        # Tambahkan data ke file Excel
        for item in data:
            row = [item['no_BKK'], item['tanggal_BKK'], item['ajuan_id'], item['dibayarkan_kepada'], item['uraian'],
                   item['kode_bank'], item['nomer_cek']]
            ws.append(row)

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=BuktiKasKeluar.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'


    def export_as_pdf_global(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="bukti_kas_keluar.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))
        elements = []
        title_style = styles['Heading1']
        title_style.alignment = TA_CENTER
        title = Paragraph('Bukti Kas Keluar', style=title_style)
        elements.append(title)
        data = []
        data.append(['No.','Nomor BKK', 'Tanggal BKK', 'Nomor Cek', 'Dibayarkan Kepada', 'Uraian', 'Kode Bank', 'Ajuan'])
        total = Decimal(0)
        row_num = 1
        for bukti_kas_keluar in queryset:
            total_ajuan = Decimal(bukti_kas_keluar.ajuan.total_ajuan)
            row = [
                row_num,
                bukti_kas_keluar.no_BKK,
                bukti_kas_keluar.tanggal_BKK,
                bukti_kas_keluar.dibayarkan_kepada,
                bukti_kas_keluar.uraian,
                bukti_kas_keluar.kode_bank,
                bukti_kas_keluar.nomer_cek,
                format_currency(total_ajuan, 'IDR', locale='id_ID'),
            ]
            data.append(row)

            # Add total_ajuan to total
            total += total_ajuan

            # Add row for total_ajuan
        data.append(['', '', '', '', '','', 'Total', format_currency(total, 'IDR', locale='id_ID')])
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # Set "NAMA KEGIATAN" in the second row to align left
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (1, 0), (-1, -1), 100),
        ]))

        elements.append(table)
        doc.build(elements)

        return response

    export_as_pdf_global.short_description = "Export selected as Global PDF"


class AjuanAdmin(admin.ModelAdmin):
    search_fields = ('nomor_pengajuan', 'nama_kegiatan', 'waktu_ajuan', 'total_ajuan')
    list_display = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab', 'RAPT',  ]
    fields = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab', 'RAPT',  ]
    readonly_fields = ['nomor_pengajuan', ]
    raw_id_fields = ["RAPT",]
    actions = ['export_as_pdf','export_to_excel']


    def get_total_ajuan(self, obj):
        return babel.numbers.format_currency(obj.total_ajuan, 'IDR', locale='id_ID')
    get_total_ajuan.short_description = 'Total Ajuan'

    def export_to_excel(self, request, queryset):
        # Query data dari model BuktiKasKeluar
        data = queryset.values('unit_ajuan__unit_ajuan', 'nomor_pengajuan', 'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'RAPT', 'RPC')

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['unit_ajuan__unit_ajuan','nomor_pengajuan','nama_kegiatan','waktu_ajuan','total_ajuan','RAPT','RPC'])

        # Tambahkan data ke file Excel
        for item in data:
            row = [item['unit_ajuan__unit_ajuan'], item['nomor_pengajuan'], item['nama_kegiatan'], item['waktu_ajuan'], item['total_ajuan'],
                   item['RAPT'], item['RPC']]
            ws.append(row)

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=Ajuan.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'

    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="ajuan.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))
        elements = []
        title_style = styles['Heading1']
        title_style.alignment = TA_CENTER
        title = Paragraph('Daftar Ajuan', style=title_style)
        elements.append(title)
        data = []
        data.append(['NO','UNIT AJUAN', 'NOMOR PENGAJUAN', 'NAMA KEGIATAN', 'WAKTU AJUAN', 'TOTAL AJUAN', 'RAPT',])
        total = Decimal(0)
        row_num = 1
        for ajuan in queryset:
            total_ajuan = Decimal(ajuan.total_ajuan)
            row = [
                row_num,
                ajuan.unit_ajuan,
                ajuan.nomor_pengajuan,
                ajuan.nama_kegiatan,
                ajuan.waktu_ajuan,
                format_currency(ajuan.total_ajuan, 'IDR', locale='id_ID'),
                ajuan.RAPT,
            ]
            data.append(row)
            row_num += 1

            # Add total_ajuan to total
            total += total_ajuan

            # Add row for total_ajuan
        data.append(['Total', '', '', '','', format_currency(total, 'IDR', locale='id_ID', )])


        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # Set "NAMA KEGIATAN" in the second row to align left
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (1, 0), (-1, -1), 100),
        ]))




        elements.append(table)
        doc.build(elements)

        return response

    export_as_pdf.short_description = "Export selected as PDF"

class AjuanInLineRAPT(admin.TabularInline):
    model = Ajuan
    extra = 0
    fields = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab']
    readonly_fields = fields


class RekapAjuanPengambilanTabunganAdmin(admin.ModelAdmin):
    search_fields = ('no_RAPT', 'jumlah')
    list_display = ('no_RAPT', 'get_jumlah','get_nomor_pengajuan','get_total_ajuan')
    readonly_fields = ('no_RAPT',)
    inlines = [AjuanInLineRAPT]

    def get_nomor_pengajuan(self, obj):
        nomor_pengajuan_list = []
        for ajuan in obj.ajuan_set.all():
            nomor_pengajuan_list.append(ajuan.nomor_pengajuan)
        if nomor_pengajuan_list:
            return ", ".join(nomor_pengajuan_list)
        else:
            return '-'

    get_nomor_pengajuan.short_description = 'nomor pengajuan'

    def get_total_ajuan(self, obj):
        total_ajuan_list = []
        for ajuan in obj.ajuan_set.all():
            total_ajuan_list.append(str(ajuan.total_ajuan))  # ubah objek Decimal menjadi str
        if total_ajuan_list:
            return ", ".join(total_ajuan_list)
        else:
            return '-'

    get_total_ajuan.short_description = 'total_ajuan'

    fieldsets = [
        (None, {'fields': ['no_RAPT']}),
        (None, {'fields': ['jumlah']}),
    ]

    actions = ["export_as_pdf", "export_to_excel"]

    def get_jumlah(self, obj):
        return babel.numbers.format_currency(obj.jumlah, 'IDR', locale='id_ID')
    get_jumlah.short_description = 'Jumlah'

    def export_to_excel(self, request, queryset):
        # Query data dari model Dana Masuk
        data = queryset.values()

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['no_RAPT', 'jumlah'])

        # Tambahkan data ke file Excel
        for item in data:
            row = [item['no_RAPT'], item['jumlah']]
            ws.append(row)

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=RAPT.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'

    def export_to_excel(self, request, queryset):
        # Query data dari model Dana Masuk
        data = queryset.values()

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['no_RAPT', 'jumlah'])

        # Tambahkan data ke file Excel
        for item in data:
            row = [item['no_RAPT'], item['jumlah']]
            ws.append(row)

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=RAPT.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'

    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="RAPT.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))

        data = []
        data.append(['no_RAPT', 'jumlah'])
        total = Decimal(0)
        for RAPT in queryset:
            row = [
                RAPT.no_RAPT,
                RAPT.jumlah
            ]
            data.append(row)

            # Add total_ajuan to total
            total += RAPT.jumlah

            # Add row for total_ajuan
        data.append(['Total', format_currency(total, 'IDR', locale='id_ID')])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))


        elements = []
        elements.append(table)
        doc.build(elements)

        return response

    export_as_pdf.short_description = "Export selected as PDF"

from django.http import HttpResponse
class CekAdmin(admin.ModelAdmin):
    raw_id_fields = ['ajuan', 'RPC', ]
    search_fields = ('ajuan', 'RPC',)
    list_display = ('ajuan', 'RPC',)
    actions = ["export_as_pdf"]

    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="Cek.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))
        elements = []

        title_style = styles['Heading1']
        title_style.alignment = TA_CENTER
        title = Paragraph('Daftar Cek', style=title_style)
        elements.append(title)
        row_num = 1
        data = [['NO.', 'NO CEK', 'AJUAN', 'NOMER BANK TERTARIK']]
        for cek in queryset:
            row =[
                row_num,
                cek.no_cek,
                cek.ajuan,
                cek.RPC,
            ]
            data.append(row)
            row_num += 1
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # Set "NAMA KEGIATAN" in the second row to align left
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('WORDWRAP', (1, 0), (-1, -1), 100),
        ]))
        elements.append(table)
        doc.build(elements)
        return response

    export_as_pdf.short_description = "Export selected as PDF"


class CekInLineRPC(admin.TabularInline):
    model = Cek
    extra = 0
    fields = ['no_cek', 'keterangan', 'nomer_bank_tertarik', 'get_total_ajuan','get_nama_kegiatan',]
    readonly_fields = fields

    def get_total_ajuan(self, obj):
        return obj.ajuan.total_ajuan

    get_total_ajuan.short_description = 'Total Ajuan'

    def get_nama_kegiatan(self, obj):
        return obj.ajuan.nama_kegiatan

    get_total_nama_kegiatan_description = 'Nama Kegiatan'


class RekapPencairanCekAdmin(admin.ModelAdmin):
    search_fields = ('no_RPC',)
    readonly_fields = ('no_RPC', 'jumlah',)
    inlines = [CekInLineRPC]
    list_display = ('no_RPC', 'jumlah', 'nomer_bank_tertarik', 'get_total_ajuan', 'get_nama_kegiatan')

    def get_total_ajuan(self, obj):
        total_ajuan_list = []
        for cek in obj.cek_set.all():
            if cek.ajuan:
                total_ajuan_list.append(str(cek.ajuan.total_ajuan))
        if total_ajuan_list:
            return ", ".join(total_ajuan_list)
        else:
            return '-'

    get_total_ajuan.short_description = 'Total Ajuan'

    def get_nama_kegiatan(self, obj):
        nama_kegiatan_list = []
        for cek in obj.cek_set.all():
            ajuan = cek.ajuan
            if ajuan:
                nama_kegiatan_list.append(ajuan.nama_kegiatan)
        if nama_kegiatan_list:
            return ", ".join(nama_kegiatan_list)
        else:
            return '-'

    get_nama_kegiatan.short_description = 'Nama Kegiatan'

    def nomer_bank_tertarik(self, obj):
        nomer_bank_tertarik_list = []
        for cek in obj.cek_set.all():
            nomer_bank_tertarik_list.append(str(cek.nomer_bank_tertarik))
        if nomer_bank_tertarik_list:
            return ", ".join(nomer_bank_tertarik_list)
        else:
            return '-'

    def save_model(self, request, obj, form, change):
        total_ajuan = 0
        for cek in obj.cek_set.all():
            total_ajuan += cek.ajuan.total_ajuan
        obj.jumlah = total_ajuan
        super().save_model(request, obj, form, change)

    actions = ['export_as_pdf']

    def save_model(self, request, obj, form, change):
        if obj.pk:
            total_ajuan = sum(cek.ajuan.total_ajuan for cek in obj.cek_set.all() if cek.ajuan)
            obj.jumlah = total_ajuan
        super().save_model(request, obj, form, change)

    actions = ['export_as_pdf']


    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="RPC.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))
        elements = []

        # Add title
        title_style = styles['Heading1']
        title_style.alignment = TA_CENTER
        title = Paragraph('Daftar RPC' , style=title_style)
        elements.append(title)

        for rpc in queryset:
            # Add RPC number
            rpc_num = Paragraph('RPC No. ' + str(rpc.no_RPC), style=styles['Heading2'])
            elements.append(rpc_num)

            # Add table
            data = [['NO.', 'NAMA KEGIATAN', 'TOTAL AJUAN', 'NO CEK', 'NOMER BANK TERTARIK']]
            total = 0
            row_num = 1
            for cek in rpc.cek_set.all():
                ajuan = cek.ajuan
                row = [
                    row_num,
                    ajuan.nama_kegiatan if ajuan else 'Batal',
                    format_currency(cek.ajuan.total_ajuan, 'IDR', locale='id_ID') if ajuan else '',
                    cek.no_cek,
                    cek.nomer_bank_tertarik,
                ]
                data.append(row)
                row_num += 1
                if ajuan:
                    total += ajuan.total_ajuan

            data.append(['JUMLAH', '', format_currency(total, 'IDR', locale='id_ID')])

            # Create table
            table = Table(data,colWidths=[70, 250, 150, 130, 130])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # Set "NAMA KEGIATAN" in the second row to align left
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('WORDWRAP', (1, 0), (-1, -1), 100),
            ]))
            elements.append(table)

        space = Spacer(1, 10 * mm)
        elements.append(space)
        tanda_tangan = [['', '', '']]
        label_ttd = ['Disetujui', '                    ', 'Diperiksa', '                    ', 'Diajukan']
        tanda_tangan.append(label_ttd)
        space_ttd = ' '
        tanda_tangan.append(space_ttd)
        tanda_tangan.append(space_ttd)
        tanda_tangan.append(space_ttd)
        data_ttd = ['(H. Misbah Rosyadi)', '                    ', '(M. Dhiya Ulhaq, S.E)', '                    ',
                    '(Ita Aryani)']
        data_ttd_2 = ['Ketua Yayasan', '                    ', 'Direktur Keuangan', '                    ', 'Bendahara']
        tanda_tangan.append(data_ttd)
        tanda_tangan.append(data_ttd_2)
        table_ttd = Table(tanda_tangan)
        table_ttd.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        ]))
        elements.append(table_ttd)

        doc.build(elements)
        return response

    export_as_pdf.short_description = "Export selected as PDF"



admin.site.register(UnitAjuan)
admin.site.register(BankTertarik)
admin.site.register(Ajuan, AjuanAdmin)
admin.site.register(DanaMasuk, DanaMasukAdmin)
admin.site.register(BuktiKasKeluar, BuktiKasKeluarAdmin)
admin.site.register(RekapPencairanCek, RekapPencairanCekAdmin)
admin.site.register(RekapAjuanPengambilanTabungan, RekapAjuanPengambilanTabunganAdmin)
admin.site.register(Cek, CekAdmin)
admin.site.site_header = 'Dashboard Sistem Ajuan'
admin.site.site_title = 'Sistem Ajuan'
admin.site.index_title = 'Selamat datang di Dashboard Sistem Ajuan'