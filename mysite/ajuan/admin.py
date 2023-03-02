from .models import *
from django.http import HttpResponse
from django.contrib import admin
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from .models import BuktiKasKeluar
from babel.numbers import format_currency
from decimal import Decimal
from reportlab.lib.pagesizes import letter
from openpyxl import Workbook


class DanaMasukAdmin(admin.ModelAdmin):
    search_fields = ('nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana')
    list_display = ('nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana')
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
        response['Content-Disposition'] = 'attachment; filename="Dana_masuk.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))

        data = []
        data.append(['nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana'])
        total = Decimal(0)
        for dana_masuk in queryset:
            total_ajuan = Decimal(dana_masuk.total_dana)
            row = [
                dana_masuk.nama_dana_masuk,
                dana_masuk.waktu_masuk,
                dana_masuk.penanggung_jawab,
                dana_masuk.total_dana,
            ]
            data.append(row)

            # Add total_ajuan to total
            total += total_ajuan

            # Add row for total_ajuan
        data.append(['', '', '', '', '', 'Total', format_currency(total, 'IDR', locale='id_ID')])

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


class BuktiKasKeluarAdmin(admin.ModelAdmin):
    search_fields = ('no_BKK', 'tanggal_BKK', 'dibayarkan_kepada', 'uraian', 'kode_bank', 'nomer_cek')
    list_display = ('no_BKK', 'tanggal_BKK', 'ajuan', 'dibayarkan_kepada', 'uraian', 'kode_bank', 'nomer_cek')
    actions = ["export_as_pdf","export_to_excel"]
    raw_id_fields = ['ajuan']


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


    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="bukti_kas_keluar.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))

        data = []
        data.append(['Nomor BKK', 'Tanggal BKK', 'Ajuan', 'Dibayarkan Kepada', 'Uraian', 'Kode Bank', 'Nomor Cek'])
        total = Decimal(0)
        for bukti_kas_keluar in queryset:
            total_ajuan = Decimal(bukti_kas_keluar.ajuan.total_ajuan)
            row = [
                bukti_kas_keluar.no_BKK,
                bukti_kas_keluar.tanggal_BKK,
                total_ajuan,
                bukti_kas_keluar.dibayarkan_kepada,
                bukti_kas_keluar.uraian,
                bukti_kas_keluar.kode_bank,
                bukti_kas_keluar.nomer_cek,
            ]
            data.append(row)

            # Add total_ajuan to total
            total += total_ajuan

            # Add row for total_ajuan
        data.append(['', '', '', '', '', 'Total', format_currency(total, 'IDR', locale='id_ID')])



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


class AjuanAdmin(admin.ModelAdmin):
    search_fields = ('nomor_pengajuan', 'nama_kegiatan', 'waktu_ajuan', 'total_ajuan')
    list_display = ['unit_ajuan','nomor_pengajuan','nama_kegiatan','waktu_ajuan','total_ajuan','RAPT','RPC']
    raw_id_fields = ["RAPT",
                     "RPC",
                     ]
    actions = ['export_as_pdf','export_to_excel']

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

        data = []
        data.append(['unit_ajuan', 'nomor_pengajuan', 'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'RAPT', 'RPC'])
        total = Decimal(0)
        for ajuan in queryset:
            total_ajuan = Decimal(ajuan.total_ajuan)
            row = [
                ajuan.unit_ajuan,
                ajuan.nomor_pengajuan,
                ajuan.nama_kegiatan,
                ajuan.waktu_ajuan,
                ajuan.total_ajuan,
                ajuan.RAPT,
                ajuan.RPC,
            ]
            data.append(row)

            # Add total_ajuan to total
            total += total_ajuan

            # Add row for total_ajuan
        data.append(['', '', '', '', '', 'Total', format_currency(total, 'IDR', locale='id_ID')])


        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 14),
            ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))


        elements = []
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
    list_display = ('no_RAPT', 'jumlah')
    inlines = [AjuanInLineRAPT]

    fieldsets = [
        (None, {'fields': ['no_RAPT']}),
        (None, {'fields': ['jumlah']}),
    ]

    actions = ["export_as_pdf", "export_to_excel"]

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

class AjuanInLineRPC(admin.TabularInline):
    model = Ajuan
    extra = 0
    fields = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab']
    readonly_fields = fields


class RekapPencairanCekAdmin(admin.ModelAdmin):
    search_fields = ('no_RPC', 'jumlah')
    list_display = ('no_RPC', 'jumlah')
    inlines = [AjuanInLineRPC]

    fieldsets = [
        (None, {'fields': ['no_RPC']}),
        (None, {'fields': ['jumlah']}),
    ]

    actions = ['export_as_pdf']


    def export_to_excel(self, request, queryset):
        # Query data dari model Dana Masuk
        data = queryset.values()

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['no_RPC', 'jumlah'])

        # Tambahkan data ke file Excel
        for item in data:
            row = [item['no_RPC'], item['jumlah']]
            ws.append(row)

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=RPC.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'

    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="RPC.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))

        data = []
        data.append(['no_RPC', 'jumlah'])
        total = Decimal(0)
        for RPC in queryset:
            row = [
                RPC.no_RPC,
                RPC.jumlah
            ]
            data.append(row)

            # Add total_ajuan to total
            total += RPC.jumlah

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




admin.site.register(UnitAjuan)
admin.site.register(Ajuan, AjuanAdmin)
admin.site.register(DanaMasuk, DanaMasukAdmin)
admin.site.register(BuktiKasKeluar, BuktiKasKeluarAdmin)
admin.site.register(RekapPencairanCek, RekapPencairanCekAdmin)
admin.site.register(RekapAjuanPengambilanTabungan, RekapAjuanPengambilanTabunganAdmin)