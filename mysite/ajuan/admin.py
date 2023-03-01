from .models import *
from django.http import HttpResponse
from django.contrib import admin
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from .models import BuktiKasKeluar
from babel.numbers import format_currency
from decimal import Decimal


class DanaMasukAdmin(admin.ModelAdmin):
    search_fields = ('nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana')
    list_display = ('nama_dana_masuk', 'waktu_masuk', 'penanggung_jawab', 'total_dana')



class BuktiKasKeluarAdmin(admin.ModelAdmin):
    search_fields = ('no_BKK', 'tanggal_BKK', 'dibayarkan_kepada', 'uraian', 'kode_bank', 'nomer_cek')
    list_display = ('no_BKK', 'tanggal_BKK', 'ajuan', 'dibayarkan_kepada', 'uraian', 'kode_bank', 'nomer_cek')
    actions = ["export_as_pdf"]
    raw_id_fields = ['ajuan']

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
    actions = ['export_as_pdf']

    # def get_queryset(self, request):
    #     queryset = super().get_queryset(request)
    #     total_ajuan_accumulated = queryset.aggregate(total_ajuan_accumulated=Sum('total_ajuan'))['total_ajuan_accumulated']
    #     self.total_ajuan_accumulated = total_ajuan_accumulated or 0
    #     return queryset
    #
    # def total_ajuan_accumulated(self, obj):
    #     return format_html('<b>{}</b>', obj.total_ajuan_accumulated)
    # total_ajuan_accumulated.short_description = 'Akumulasi Total Ajuan'


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

    actions = ['export_as_pdf']

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