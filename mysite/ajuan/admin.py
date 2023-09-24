import babel
from .models import *
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from .models import BuktiKasKeluar
from babel.numbers import format_currency
from decimal import Decimal
from openpyxl import Workbook
from django.contrib import admin
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import mm, cm
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl.styles import Alignment
import babel.numbers
import decimal
from django.contrib.admin.widgets import FilteredSelectMultiple
import os
from django.conf import settings
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.units import cm
from babel.numbers import format_currency
from reportlab.lib.units import inch
from terbilang import Terbilang



styles = getSampleStyleSheet()
class DanaMasukAdmin(admin.ModelAdmin):

    search_fields = ['waktu_masuk', 'uraian', 'bank_penerima', 'total_dana']
    list_display = ('waktu_masuk', 'uraian', 'bank_penerima', 'get_total_dana')
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
        data.append(['No.','Uraian', 'Waktu Masuk', 'Total Dana'])
        total = Decimal(0)
        row_num = 1
        for dana_masuk in queryset:
            total_ajuan = Decimal(dana_masuk.total_dana)
            row = [
                row_num,
                dana_masuk.uraian,
                dana_masuk.waktu_masuk,
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
    readonly_fields = ['no_BKK', ]
    fields = ['no_BKK', 'tanggal_BKK', 'ajuan', 'dibayarkan_kepada', 'uraian', 'nomer_cek','nomer_bank_tertarik']
    search_fields = ['no_BKK','ajuan__nomor_pengajuan', 'tanggal_BKK', 'dibayarkan_kepada', 'uraian','nomer_bank_tertarik__nomer_bank_tertarik','nomer_cek__no_cek']
    list_display = ('no_BKK', 'tanggal_BKK', 'ajuan','get_total_ajuan', 'dibayarkan_kepada', 'uraian', 'nomer_bank_tertarik', 'nomer_cek')
    actions = ["export_to_excel", "export_as_pdf_global", "export_to_pdf_satuan","test_add_logo"]
    raw_id_fields = ['ajuan', 'nomer_cek']


    def set_table_style(self, table):
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.white),  # Mengubah warna latar belakang header menjadi putih
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('TOPPADDING', (0, 0), (-1, 0), 2.5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),  # Mengubah warna latar belakang sel lainnya menjadi putih
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))

    def generate_first_table(self, elements, data):
        def wrap_text_by_words(text, limit=7):
            words = text.split()
            return '\n'.join([' '.join(words[i:i + limit]) for i in range(0, len(words), limit)])

        logo_path = os.path.join(settings.STATIC_ROOT, 'images/Sistem Ajuan Bu Ita.png')
        logo = Image(logo_path)
        logo.drawHeight = 1.5 * cm
        logo.drawWidth = 2 * cm

        total_ajuan = data.ajuan.total_ajuan if data and data.ajuan and data.ajuan.total_ajuan else 0
        total_ajuan_str = str(total_ajuan)
        t = Terbilang()
        t.parse(total_ajuan_str)
        t_gr = t.getresult()
        t_gr_string_title = t_gr.title() + ' Rupiah'
        t_gr_string_title = wrap_text_by_words(t_gr_string_title, 7)


        data_0 = [
            [logo, '\n BUKTI KAS KELUAR\n_________________',
             'Nomer BKK: {}\nTanggal: {}'.format(data.no_BKK, data.tanggal_BKK) if data else ''],
            ['Perkiraan', 'Uraian', 'Jumlah'],
            ['', '', ''],
            ['', '', ''],
            ['Dibayarkan Kepada:\n {}'.format(data.dibayarkan_kepada), data.uraian,
             format_currency(data.ajuan.total_ajuan, 'IDR', locale='id_ID') if data.ajuan.total_ajuan else ''],
            ['Terbilang', '{}'.format(t_gr_string_title), '']
        ]

        table_0 = Table(data_0, colWidths=[150, 250, 150])
        self.set_table_style(table_0)
        elements.append(table_0)


    def generate_second_table(self, elements, data):
        data_2 = [
            ['Nomer Bank Tertarik: {}'.format(
                data.nomer_bank_tertarik.nomer_bank_tertarik if data.nomer_bank_tertarik else ''),
             'Nomer Cek: {}'.format(data.nomer_cek.no_cek if data.nomer_cek else '')]
        ]
        table_2 = Table(data_2, colWidths=[275, 275])
        self.set_table_style(table_2)
        elements.append(table_2)

    def generate_signature_table(self, elements, data):
        tanda_tangan = [['Pemberi \n \n \n____________', 'Mengetahui \n \n \n____________',
                         'Penerima \n \n \n____________']]
        table_ttd = Table(tanda_tangan, colWidths=[175, 200, 175])
        self.set_table_style(table_ttd)
        elements.append(table_ttd)



    def export_to_pdf_satuan(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="BKK Satuan.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(A4), topMargin=0.5*inch, bottomMargin=0.1*inch)


        elements = []
        for data in queryset:
            # Kode untuk tabel pertama
            self.generate_first_table(elements, data)

            # Kode untuk tabel kedua
            self.generate_second_table(elements, data)

            # Kode untuk tanda tangan
            self.generate_signature_table(elements, data)

            # Kode untuk spacer
            data_spacer = [('')]
            table_spacer = Table(data_spacer, colWidths=[500])
            elements.append(table_spacer)

        doc.build(elements)

        return response

    export_to_pdf_satuan.short_description = "Export selected as PDF satuan"

    def get_total_ajuan(self, obj):
        if hasattr(obj, 'ajuan') and hasattr(obj.ajuan, 'total_ajuan'):
            total_ajuan = decimal.Decimal(str(obj.ajuan.total_ajuan))
        else:
            total_ajuan = decimal.Decimal('0')
        return babel.numbers.format_currency(total_ajuan, 'IDR', locale='id_ID')
    get_total_ajuan.short_description = 'Total Ajuan'

    def export_to_excel(self, request, queryset):
        # Query data dari model BuktiKasKeluar
        data = queryset.values('no_BKK', 'tanggal_BKK', 'ajuan_id', 'dibayarkan_kepada', 'uraian',
                               'nomer_bank_tertarik', 'nomer_cek')

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['No.', 'No BKK', 'Tanggal BKK', 'Ajuan', 'Dibayarkan Kepada', 'Uraian', 'Nomer Bank Tertarik',
                   'Nomor Cek'])
        row_num = 1

        # Tambahkan data ke file Excel
        for item in data:
            row = [row_num, item['no_BKK'], item['tanggal_BKK'], item['ajuan_id'], item['dibayarkan_kepada'],
                   item['uraian'], item['nomer_bank_tertarik'], item['nomer_cek']]
            ws.append(row)
            row_num += 1

        # Konversi file Excel ke HttpResponse
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=BuktiKasKeluar.xlsx'
        wb.save(response)

        return response

    export_to_excel.short_description = 'Export to Excel'

    def export_as_pdf_global(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="BKK Global.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))
        elements = []
        title_style = styles['Heading1']
        title_style.alignment = TA_CENTER
        title = Paragraph('Bukti Kas Keluar', style=title_style)
        elements.append(title)
        data = []
        data.append(
            ['No.', 'Nomor BKK', 'Tanggal BKK', 'Nomor Cek', 'Dibayarkan Kepada', 'Uraian', 'Kode Bank', 'Ajuan'])
        total = Decimal(0)
        row_num = 1
        for bukti_kas_keluar in queryset:
            total_ajuan = Decimal(bukti_kas_keluar.ajuan.total_ajuan) if bukti_kas_keluar.ajuan else Decimal(0)
            row = [
                row_num,
                bukti_kas_keluar.no_BKK,
                bukti_kas_keluar.tanggal_BKK,
                bukti_kas_keluar.dibayarkan_kepada,
                bukti_kas_keluar.uraian,
                bukti_kas_keluar.nomer_bank_tertarik,
                bukti_kas_keluar.nomer_cek,
                format_currency(total_ajuan, 'IDR', locale='id_ID') if total_ajuan else '-',
            ]
            data.append(row)
            row_num += 1

            # Add total_ajuan to total
            total += total_ajuan

        # Add row for total_ajuan
        data.append(['', '', '', '', '', '', 'Total', format_currency(total, 'IDR', locale='id_ID')])
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 2), (-1, -1), 'LEFT'),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            ('ALIGN', (0, 0), (7, 0), 'CENTER'),
            ('ALIGN', (7, 1), (7, -1), 'RIGHT'),
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
    list_display = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'get_total_ajuan', 'penanggung_jawab', 'RAPT',  ]
    fields = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab', 'RAPT',  ]
    readonly_fields = ['nomor_pengajuan', ]
    raw_id_fields = ['RAPT',]
    actions = ['export_as_pdf','export_to_excel']


    def get_total_ajuan(self, obj):
        return babel.numbers.format_currency(obj.total_ajuan, 'IDR', locale='id_ID')
    get_total_ajuan.short_description = 'Total Ajuan'

    def export_to_excel(self, request, queryset):
        # Query data dari model BuktiKasKeluar
        data = queryset.values('unit_ajuan__unit_ajuan', 'nomor_pengajuan', 'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'RAPT')

        # Buat file Excel dan tambahkan header
        wb = Workbook()
        ws = wb.active
        ws.append(['No.','Unit Ajuan','Nomor Pengajuan','Nama Kegiatan','Waktu Ajuan','Total Ajuan','RAPT'])
        # set rata tengah pada header
        header_row = ws[1]
        for cell in header_row:
            cell.alignment = Alignment(horizontal='center')

        row_num = 1
        # Tambahkan data ke file Excel
        for item in data:
            row = [row_num, item['unit_ajuan__unit_ajuan'], item['nomor_pengajuan'], item['nama_kegiatan'], item['waktu_ajuan'], item['total_ajuan'],
                   item['RAPT']]
            ws.append(row)
            row_num += 1

            # Menyesuaikan lebar kolom secara otomatis
            for column_cells in ws.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = length

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
        data.append(['', '', '', '','Total', format_currency(total, 'IDR', locale='id_ID', )])


        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (3, 1), (3, -2), 'LEFT'),
            ('ALIGN', (2, 1), (2, -2), 'LEFT'),
            ('ALIGN', (1, 1), (1, -2), 'LEFT'),
            ('ALIGN', (5, 1), (5, -2), 'RIGHT'),
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
    list_display = ('no_RAPT','get_nomor_pengajuan','get_total_ajuan', 'get_jumlah')
    readonly_fields = ('no_RAPT','jumlah',)
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

    get_total_ajuan.short_description = 'total ajuan'

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
from django.db.models import Q
class CekAdmin(admin.ModelAdmin):
    raw_id_fields = ('RPC',)
    search_fields = ('no_cek', 'nomer_bank_tertarik__nomer_bank_tertarik', 'RPC__no_RPC',)
    list_display = ('tanggal', 'no_cek', 'nomer_bank_tertarik', 'display_many_to_many', 'total_cek', 'RPC',)
    actions = ["export_as_pdf"]
    formfield_overrides = {
        models.ManyToManyField: {'widget': FilteredSelectMultiple('Ajuan', False)},
    }
    def get_form(self, request, obj=None, **kwargs):
        request._obj_ = obj
        return super(CekAdmin, self).get_form(request, obj, **kwargs)
    def formfield_for_manytomany(self, db_field, request, **kwargs):
        if db_field.name == "ajuan_terkait":
            if hasattr(request, '_obj_') and request._obj_ is not None:
                kwargs['queryset'] = Ajuan.objects.filter(
                    Q(is_selected=False) | Q(ceks_terkait=request._obj_)
                )
            else:
                kwargs['queryset'] = Ajuan.objects.filter(is_selected=False)
        return super().formfield_for_manytomany(db_field, request, **kwargs)



    def display_many_to_many(self, obj):
        # Ambil nomor_pengajuan dari setiap objek Ajuan yang terkait dalam ajuan_terkait
        nomor_pengajuan_list = [str(ajuan.nomor_pengajuan) for ajuan in obj.ajuan_terkait.all()]

        # Menggabungkan nomor_pengajuan menjadi satu string dengan koma dan spasi sebagai pemisah
        return ', '.join(nomor_pengajuan_list)
    display_many_to_many.short_description = 'Ajuan Terkait'

    def export_as_pdf(self, request, queryset):
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="Cek.pdf"'

        doc = SimpleDocTemplate(response, pagesize=landscape(letter))
        elements = []

        styles = getSampleStyleSheet()
        title_style = styles['Heading1']
        title_style.alignment = 1  # TA_CENTER is 1
        title = Paragraph('Daftar Cek', style=title_style)
        elements.append(title)

        row_num = 1
        data = [['No.', 'Tanggal', 'No.Cek', 'Nomer Bank Tertarik','Ajuan', 'RPC', 'Total Ajuan']]

        total_sum = 0  # Inisialisasi total

        for cek in queryset:
            if cek.total_cek is not None:
                formatted_total_cek = format_currency(cek.total_cek, 'IDR', locale='id_ID')
            else:
                formatted_total_cek = "-"  # Or any string you want to display when total_cek is None
            ajuans = ", ".join(str(ajuan.nomor_pengajuan) for ajuan in cek.ajuan_terkait.all())
            total_ajuan = cek.ajuan_terkait.aggregate(Sum('total_ajuan'))[
                'total_ajuan__sum']  # asumsikan total_ajuan adalah field dalam model Ajuan

            # Menambahkan 'cek.tanggal' dan 'total_ajuan' ke baris tabel
            row = [
                row_num,
                cek.tanggal,
                cek.no_cek,
                cek.nomer_bank_tertarik,
                ajuans,
                cek.RPC,
                formatted_total_cek,  # Total Ajuan

            ]
            data.append(row)
            row_num += 1

            if cek.total_cek is not None:
                total_sum += cek.total_cek

        total_row = [
            'Total', '', '', '', '', '', format_currency(total_sum, 'IDR', locale='id_ID')
        ]
        data.append(total_row)

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('ALIGN', (6, 1), (6, -1), 'RIGHT'),
            ('ALIGN', (0, 1), (1, -2), 'CENTER'),
            ('ALIGN', (0, 0), (6, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-0, 0), 'Helvetica-Bold'),
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
    fields = ['no_cek', 'nomer_bank_tertarik','get_nama_kegiatan',]
    readonly_fields = fields

    # def get_total_ajuan(self, obj):
    #     return obj.ajuan.total_ajuan
    #
    # get_total_ajuan.short_description = 'Total Ajuan'

    def get_nama_kegiatan(self, obj):
        return obj.ajuan.nama_kegiatan

    get_total_nama_kegiatan_description = 'Nama Kegiatan'


class RekapPencairanCekAdmin(admin.ModelAdmin):
    search_fields = ('no_RPC',)
    readonly_fields = ('no_RPC', 'jumlah',)
    inlines = [CekInLineRPC]
    # list_display = ('no_RPC', 'jumlah', 'nomer_bank_tertarik', 'get_total_ajuan', 'get_nama_kegiatan')

    actions = ['export_as_pdf', 'update_jumlah_RPC']

    def get_jumlah(self, obj):
        return babel.numbers.format_currency(obj.jumlah, 'IDR', locale='id_ID')
    get_jumlah.short_description = 'Jumlah'
    def update_jumlah_RPC(self, request, queryset):
        for obj in queryset:
            total_ajuan = sum(cek.ajuan.total_ajuan for cek in obj.cek_set.all() if cek.ajuan)
            obj.jumlah = total_ajuan
            obj.save(update_fields=['jumlah'])
        self.message_user(request, f'Successfully updated {queryset.count()} Rekap Pencairan Cek(s).')

    update_jumlah_RPC.short_description = 'Update Jumlah RPC'

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



    def save_model(self, request, obj, form, change):
        if obj.pk:
            total_ajuan = sum(cek.ajuan.total_ajuan for cek in obj.cek_set.all() if cek.ajuan)
            obj.jumlah = total_ajuan
        super().save_model(request, obj, form, change)



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

class RekapBankTertarikAdmin(admin.ModelAdmin):
    search_fields = ('no_RBT',)
    raw_id_fields = ['dana_masuk','no_cek' ]
    # readonly_fields = ('no_RBT', 'tanggal','uraian',)

    def display_ajuan(self, obj):
        return ', '.join([ajuan.nomor_pengajuan for ajuan in obj.ajuan.all()])
    list_display = ('no_RBT', 'tanggal','no_cek', 'dana_masuk','dana_keluar')
    formfield_overrides = {
        models.ManyToManyField: {'widget': FilteredSelectMultiple('Sumber', False)},
    }


class BuktiKasKeluarInLineBankTertarik(admin.TabularInline):
    model = BuktiKasKeluar
    extra = 0
    fields = ['no_BKK','tanggal_BKK',  'ajuan', 'dibayarkan_kepada', 'uraian', 'nomer_cek', 'nomer_bank_tertarik']
    readonly_fields = fields

class BankTertarikAdmin(admin.ModelAdmin):
    inlines = [BuktiKasKeluarInLineBankTertarik]


admin.site.register(UnitAjuan)
admin.site.register(BankTertarik, BankTertarikAdmin)
admin.site.register(Ajuan, AjuanAdmin)
admin.site.register(DanaMasuk, DanaMasukAdmin)
admin.site.register(BuktiKasKeluar, BuktiKasKeluarAdmin)
admin.site.register(RekapPencairanCek, RekapPencairanCekAdmin)
admin.site.register(RekapAjuanPengambilanTabungan, RekapAjuanPengambilanTabunganAdmin)
admin.site.register(Cek, CekAdmin)
admin.site.register(RekapBankTertarik, RekapBankTertarikAdmin)
admin.site.site_header = 'Sistem Ajuan Bu Ita'
admin.site.site_title = 'Selamat Datang di Sistem Ajuan Bu Ita'
admin.site.index_title = 'Selamat datang di Dashboard Sistem Ajuan Bu Ita'