from django.contrib import admin
from .models import *
from django.contrib.admin.widgets import AdminTextInputWidget
from django import forms



class AjuanAdmin(admin.ModelAdmin):
    raw_id_fields = ["RAPT",
                     "RPC",
                     ]


class BuktiKasKeluarAdmin(admin.ModelAdmin):
    raw_id_fields = ["ajuan"]

class AjuanInLineRAPT(admin.TabularInline):
    model = Ajuan
    extra = 0
    fields = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab']
    readonly_fields = fields


class RekapAjuanPengambilanTabunganAdmin(admin.ModelAdmin):
    fieldsets = [
        (None, {'fields': ['no_RAPT']}),
        (None, {'fields': ['jumlah']}),
    ]
    inlines = [AjuanInLineRAPT]


class AjuanInLineRPC(admin.TabularInline):
    model = Ajuan
    extra = 0
    fields = ['unit_ajuan','nomor_pengajuan',  'nama_kegiatan', 'waktu_ajuan', 'total_ajuan', 'penanggung_jawab']
    readonly_fields = fields


class RekapPencairanCekAdmin(admin.ModelAdmin):
    fieldsets = [
        (None, {'fields': ['no_RPC']}),
        (None, {'fields': ['jumlah']}),
    ]
    inlines = [AjuanInLineRPC]





admin.site.register(UnitAjuan)
admin.site.register(Ajuan, AjuanAdmin)
admin.site.register(DanaMasuk)
admin.site.register(BuktiKasKeluar, BuktiKasKeluarAdmin)
admin.site.register(RekapPencairanCek, RekapPencairanCekAdmin)
admin.site.register(RekapAjuanPengambilanTabungan, RekapAjuanPengambilanTabunganAdmin)