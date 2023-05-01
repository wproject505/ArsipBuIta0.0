from django.db import models
from django.db.models import Sum
from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver



class UnitAjuan(models.Model):
    unit_ajuan = models.CharField(null=False, max_length=10)
    def __str__(self):
        return self.unit_ajuan
    class Meta:
        verbose_name_plural = 'Unit Ajuan'


class BankTertarik(models.Model):
    nomer_bank_tertarik = models.CharField(null=False,max_length=25)
    def __str__(self):
        return self.nomer_bank_tertarik

    class Meta:
        verbose_name_plural = 'Bank Tertarik'

class DanaMasuk(models.Model):
    nama_dana_masuk = models.CharField(null=False, max_length=50)
    waktu_masuk = models.DateField(null=True)
    penanggung_jawab = models.CharField(null=True, max_length=30)
    total_dana = models.DecimalField(max_digits=20, decimal_places=0)
    # nomor_RAPT = models.ForeignKey(RekapAjuanPengambilanTabungan, null=True, blank=True, on_delete=models.SET_NULL)
    def __str__(self):
        return self.nama_dana_masuk

    class Meta:
        verbose_name_plural = 'Dana Masuk'


class RekapAjuanPengambilanTabungan(models.Model):
    no_RAPT = models.CharField(max_length=20, null=True, blank=True)
    jumlah = models.DecimalField(max_digits=50, decimal_places=0, null=True, blank=True)

    def __str__(self):
        return self.no_RAPT

    class Meta:
        verbose_name_plural = 'RAPT'

    def save(self, *args, **kwargs):
        if not self.no_RAPT:
            last_id = RekapAjuanPengambilanTabungan.objects.all().order_by('-pk').first()
            if last_id:
                no_RAPT = f'RAPT{str(last_id.pk + 1).zfill(5)}'
            else:
                no_RAPT = 'RAPT00001'
            self.no_RAPT = no_RAPT

        total_ajuan = Ajuan.objects.filter(RAPT=self).aggregate(total=Sum('total_ajuan'))['total']
        self.jumlah = total_ajuan or 0

        super().save(*args, **kwargs)





class RekapPencairanCek(models.Model):
    no_RPC = models.CharField(max_length=20, null=True, blank=True)
    jumlah = models.DecimalField(max_digits=50, decimal_places=0, null=True, blank=True)

    def __str__(self):
        return self.no_RPC

    class Meta:
        verbose_name_plural = 'RPC'

    def save(self, *args, **kwargs):
        if not self.no_RPC:
            last_id = RekapPencairanCek.objects.all().order_by('-pk').first()
            if last_id:
                no_RPC = f'RPC{str(last_id.pk + 1).zfill(5)}'
            else:
                no_RPC = 'RPC00001'
            self.no_RPC = no_RPC

        super().save(*args, **kwargs)



class Ajuan(models.Model):
    unit_ajuan = models.ForeignKey(UnitAjuan, null=True, on_delete=models.SET_NULL)
    nomor_pengajuan = models.CharField(max_length=50,  null=True, blank=True, help_text="nomor pengajuan akan terisi otomatis")
    nama_kegiatan = models.CharField(null=False, blank=True, max_length=50)
    waktu_ajuan = models.DateField(blank=True, null=True)
    penanggung_jawab = models.CharField(null=True, blank=True, max_length=30)
    total_ajuan = models.DecimalField(max_digits=20, blank=False, null=False, decimal_places=0)
    RAPT = models.ForeignKey(RekapAjuanPengambilanTabungan, null=True, blank=True, on_delete=models.SET_NULL)

    def __str__(self):
        return self.nomor_pengajuan

    class Meta:
        verbose_name_plural = 'Ajuan'


# @receiver(post_save, sender=Ajuan)
# def create_rapt(sender, instance, created, **kwargs):
#     if created:
#         rapt = RekapAjuanPengambilanTabungan.objects.create(
#             no_RAPT=instance.nomor_pengajuan,
#             jumlah=instance.total_ajuan,
#         )
#         instance.RAPT = rapt
#         instance.save()

class Cek(models.Model):
    no_cek  = models.CharField(max_length=50, null=True,)
    keterangan = models.CharField(max_length=50, null=True, )
    nomer_bank_tertarik = models.ForeignKey(BankTertarik, null=True, on_delete=models.SET_NULL)
    ajuan = models.ForeignKey(Ajuan, null=True, blank=True, on_delete=models.SET_NULL)
    total_ajuan = models.DecimalField(max_digits=20, blank=False, null=False, decimal_places=0)
    RPC = models.ForeignKey(RekapPencairanCek, null=True, blank=True, on_delete=models.SET_NULL)



    def __str__(self):
        return self.no_cek

    class Meta:
        verbose_name_plural = 'Cek'




class BuktiKasKeluar(models.Model):
    no_BKK = models.CharField(max_length=50, blank=True, help_text="nomor BKK akan terisi otomatis")
    tanggal_BKK = models.DateField(null=True)
    ajuan = models.ForeignKey(Ajuan, null=True, blank=True, on_delete=models.SET_NULL)
    dibayarkan_kepada = models.CharField(null=False, max_length=20)
    uraian = models.CharField(null=False, max_length=50)
    nomer_bank_tertarik = models.ForeignKey(BankTertarik, null=True, on_delete=models.SET_NULL)
    nomer_cek = models.ForeignKey(Cek, null=True, blank=True, on_delete=models.SET_NULL)
    def __str__(self):
        return self.no_BKK

    class Meta:
        verbose_name_plural = 'Bukti Kas Keluar'

    def save(self, *args, **kwargs):
        last_id = BuktiKasKeluar.objects.order_by('-pk').first()
        if last_id:
            id_num = str(last_id.pk + 1).zfill(4)
        else:
            id_num = '0000'
        self.no_BKK = f'BK{id_num}'

        no_BKK = f'BKK{id_num}'
        count = 0
        while BuktiKasKeluar.objects.filter(no_BKK=no_BKK).exists():
            count += 1
            no_BKK = f'BKK{id_num}({count})'

        super().save(*args, **kwargs)

