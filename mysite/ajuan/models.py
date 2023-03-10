from django.db import models


class UnitAjuan(models.Model):
    unit_ajuan = models.CharField(null=False, max_length=10)
    def __str__(self):
        return self.unit_ajuan


class DanaMasuk(models.Model):
    nama_dana_masuk = models.CharField(null=False, max_length=50)
    waktu_masuk = models.DateField(null=True)
    penanggung_jawab = models.CharField(null=True, max_length=30)
    total_dana = models.DecimalField(max_digits=20, decimal_places=0)
    # nomor_RAPT = models.ForeignKey(RekapAjuanPengambilanTabungan, null=True, blank=True, on_delete=models.SET_NULL)
    def __str__(self):
        return self.nama_dana_masuk


class RekapAjuanPengambilanTabungan(models.Model):
    no_RAPT = models.CharField(max_length=20, unique=True, null=True, blank=True)
    jumlah = models.DecimalField(max_digits=50, decimal_places=0, null=True, blank=True)

    def __str__(self):
        return self.no_RAPT

    def save(self, *args, **kwargs):
        if not self.no_RAPT:
            last_id = RekapAjuanPengambilanTabungan.objects.all().order_by('-id').first()
            if last_id:
                id_num = str(last_id.id + 1).zfill(4)
            else:
                id_num = '01'
            self.no_RAPT = f'RAPT{id_num}'

        total_ajuan = Ajuan.objects.filter(RAPT=self).aggregate(total=Sum('total_ajuan'))['total']
        self.jumlah = total_ajuan or 0

        super().save(*args, **kwargs)



from django.db.models import Sum

class RekapPencairanCek(models.Model):
    no_RPC = models.CharField(max_length=20, unique=True, null=True, blank=True)
    jumlah = models.DecimalField(max_digits=50, decimal_places=0, null=True, blank=True)

    def __str__(self):
        return self.no_RPC

    def save(self, *args, **kwargs):
        if not self.no_RPC:
            last_id = RekapPencairanCek.objects.all().order_by('-pk').first()
            if last_id:
                id_num = str(last_id.pk + 1).zfill(4)
            else:
                id_num = '0001'
            self.no_RPC = f'RPC{id_num}'

        total_ajuan = Ajuan.objects.filter(RPC=self).aggregate(total=Sum('total_ajuan'))['total']
        self.jumlah = total_ajuan or 0

        super().save(*args, **kwargs)


class Ajuan(models.Model):
    unit_ajuan = models.ForeignKey(UnitAjuan, null=True, on_delete=models.SET_NULL)
    nomor_pengajuan = models.CharField(max_length=50, unique=True, null=True, blank=True, help_text="nomor pengajuan akan terisi otomatis")
    nama_kegiatan = models.CharField(null=False, blank=True, max_length=50)
    waktu_ajuan = models.DateField(blank=True, null=True)
    penanggung_jawab = models.CharField(null=True, blank=True, max_length=30)
    total_ajuan = models.DecimalField(max_digits=20, blank=False, null=False, decimal_places=0)
    RAPT = models.ForeignKey(RekapAjuanPengambilanTabungan, null=True, blank=True, on_delete=models.SET_NULL)
    RPC = models.ForeignKey(RekapPencairanCek, null=True, blank=True, on_delete=models.SET_NULL)


    def save(self, *args, **kwargs):
        if not self.nomor_pengajuan:
            last_id = Ajuan.objects.all().order_by('-pk').first()
            if last_id:
                id_num = str(last_id.pk + 1).zfill(4)
            else:
                id_num = '0001'
            unit_ajuan = self.unit_ajuan.unit_ajuan
            nomor_pengajuan = f'{unit_ajuan}{id_num}'
            count = 1
            while Ajuan.objects.filter(nomor_pengajuan=nomor_pengajuan).exists():
                count += 1
                nomor_pengajuan = f'{unit_ajuan}{id_num}({count})'
            self.nomor_pengajuan = nomor_pengajuan


        if self.RAPT:
            total_ajuan = Ajuan.objects.filter(RAPT=self.RAPT).aggregate(total=Sum('total_ajuan'))['total']
            self.RAPT.jumlah = total_ajuan or 0  # default value is 0 if total_ajuan is None
            self.RAPT.save()

        if self.RPC:
            total_ajuan = Ajuan.objects.filter(RPC=self.RPC).aggregate(total=Sum('total_ajuan'))['total']
            self.RPC.jumlah = total_ajuan or 0  # default value is 0 if total_ajuan is None
            self.RPC.save()

        super(Ajuan, self).save(*args, **kwargs)

    def __str__(self):
        return self.nomor_pengajuan


class BuktiKasKeluar(models.Model):
    no_BKK = models.CharField(max_length=50, blank=True, help_text="nomor BKK akan terisi otomatis")
    tanggal_BKK = models.DateField(null=True)
    ajuan = models.ForeignKey(Ajuan, null=True, blank=True, on_delete=models.SET_NULL)
    dibayarkan_kepada = models.CharField(null=False, max_length=20)
    uraian = models.CharField(null=False, max_length=50)
    kode_bank = models.CharField(null=False, max_length=20)
    nomer_cek = models.CharField(null=False, max_length=20)
    def __str__(self):
        return self.no_BKK

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


