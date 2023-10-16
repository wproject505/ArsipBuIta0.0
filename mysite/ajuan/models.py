from django.db import models
from django.db.models import Sum
from django.db.models.signals import pre_save
from django.db.models.signals import m2m_changed
from django.dispatch import receiver
class UnitAjuan(models.Model):
    unit_ajuan = models.CharField(null=False, max_length=30)
    def __str__(self):
        return self.unit_ajuan
    class Meta:
        verbose_name_plural = 'Unit Ajuan'


class BankTertarik(models.Model):
    nomer_bank_tertarik = models.CharField(null=False,max_length=25)
    def __str__(self):
        return self.nomer_bank_tertarik

    class Meta:
        verbose_name_plural = 'Bank'

class DanaMasuk(models.Model):
    waktu_masuk = models.DateField(null=True)
    uraian = models.CharField(null=False, max_length=50)
    bank_penerima = models.ForeignKey(BankTertarik, blank=True, null=True,on_delete=models.SET_NULL)
    total_dana = models.DecimalField(max_digits=20, decimal_places=0)
    def __str__(self):
        return self.uraian

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
                no_RAPT = 'RAPT0001'
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
    nama_kegiatan = models.TextField(null=True, blank=True)
    waktu_ajuan = models.DateField(blank=True, null=True)
    penanggung_jawab = models.CharField(null=True, blank=True, max_length=30)
    total_ajuan = models.DecimalField(max_digits=20, blank=False, null=True, decimal_places=0)
    RAPT = models.ForeignKey(RekapAjuanPengambilanTabungan, null=True, blank=True, on_delete=models.SET_NULL)
    # Field status untuk menandai apakah Ajuan sudah dipilih atau belum
    is_selected = models.BooleanField(default=False)


    def __str__(self):
        if self.nomor_pengajuan:
            return self.nomor_pengajuan
        return str(self.id)  # mengembalikan id sebagai fallback


    class Meta:
        verbose_name_plural = 'Ajuan'


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


        super(Ajuan, self).save(*args, **kwargs)







class Cek(models.Model):
    tanggal = models.DateField(null=True)
    no_cek = models.CharField(max_length=50, null=True)
    nomer_bank_tertarik = models.ForeignKey(BankTertarik, null=True, on_delete=models.SET_NULL)
    ajuan_terkait = models.ManyToManyField(Ajuan, blank=True, related_name='ceks_terkait')
    RPC = models.ForeignKey(RekapPencairanCek, null=True, blank=True, on_delete=models.SET_NULL)
    total_cek = models.DecimalField(max_digits=20, blank=True, null=True, decimal_places=0)

    def __str__(self):
        return self.no_cek

    class Meta:
        verbose_name_plural = 'Cek'


@receiver(m2m_changed, sender=Cek.ajuan_terkait.through)
def update_ajuan_and_total_cek(sender, instance, action, pk_set, **kwargs):
    if action == 'post_add':
        Ajuan.objects.filter(pk__in=pk_set).update(is_selected=True)
    elif action == 'post_remove':
        Ajuan.objects.filter(pk__in=pk_set).update(is_selected=False)

    total_cek = instance.ajuan_terkait.aggregate(total_ajuan=Sum('total_ajuan'))['total_ajuan']
    instance.total_cek = total_cek or 0
    instance.save()

class BuktiKasKeluar(models.Model):
    no_BKK = models.CharField(max_length=50, blank=True, help_text="nomor BKK akan terisi otomatis")
    tanggal_BKK = models.DateField(null=True)
    ajuan = models.ForeignKey(Ajuan, null=True, blank=True, on_delete=models.SET_NULL)
    dibayarkan_kepada = models.CharField(null=False, max_length=20)
    uraian = models.CharField(null=False, max_length=50)
    nomer_cek = models.ForeignKey(Cek, null=True, blank=True, on_delete=models.SET_NULL)
    nomer_bank_tertarik = models.ForeignKey(BankTertarik, null=True, blank=True, on_delete=models.SET_NULL)

    def __str__(self):
        return self.no_BKK

    class Meta:
        verbose_name_plural = 'Bukti Kas Keluar'

    def save(self, *args, **kwargs):
        last_id = BuktiKasKeluar.objects.order_by('-pk').first()
        if last_id:
            id_num = str(last_id.pk + 1).zfill(4)
        else:
            id_num = '00001'
        self.no_BKK = f'BK{id_num}'

        no_BKK = f'BKK{id_num}'
        count = 0
        while BuktiKasKeluar.objects.filter(no_BKK=no_BKK).exists():
            count += 1
            no_BKK = f'BKK{id_num}({count})'

        super().save(*args, **kwargs)

@receiver(pre_save, sender=BuktiKasKeluar)
def update_nomer_bank_tertarik(sender, instance, **kwargs):
    if instance.nomer_cek:
        # Jika nomer_cek telah dipilih, ambil nilai nomer_bank_tertarik dari nomer_cek
        instance.nomer_bank_tertarik = instance.nomer_cek.nomer_bank_tertarik
    else:
        # Jika nomer_cek tidak dipilih, set nomer_bank_tertarik menjadi None atau sesuai dengan kebutuhan Anda
        instance.nomer_bank_tertarik = None

# Register the signal
pre_save.connect(update_nomer_bank_tertarik, sender=BuktiKasKeluar)

class RekapBankTertarik(models.Model):
    no_RBT = models.CharField(max_length=50, blank=True, help_text="nomor RBT akan terisi otomatis")
    tanggal = models.DateField(null=True)
    no_cek = models.ForeignKey(Cek, null=True, blank=True, on_delete=models.SET_NULL)
    dana_masuk = models.ForeignKey(DanaMasuk, null=True, blank=True, on_delete=models.SET_NULL)
    dana_keluar = models.DecimalField(max_digits=20, decimal_places=0)

    class Meta:
        verbose_name_plural = 'Rekap Bank Tertarik'



