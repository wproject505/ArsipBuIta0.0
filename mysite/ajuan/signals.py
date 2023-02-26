from .models import *
from django.db.models.signals import pre_save, post_save
from django.dispatch import receiver



@receiver(post_save, sender=Ajuan)
def update_rekap_pencairan_cek(sender, instance, **kwargs):
    nomor_rpc = instance.nomor_RPC

    if nomor_rpc is not None:
        # Menghitung total_ajuan pada semua objek Ajuan yang memiliki nomor_RPC yang sama
        total_ajuan = Ajuan.objects.filter(nomor_RPC=nomor_rpc).aggregate(total_ajuan=models.Sum('total_ajuan'))['total_ajuan'] or 0

        # Memperbarui nilai jumlah pada objek RekapPencairanCek
        rekap_pencairan = RekapPencairanCek.objects.get(no_RPC=nomor_rpc)
        rekap_pencairan.jumlah = total_ajuan
        rekap_pencairan.save()