from django.contrib import admin
from .models import Client, Reason


# class ClientAdmin(admin.ModelAdmin):
#     list_display = ('client_name')


# class ReasonAdmin(admin.ModelAdmin):
#     list_display = ('reason')


admin.site.register(Client)
admin.site.register(Reason)
