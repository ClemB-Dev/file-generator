from django.db import models


class Client(models.Model):
    client_name = models.CharField(max_length=200)

    def __str__(self):
        return self.client_name


class Reason(models.Model):
    reason = models.TextField()

    def __str__(self):
        return self.reason
