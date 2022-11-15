from django.urls import path
from .views import index, file_generator


urlpatterns = [
    path('', index, name='index'),
    path('file_generator/', file_generator, name='file_generator'),
]
