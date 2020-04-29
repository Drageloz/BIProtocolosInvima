from django.urls import path
from . import views
from invima.dash_apps.finished_apps import simpleexample
from invima.dash_apps.finished_apps import tiempoprotocolos


urlpatterns = [
    path('', views.tiempoProtocolos, name='tiempoProtocolos')
]