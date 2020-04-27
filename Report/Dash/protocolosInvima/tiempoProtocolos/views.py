from django.shortcuts import render

# Create your views here.

def tiempoProtocolos(request):
    return render(request, 'tiempoProtocolos/welcome.html')