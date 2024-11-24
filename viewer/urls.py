from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),  # Main page
    path('download_csv/', views.download_csv, name='download_csv'),  # CSV download
]
