from django.contrib import admin
from django.urls import path
from geoCodeXML import views as geoCode_views


urlpatterns = [
    path('admin/', admin.site.urls),
    path('',geoCode_views.index, name='index')
]