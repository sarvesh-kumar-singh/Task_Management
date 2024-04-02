"""
URL configuration for task_management_project project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from task_management import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('add_user/', views.add_user, name='add_user'),
    path('add_task/', views.add_task, name='add_task'),
    path('user_list/', views.user_list, name='user_list'),
    path('task_list/', views.task_list, name='task_list'),
    path('export_users_and_tasks/', views.export_users_and_tasks, name='export_users_and_tasks'),
]
