
from django.shortcuts import render, redirect
from django.core.paginator import Paginator
from .models import User, Task
from .forms import UserForm, TaskForm
from django.http import HttpResponse
import xlwt

def add_user(request):
    if request.method == 'POST':
        form = UserForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('user_list')
    else:
        form = UserForm()
    return render(request, 'add_user.html', {'form': form})

def add_task(request):
    if request.method == 'POST':
        form = TaskForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('task_list')
    else:
        form = TaskForm()
    return render(request, 'add_task.html', {'form': form})

def user_list(request):
    users = User.objects.all()
    paginator = Paginator(users, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    return render(request, 'user_list.html', {'page_obj': page_obj})

def task_list(request):
    tasks = Task.objects.all()
    paginator = Paginator(tasks, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    return render(request, 'task_list.html', {'page_obj': page_obj})

def export_users_and_tasks(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="users_and_tasks.xls"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws_users = wb.add_sheet('Users')
    ws_tasks = wb.add_sheet('Tasks')
    
    # Write headers for users
    row_num = 0
    columns = ['ID', 'Name', 'Email', 'Mobile']
    for col_num, column_title in enumerate(columns):
        ws_users.write(row_num, col_num, column_title)
    
    # Write headers for tasks
    row_num = 0
    columns = ['User ID', 'User Name', 'Task Detail', 'Task Type']
    for col_num, column_title in enumerate(columns):
        ws_tasks.write(row_num, col_num, column_title)
    
    # Write users data
    users = User.objects.all()
    for obj in users:
        row_num += 1
        ws_users.write(row_num, 0, obj.id)
        ws_users.write(row_num, 1, obj.name)
        ws_users.write(row_num, 2, obj.email)
        ws_users.write(row_num, 3, obj.mobile)
    
    # Write tasks data
    tasks = Task.objects.all()
    for obj in tasks:
        row_num += 1
        ws_tasks.write(row_num, 0, obj.user.id)
        ws_tasks.write(row_num, 1, obj.user.name)
        ws_tasks.write(row_num, 2, obj.task_detail)
        ws_tasks.write(row_num, 3, obj.task_type)
    
    wb.save(response)
    return response

