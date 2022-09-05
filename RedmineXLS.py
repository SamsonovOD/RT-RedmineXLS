from redminelib import Redmine
import datetime
from datetime import timedelta
import xlsxwriter 

redmine = Redmine('http://lab.rt.ru/', username=open("username.txt", "r").read(), password=open("password.txt", "r").read()) # Авторизация RedMine API

def rm_get_id(s): # Полчуить ID проекта по имени
    for i in redmine.project.all():
        if(s) == i.name:
            return i.id

def rm_get_projects(): # Вывести список проектов (Предположительно до 61: Sagemcom F@st 5655V2 включительно)
    return redmine.project.all(limit=47)
    
def rm_get_projtickets(p_id): #Вывести список тикетов в проекте
    return redmine.issue.filter(project_id=p_id, status_id='*')
    
def rm_get_subprojtickets(p_id, sp_id): #Вывести список тикетов в подроекте проекта
    return redmine.issue.filter(project_id=p_id, subproject_id=sp_id, status_id='*')
    
if __name__ == "__main__":
    worktable = xlsxwriter.Workbook('redmine_'+str(datetime.datetime.now()).split(" ")[0]+'.xls')
    worksheet = worktable.add_worksheet()
    row = 0
    worksheet.write(row, 0, "Проект") 
    worksheet.write(row, 1, "Вендор") 
    worksheet.write(row, 2, "Модель") 
    worksheet.write(row, 3, "HW") 
    worksheet.write(row, 4, "Задачи") 
    worksheet.write(row, 5, "Открыто")
    worksheet.write(row, 6, "Недавно обновлено") 
    worksheet.write(row, 7, "Недавно закрыто")  
    row = 1
    current = "P"
    for project in rm_get_projects():
        if hasattr(project, 'parent'):
            if project.name.find("HW:") != -1:
                model = project.name[:project.name.find("HW:")-2]
                hw = project.name[project.name.find("HW:")+4:-1]
            else:
                if project.name.find("HW") != -1:
                    model = project.name[:project.name.find("HW")-1]
                    hw = project.name[project.name.find("HW")+3:]
                else:
                    model = project.name
                    hw = "-"
            tc = 0
            to = 0
            tru = 0
            trs = 0
            for ticket in rm_get_projtickets(project.id):
                tc += 1
                if str(ticket.status) != "Closed":
                    to += 1
                if ticket.updated_on > datetime.datetime.now() - timedelta(days=7):
                    tru += 1
                    if str(ticket.status) == "Closed":
                        trs += 1
            worksheet.write(row, 0, current + " " + project.name + " (" + project.identifier + ")") 
            worksheet.write(row, 1, current) 
            worksheet.write(row, 2, model) 
            worksheet.write(row, 3, hw) 
            worksheet.write(row, 4, str(tc)) 
            worksheet.write(row, 5, str(to)) 
            worksheet.write(row, 6, str(tru))
            worksheet.write(row, 7, str(trs))
            row += 1
        else:
            current = project.name
    worktable.close()
    print("Done.")