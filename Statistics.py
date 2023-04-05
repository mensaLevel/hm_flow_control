from dataclasses import dataclass
import openpyxl
from datetime import datetime
import time


@dataclass
class Statistics: 
    name: str
 
    @staticmethod
    def create(cls):
        
        # Создание нового файла Excel
        cls.workbook = openpyxl.Workbook()
        cls.f = ["" for i in range (10)]
        # Выбор активного листа
        cls.worksheet = cls.workbook.active
        # Заполнение заголовков: № итерации, эталонное значение, номера рабочих мест.
        headers = ["#", "Time", "Reference","#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10"]
        cls.worksheet.append(headers)
        # Данные для записи
        cls.flow_data = []
        cls.flow_i = 0  # Счетчик проверки дубликатов строк
                
    @staticmethod
    def update(cls, iteration="", column="", reference="", flow="", status=True):
                
        if status == True:
            if flow != "": cls.f[column]= str(flow) #*round(random.uniform(0.5, 1.5), 1))
            cls.row = [iteration, datetime.now().strftime("%H:%M:%S"), reference] +cls.f
            
            if iteration > cls.flow_i:
                cls.flow_data.append(cls.row)
                cls.flow_i = iteration
                print(cls.row) # Отладочная информация
    
    @staticmethod    
    def save(cls):
        
        # Запись данных в цикле
        for row in cls.flow_data:
            cls.worksheet.append(row)
        
        # Сохранение файла
        cls.workbook.save(f"{datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.xlsx")
        cls.instance = False        

'''
########### TEST #########
stat = Statistics("Flow")

stat.create(stat)
for i in range(15):
    for u in range(10):
        time.sleep(0.01)
        
        stat.update(stat, iteration=i, column=u, reference="0.1", flow=0.15)

stat.save(stat)
'''