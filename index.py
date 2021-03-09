from tkinter import ttk, Frame, BOTH, LEFT
from tkinter import *
from openpyxl import load_workbook, Workbook
import tkinter.font as tkFont

counter = 0
wb = load_workbook('registro.xlsx')
sheet = wb['Registros']
beginrow = 2
finalrow = 1000
listB = [sheet['A' + str(i)].value for i in range(beginrow , finalrow + 1)]

class Product():

    def __init__(self, window):

        self.wind = window
        self.wind.title('Formulario de registro')
        fontStyle = tkFont.Font(family = 'Lucila Grande', size = 15)

        #Creating a frame container
        frame = LabelFrame(self.wind, text = 'Responde las preguntas', font = fontStyle)
        frame.grid(row = 0, column = 0, columnspan = 3, pady = 16)

        #Name input
        label = ttk.Label(frame, text = "Nombre: ", font = fontStyle, justify="right").grid(row = 1, column = 0)
        self.name = Entry(frame, width = 30, font = fontStyle)
        self.name.focus()
        self.name.grid(row = 1, column = 1, padx=5, pady=5, ipady=5)

        # Lastname input
        Label(frame, text = "Apellido: ", font = fontStyle).grid(row = 2, column = 0)
        self.lastname = Entry(frame, width = 30, font = fontStyle)
        self.lastname.grid(row = 2, column = 1, padx=5, pady=5, ipady=5)

        # Phone number input
        Label(frame, text = "Número de contacto: ", font = fontStyle).grid(row = 3, column = 0)
        self.num = Entry(frame, width = 30, font = fontStyle)
        self.num.grid(row = 3, column = 1, padx=5, pady=5, ipady=5)

        # Email input
        Label(frame, text = "Correo electronico: ", font = fontStyle).grid(row = 4, column = 0)
        self.email = Entry(frame, width = 30, font = fontStyle)
        self.email.grid(row = 4, column = 1, padx=5, pady=5, ipady=5)

        #Economic activity input
        Label(frame, text = 'Tipo de actividad economica: ', font = fontStyle).grid(row = 5, column = 0)
        self.comboType = ttk.Combobox(frame, state='readonly', width = 30, font = fontStyle)
        self.comboType.grid(row = 5, column = 1, padx=5, pady=5, ipady=5)
        options = ["Palma", "Caucho"]
        self.comboType["values"] = options

        # production zone input
        Label(frame, text = 'Zona donde realiza producción : ', font = fontStyle).grid(row = 6, column = 0)
        self.comboZone = ttk.Combobox(frame, state='readonly', width = 30, font = fontStyle)
        self.comboZone.grid(row = 6, column = 1, padx=5, pady=5, ipady=5)
        options = ["Corregimiento El Centro", "Corregimiento La Fortuna", "Corregimiento Ciénaga del Opón", "Corregimiento Meseta de San Rafael", "Corregimiento El Llanito", "Corregimiento San Rafael de Chucuri" ]
        self.comboZone["values"] = options

        # Qty input
        Label(frame, text = "Capacidad de producción en toneladas: ", font = fontStyle).grid(row = 7, column = 0)
        self.qty = Entry(frame, width = 30, font = fontStyle)
        self.qty.grid(row = 7, column = 1, padx=5, pady=5, ipady=5)

        # bussines input
        Label(frame, text = "Entidades o empresas con las que ha comercializado: ", font = fontStyle).grid(row = 8, column = 0)
        self.buss1 = Entry(frame, width = 30, font = fontStyle)
        self.buss1.grid(row = 8, column = 1, padx=5, pady=5, ipady=5)

        # Workers input
        Label(frame, text = "Numero de trabajadores en la producción: ", font = fontStyle).grid(row = 9, column = 0)
        self.work = Entry(frame, width = 30, font = fontStyle)
        self.work.grid(row = 9, column = 1, padx=5, pady=5, ipady=5)

        # time input
        Label(frame, text = "Numero de tiempo que lleva realizando la actividad productiva anual: ", font = fontStyle).grid(row = 10, column = 0)
        self.time = Entry(frame, width = 30, font = fontStyle)
        self.time.grid(row = 10, column = 1, padx=5, pady=5, ipady=5)

        # Area input
        Label(frame, text = "Extensión de tierra donde produce en hectáreas: ", font = fontStyle).grid(row = 11, column = 0)
        self.area = Entry(frame, width = 30, font = fontStyle)
        self.area.grid(row = 11, column = 1, padx=5, pady=5, ipady=5)

        # Logistic input
        Label(frame, text = "Logística para el transporte del producto : ", font = fontStyle).grid(row = 12, column = 0)
        self.comboLogistic = ttk.Combobox(frame, state='readonly', width = 30, font = fontStyle)
        self.comboLogistic.grid(row = 12, column = 1, padx=5, pady=5, ipady=5)
        options = ["No cuenta con transporte", "Transporte local", "Transporte departamental", "Transporte nacional"]
        self.comboLogistic["values"] = options

        # register input
        Label(frame, text = "Registro mercantil: ", font = fontStyle).grid(row = 13, column = 0)
        self.comboRegister = ttk.Combobox(frame, state='readonly', width = 30, font = fontStyle)
        self.comboRegister.grid(row = 13, column = 1, padx=5, pady=5, ipady=5)
        options = ["Si", "No"]
        self.comboRegister["values"] = options

        # Button add
        button = Button(frame, text = "Enviar", font = fontStyle, bg = 'blue', fg = 'white', command = self.obtain_data).grid(row = 15, columnspan = 2, sticky = W + E, padx=5, pady=5, ipady=5)

        #Output messages
        self.message = Label(text = '', fg = 'red')
        self.message.grid(row = 15, columnspan = 2, sticky = W + E)

    def validation(self):
        return len(self.name.get()) != 0 and len(self.lastname.get()) != 0 and len(self.num.get()) != 0 and len(self.email.get()) != 0 and len(self.comboType.get()) != 0 and len(self.comboZone.get()) != 0 and len(self.qty.get()) != 0 and len(self.buss1.get()) != 0 and len(self.work.get()) != 0 and len(self.time.get()) != 0 and len(self.area.get()) != 0 and len(self.comboLogistic.get()) != 0 and len(self.comboRegister.get()) != 0

    def obtain_data(self):
        global counter, comboType, comboZone, comboLogistic, comboRegister, frame, name, lastname, num, email, qty, buss1, work, time, area, sheet
        if self.validation():
            counter += 1
            print(counter)
            number = 1

            #Listas de cada fila / list of rows
            row_name = sheet['B']
            row_lastname = sheet['C']
            row_num = sheet['D']
            row_email = sheet['E']
            row_type = sheet['F']
            row_zone = sheet['G']
            row_qty = sheet['H']
            row_buss1 = sheet['I']
            row_work = sheet['J']
            row_time = sheet['K']
            row_area = sheet['L']
            row_logistic = sheet['M']
            row_register = sheet['N']
            #Nombre / Name
            for i in row_name:
                if i.value is not None:
                    number += 1
                elif i.value is None:
                    rowname = "B" + str(number)
                    rowlastname = "C" + str(number)
                    rownum = "D" + str(number)
                    rowemail = "E" + str(number)
                    rowtype = "F" + str(number)
                    rowzone = "G" + str(number)
                    rowqty = "H" + str(number)
                    rowbuss1 = "I" + str(number)
                    rowwork = "J" + str(number)
                    rowtime = "K" + str(number)
                    rowarea = "L" + str(number)
                    rowlogistic = "M" + str(number)
                    rowregister = "N" + str(number)

                    sheet[rowname] = self.name.get()
                    sheet[rowlastname] = self.lastname.get()
                    sheet[rownum] = self.num.get()
                    sheet[rowemail] = self.email.get()
                    sheet[rowtype] = self.comboType.get()
                    sheet[rowzone] = self.comboZone.get()
                    sheet[rowqty] = self.qty.get()
                    sheet[rowbuss1] = self.buss1.get()
                    sheet[rowwork] = self.work.get()
                    sheet[rowtime] = self.time.get()
                    sheet[rowarea] = self.area.get()
                    sheet[rowlogistic] = self.comboLogistic.get()
                    sheet[rowregister] = self.comboRegister.get()
                    break

            wb.save('registro.xlsx')



            #Borrar inputs y combobox / clean inputs and combobox
            self.message['text'] = 'Se ha registrado satisfactoriamente'
            self.comboType.set("")
            self.comboZone.set("")
            self.comboLogistic.set("")
            self.comboRegister.set("")
            self.name.delete(0, END)
            self.lastname.delete(0, END)
            self.num.delete(0, END)
            self.email.delete(0, END)
            self.qty.delete(0, END)
            self.buss1.delete(0, END)
            self.work.delete(0, END)
            self.time.delete(0, END)
            self.area.delete(0, END)
        else:
            print("Debe rellenar todos los datos")

if __name__ == '__main__':
    window = Tk()
    application = Product(window)
    window.mainloop()
