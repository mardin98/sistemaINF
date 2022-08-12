import datetime
import time
import win32api
import win32con
import wx
import wx.grid
import sqlite3
from time import localtime, strftime
import os
from skimage import io as iio
import io
import zlib
import dlib
import numpy as np
import cv2
import _thread
import threading
import win32com.client
import tkinter as tk
from tkinter import filedialog
import csv
import pymysql

spk = win32com.client.Dispatch("SAPI.SpVoice")
#Prueba
ID_NEW_REGISTER = 160
ID_FINISH_REGISTER = 161

ID_START_PUNCHCARD = 190
ID_END_PUNCARD = 191

ID_FIN_PUNCHCARD = 194
ID_FIN2_PUNCARD = 195

ID_TODAY_LOGCAT = 220
ID_CUSTOM_LOGCAT = 260

ID_WORKING_HOURS = 301
ID_OFFWORK_HOURS = 302
ID_DELETE = 303

ID_OPEN_LOGCAT = 283 #Cuenteo Asistencia DB
ID_CLOSE_LOGCAT = 284

ID_WORKER_UNAVIABLE = -1

PATH_FACE = "data/face_img_database/"
# face recognition model, the object maps human faces into 128D vectors
facerec = dlib.face_recognition_model_v1("model/dlib_face_recognition_resnet_model_v1.dat")

detector = dlib.get_frontal_face_detector()
predictor = dlib.shape_predictor('model/shape_predictor_68_face_landmarks.dat')


def speak_info(info):
    spk.Speak(info)


def return_euclidean_distance(feature_1, feature_2):
    feature_1 = np.array(feature_1)
    feature_2 = np.array(feature_2)
    dist = np.sqrt(np.sum(np.square(feature_1 - feature_2)))
    print("distancia euclidiana: ", dist)
    if dist > 0.4:
        return "diferente"
    else:
        return "similar"


class WAS(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, title="Sistema de tarjeta perforada de monitoreo inteligente", size=(1032, 560))

        self.Folderpath = None
        self.initMenu()
        self.initInfoText()
        self.initGallery()
        self.initDatabase()
        self.initData()

    def initData(self):
        self.name = ""
        self.id = ID_WORKER_UNAVIABLE
        self.face_feature = ""
        self.pic_num = 0
        self.flag_registed = False
        self.loadDataBase(1)

    def initMenu(self):

        menuBar = wx.MenuBar() # generar barra de menú
        menu_Font = wx.Font()  # Font(faceName="consolas",pointsize=20)
        menu_Font.SetPointSize(14)
        menu_Font.SetWeight(wx.BOLD)

        registerMenu = wx.Menu() # generar submenú
        self.new_register = wx.MenuItem(registerMenu, ID_NEW_REGISTER, "nueva entrada")
        self.new_register.SetBitmap(wx.Bitmap("drawable/new_register.png"))
        self.new_register.SetTextColour("SLATE BLACK")
        self.new_register.SetFont(menu_Font)
        registerMenu.Append(self.new_register)

        self.finish_register = wx.MenuItem(registerMenu, ID_FINISH_REGISTER, "completar la entrada")
        self.finish_register.SetBitmap(wx.Bitmap("drawable/finish_register.png"))
        self.finish_register.SetTextColour("SLATE BLACK")
        self.finish_register.SetFont(menu_Font)
        self.finish_register.Enable(False)
        registerMenu.Append(self.finish_register)

        puncardMenu = wx.Menu()
        self.start_punchcard = wx.MenuItem(puncardMenu, ID_START_PUNCHCARD, "empezar a iniciar sesión")
        self.start_punchcard.SetBitmap(wx.Bitmap("drawable/start_punchcard.png"))
        self.start_punchcard.SetTextColour("SLATE BLACK")
        self.start_punchcard.SetFont(menu_Font)
        puncardMenu.Append(self.start_punchcard)

        self.end_puncard = wx.MenuItem(puncardMenu, ID_END_PUNCARD, "finalizar sesión")
        self.end_puncard.SetBitmap(wx.Bitmap("drawable/end_puncard.png"))
        self.end_puncard.SetTextColour("SLATE BLACK")
        self.end_puncard.SetFont(menu_Font)
        self.end_puncard.Enable(False)
        puncardMenu.Append(self.end_puncard)

        #Prueba registro salida
        salidaR = wx.Menu()
        self.salida_punchcard = wx.MenuItem(salidaR, ID_FIN_PUNCHCARD, "Salir de  sesión")
        self.salida_punchcard.SetBitmap(wx.Bitmap("drawable/start_punchcard.png"))
        self.salida_punchcard.SetTextColour("SLATE BLACK")
        self.salida_punchcard.SetFont(menu_Font)
        salidaR.Append(self.salida_punchcard)

        self.salida2_puncard = wx.MenuItem(salidaR, ID_FIN2_PUNCARD, "finalizar sesión")
        self.salida2_puncard.SetBitmap(wx.Bitmap("drawable/end_puncard.png"))
        self.salida2_puncard.SetTextColour("SLATE BLACK")
        self.salida2_puncard.SetFont(menu_Font)
        self.salida2_puncard.Enable(False)
        salidaR.Append(self.salida2_puncard)
        #Fin prueba

        logcatMenu = wx.Menu()
        self.today_logcat = wx.MenuItem(logcatMenu, ID_TODAY_LOGCAT, "Salida del registro de hoy")
        self.today_logcat.SetBitmap(wx.Bitmap("drawable/open_logcat.png"))
        self.today_logcat.SetFont(menu_Font)
        self.today_logcat.SetTextColour("SLATE BLACK")
        logcatMenu.Append(self.today_logcat)

        self.custom_logcat = wx.MenuItem(logcatMenu, ID_CUSTOM_LOGCAT, "registro personalizado de salida")
        self.custom_logcat.SetBitmap(wx.Bitmap("drawable/open_logcat.png"))
        self.custom_logcat.SetFont(menu_Font)
        self.custom_logcat.SetTextColour("SLATE BLACK")
        logcatMenu.Append(self.custom_logcat)

        setMenu = wx.Menu()
        self.working_hours = wx.MenuItem(setMenu, ID_WORKING_HOURS, "Horas Laborales")
        self.working_hours.SetBitmap(wx.Bitmap("drawable/close_logcat.png"))
        self.working_hours.SetFont(menu_Font)
        self.working_hours.SetTextColour("SLATE BLACK")
        setMenu.Append(self.working_hours)

        self.offwork_hours = wx.MenuItem(setMenu, ID_OFFWORK_HOURS, "tiempo libre")
        self.offwork_hours.SetBitmap(wx.Bitmap("drawable/open_logcat.png"))
        self.offwork_hours.SetFont(menu_Font)
        self.offwork_hours.SetTextColour("SLATE BLACK")
        setMenu.Append(self.offwork_hours)

        self.delete = wx.MenuItem(setMenu, ID_DELETE, "eliminar personas")
        self.delete.SetBitmap(wx.Bitmap("drawable/end_puncard.png"))
        self.delete.SetFont(menu_Font)
        self.delete.SetTextColour("SLATE BLACK")
        setMenu.Append(self.delete)


        #Prueba Asitencia DB

        RegistroA = wx.Menu()
        self.open_logcat = wx.MenuItem(RegistroA, ID_OPEN_LOGCAT, "Registro Abierto")
        self.open_logcat.SetBitmap(wx.Bitmap("drawable/open_logcat.png"))
        self.open_logcat.SetFont(menu_Font)
        self.open_logcat.SetTextColour("SLATE BLUE")
        RegistroA.Append(self.open_logcat)

        self.close_logcat = wx.MenuItem(RegistroA, ID_CLOSE_LOGCAT, "Cerrar registro")
        self.close_logcat.SetBitmap(wx.Bitmap("drawable/close_logcat.png"))
        self.close_logcat.SetFont(menu_Font)
        self.close_logcat.SetTextColour("SLATE BLUE")
        RegistroA.Append(self.close_logcat)

        #finprueba Asitencia DB

        menuBar.Append(registerMenu, "Registro")
        menuBar.Append(puncardMenu, "&Entrada")
        menuBar.Append(salidaR, "&Salida")
        menuBar.Append(logcatMenu, "&Exportar A")
        menuBar.Append(RegistroA, "Asistencia") # Menu Asistencia DB
        menuBar.Append(setMenu, "&configurar")

        self.SetMenuBar(menuBar)

        self.Bind(wx.EVT_MENU, self.OnNewRegisterClicked, id=ID_NEW_REGISTER)
        self.Bind(wx.EVT_MENU, self.OnFinishRegisterClicked, id=ID_FINISH_REGISTER)
        self.Bind(wx.EVT_MENU, self.OnStartPunchCardClicked, id=ID_START_PUNCHCARD)
        self.Bind(wx.EVT_MENU, self.OnEndPunchCardClicked, id=ID_END_PUNCARD)
        #Prueba
        self.Bind(wx.EVT_MENU, self.OnSalidaPunchCardClicked, id=ID_FIN_PUNCHCARD)
        self.Bind(wx.EVT_MENU, self.OnSalida2PunchCardClicked, id=ID_FIN2_PUNCARD)
        #Fin
        self.Bind(wx.EVT_MENU, self.ExportTodayLog, id=ID_TODAY_LOGCAT)
        self.Bind(wx.EVT_MENU, self.ExportCustomLog, id=ID_CUSTOM_LOGCAT)
        self.Bind(wx.EVT_MENU, self.SetWorkingHours, id=ID_WORKING_HOURS)
        self.Bind(wx.EVT_MENU, self.SetOffWorkHours, id=ID_OFFWORK_HOURS)
        self.Bind(wx.EVT_MENU, self.deleteBtn, id=ID_DELETE)
        self.Bind(wx.EVT_MENU, self.OnOpenLogcatClicked, id=ID_OPEN_LOGCAT) #Ventana de registro abierto
        self.Bind(wx.EVT_MENU, self.OnCloseLogcatClicked, id=ID_CLOSE_LOGCAT)

        #PRueba 2
    def OnOpenLogcatClicked(self, event):
            self.loadDataBase(2)
            # Debe ampliarse para mostrar el desplazamiento
            self.SetSize(1240, 560)
            grid = wx.grid.Grid(self, pos=(420, 0), size=(800, 500))
            grid.CreateGrid(100, 5)
            for i in range(100):
                for j in range(5):
                    grid.SetCellAlignment(i, j, wx.ALIGN_CENTER, wx.ALIGN_CENTER)
            grid.SetColLabelValue(0, "ID")  # etiqueta de la primera columna
            grid.SetColLabelValue(1, "Nombre")
            grid.SetColLabelValue(2, "Tiempo de marcado")
            grid.SetColLabelValue(3, "Tiempo de salida")
            grid.SetColLabelValue(4, "Estas tarde")

            grid.SetColSize(0, 120)
            grid.SetColSize(1, 120)
            grid.SetColSize(2, 150)
            grid.SetColSize(3, 150)
            grid.SetColSize(4, 150)

            # grid.SetCellTextColour("NAVY") Esta línea informa de un error en algunas máquinas
            for i, id in enumerate(self.logcat_id):
                grid.SetCellValue(i, 0, str(id))
                grid.SetCellValue(i, 1, self.logcat_name[i])
                grid.SetCellValue(i, 2, self.logcat_datetime[i])
                grid.SetCellValue(i, 3, self.logcat_datetimeSa[i])
                grid.SetCellValue(i, 4, self.logcat_late[i])




    pass

    def OnCloseLogcatClicked(self, event):
            self.initGallery()
            self.SetSize(1032, 560)





    pass

        #Prueba 2

    def SetWorkingHours(self, event):
        global working
        global setWorkingSign
        setWorkingSign = False
        self.loadDataBase(1)
        # self.working_hours.Enable(True)
        self.working_hours = wx.GetTextFromUser(message="Por favor, introduzca el horario de trabajo", caption="Consejos", default_value="07:00:00",
                                                parent=None)
        working = self.working_hours
        setWorkingSign = True
        pass

    def SetOffWorkHours(self, event):
        global offworking
        self.loadDataBase(1)
        # self.offwork_hours.Enable(True)
        self.offwork_hours = wx.GetTextFromUser(message="Por favor, introduzca la hora de cierre", caption="Consejos", default_value="16:00:00",
                                                parent=None)
        offworking = self.offwork_hours
        win32api.MessageBox(0, "Asegúrese de configurar tanto el tiempo de encendido como el de apagado y configure el tiempo de encendido primero", "recordar", win32con.MB_ICONWARNING)
        if setWorkingSign:
            self.loadDataBase(4)
        else:
            win32api.MessageBox(0, "No ha establecido horas de trabajo", "recordar", win32con.MB_ICONWARNING)
        pass

    def ExportTodayLog(self, event):
        global Folderpath1
        Folderpath1 = ""
        self.save_route1(event)
        if not Folderpath1 == "":
            self.loadDataBase(3)
            day = time.strftime("%Y-%m-%d")
            path = Folderpath1 + "/" + day + ".csv"
            f = open(path, 'w', newline='', encoding='utf-8')
            csv_writer = csv.writer(f)
            csv_writer.writerow(["Numero de serie", "Nombre", "tiempo de perforacion", "Estas tarde"])
            size = len(logcat_id)
            index = 0
            while size - 1 >= index:
                localtime1 = str(logcat_datetime[index]).replace('[', '').replace(']', '')
                csv_writer.writerow([logcat_id[index], logcat_name[index], localtime1, logcat_late[index]])
                index += 1;
            f.close()
        pass

    def ExportCustomLog(self, event):
        global dialog
        global t1
        global t2
        global Folderpath2
        Folderpath2 = ""
        dialog = wx.Dialog(self)
        Label1 = wx.StaticText(dialog, -1, "ID del empleado", pos=(30, 10))
        t1 = wx.TextCtrl(dialog, -1, '', pos=(150, 10), size=(130, -1))
        Label2 = wx.StaticText(dialog, -1, "fecha de salida (días)", pos=(30, 50))
        sampleList = [u'1', u'3', u'7', u'30']
        t2 = wx.ComboBox(dialog, -1, value="1", pos=(150, 50), size=(130, -1), choices=sampleList,
                         style=wx.CB_READONLY)
        button = wx.Button(dialog, -1, "Seleccione la ruta para guardar el archivo", pos=(60, 90))
        button.Bind(wx.EVT_BUTTON, self.save_route2, button)
        btn_confirm = wx.Button(dialog, 1, "confirmar", pos=(30, 150))
        btn_close = wx.Button(dialog, 2, "Cancelar", pos=(250, 150))
        btn_close.Bind(wx.EVT_BUTTON, self.OnClose, btn_close)
        btn_confirm.Bind(wx.EVT_BUTTON, self.DoCustomLog, btn_confirm)
        dialog.ShowModal()
        pass

    # Antes de cerrar la ventana principal, asegúrese de que esté realmente cerrada
    def OnClose(self, event):
        dlg = wx.MessageDialog(None, u'¿Está seguro de que desea cerrar esta ventana?', u'Consejos de operación', wx.YES_NO)
        if dlg.ShowModal() == wx.ID_YES:
            dialog.Destroy()

    def OnClose1(self, event):
        dlg = wx.MessageDialog(None, u'¿Está seguro de que desea cerrar esta ventana?', u'Consejos de operación', wx.YES_NO)
        if dlg.ShowModal() == wx.ID_YES:
            dialog1.Destroy()

    def OnYes(self, event):
        dlg = wx.MessageDialog(None, u'¿Está seguro de que desea eliminar al empleado con este número?', u'Consejos de operación', wx.YES_NO)
        if dlg.ShowModal() == wx.ID_YES:
            return True

    def deleteBtn(self, event):
        global dialog1
        global t4
        dialog1 = wx.Dialog(self)
        Label1 = wx.StaticText(dialog1, -1, "Ingrese la identificación del empleado: ", pos=(40, 34))
        t4 = wx.TextCtrl(dialog1, -1, '', pos=(130, 30), size=(130, -1))
        btn_confirm = wx.Button(dialog1, 1, "confirmar", pos=(30, 150))
        btn_close = wx.Button(dialog1, 2, "Cancelar", pos=(250, 150))
        btn_close.Bind(wx.EVT_BUTTON, self.OnClose1, btn_close)
        btn_confirm.Bind(wx.EVT_BUTTON, self.deleteById, btn_confirm)
        dialog1.ShowModal()

    def DoCustomLog(self, event):
        if not Folderpath2 == "":
            number = t1.GetValue()
            days = t2.GetValue()
            flag = self.findById(number, days)
            print("El número de días para consultar es：", days)
            if flag:
                row = len(find_id)
                path = Folderpath2 + '/' + find_name[0] + '.csv'
                f = open(path, 'w', newline='', encoding='utf-8')
                csv_writer = csv.writer(f)
                csv_writer.writerow(["Numero de serie", "Nombre", "tiempo de perforacion", "Estas tarde"])
                for index in range(row):
                    s1 = str(find_datetime[index]).replace('[', '').replace(']', '')
                    csv_writer.writerow([str(find_id[index]), str(find_name[index]), s1, str(find_late[index])])

                f.close()
                success = wx.MessageDialog(None, 'El registro se guardó correctamente, preste atención para verificar', 'info', wx.OK)
                success.ShowModal()
            else:
                warn = wx.MessageDialog(None, 'La identificación de entrada es incorrecta, vuelva a ingresar', 'info', wx.OK)
                warn.ShowModal()
            dialog.Destroy()
        else:
            win32api.MessageBox(0, "Ingrese la ubicación de exportación del archivo", "recordar", win32con.MB_ICONWARNING)

        pass

    def deleteById(self, event):
        global delete_name
        delete_name = []
        id = t4.GetValue()
        print("Eliminar empleado con id como:", id)
        conn = sqlite3.connect("inspurer.db")  # Establecer conexión con la base de datos
        cur = conn.cursor()  # obtener el objeto del cursor
        sql = 'select name from worker_info where id=' + id
        sql1 = 'delete from worker_info where id=' + id
        sql2 = 'delete from logcat where id=' + id
        length = len(cur.execute(sql).fetchall())
        if length <= 0:
            win32api.MessageBox(0, "No se encuentra el empleado, vuelva a ingresar la identificación", "recordar", win32con.MB_ICONWARNING)
            return False
        else:
            origin = cur.execute(sql).fetchall()
            for row in origin:
                delete_name.append(row[0])
                name = delete_name[0]
                print("nombre es", name)
            if self.OnYes(event):
                cur.execute(sql1)
                cur.execute(sql2)
                conn.commit()
                dir = PATH_FACE + name
                for file in os.listdir(dir):
                    os.remove(dir + "/" + file)
                    print("La imagen con la cara grabada ha sido eliminada", dir + "/" + file)
                os.rmdir(PATH_FACE + name)
                print("Carpeta de nombre eliminada con caras grabadas", dir)
                dialog1.Destroy()
                self.initData()
                return True

    def findById(self, id, day):
        global find_id, find_name, find_datetime, find_late
        find_id = []
        find_name = []
        find_datetime = []
        find_late = []
        DayAgo = (datetime.datetime.now() - datetime.timedelta(days=int(day)))
        # Convertir a otros formatos de cadena:
        day_before = DayAgo.strftime("%Y-%m-%d")
        today = datetime.date.today()
        first = today.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        print(last_month.strftime("%Y-%m"))
        print(last_month)
        conn = sqlite3.connect("inspurer.db")  # Establecer conexión con la base de datos
        cur = conn.cursor()  # obtener el objeto del cursor
        sql = 'select id ,name,datetime,late from logcat where id=' + id

        if day == '30':
            str = "'"
            sql1 = 'select id ,name,datetime,late from logcat where id=' + id + ' ' + 'and datetime like ' + str + '%' + last_month.strftime(
                "%Y-%m") + '%' + str
        else:
            sql1 = 'select id ,name,datetime,late from logcat where id=' + id + ' ' + 'and datetime>=' + day_before
        length = len(cur.execute(sql).fetchall())
        if length <= 0:
            return False
        else:
            cur.execute(sql1)
            origin = cur.fetchall()
            for row in origin:
                find_id.append(row[0])
                find_name.append(row[1])
                find_datetime.append(row[2])
                find_late.append(row[3])
            return True
        pass

    def save_route1(self, event):
        global Folderpath1
        root = tk.Tk()
        root.withdraw()
        Folderpath1 = filedialog.askdirectory()  # Obtener la carpeta seleccionada
        pass

    def save_route2(self, event):
        global Folderpath2
        root = tk.Tk()
        root.withdraw()
        Folderpath2 = filedialog.askdirectory()  # Obtener la carpeta seleccionada
        pass

    def register_cap(self, event):
        # crear objeto de cámara cv2
        self.cap = cv2.VideoCapture(0)

        while self.cap.isOpened():

            flag, im_rd = self.cap.read()

            # Cada cuadro de datos tiene un retraso de 1 ms, y el retraso es 0 para leer un cuadro estático
            kk = cv2.waitKey(1)
            # detalles de conteo de rostros
            dets = detector(im_rd, 1)

            # detectar rostro
            if len(dets) != 0:
                biggest_face = dets[0]
                # Toma la cara con la mayor proporción
                maxArea = 0
                for det in dets:
                    w = det.right() - det.left()
                    h = det.top() - det.bottom()
                    if w * h > maxArea:
                        biggest_face = det
                        maxArea = w * h
                        # dibujar rectángulo

                cv2.rectangle(im_rd, tuple([biggest_face.left(), biggest_face.top()]),
                              tuple([biggest_face.right(), biggest_face.bottom()]),
                              (255, 0, 0), 2)
                img_height, img_width = im_rd.shape[:2]
                image1 = cv2.cvtColor(im_rd, cv2.COLOR_BGR2RGB)
                pic = wx.Bitmap.FromBuffer(img_width, img_height, image1)
                # Mostrar la imagen en el panel
                self.bmp.SetBitmap(pic)

                # Obtenga las características de todas las caras de la imagen capturada actualmente y guárdelas en features_cap_arr
                shape = predictor(im_rd, biggest_face)
                features_cap = facerec.compute_face_descriptor(im_rd, shape)

                # Para una cara, recorre todas las características faciales almacenadas
                for i, knew_face_feature in enumerate(self.knew_face_feature):
                    # Compare una cara con todos los datos faciales almacenados
                    compare = return_euclidean_distance(features_cap, knew_face_feature)
                    if compare == "similar":  # encontrado caras similares
                        self.infoText.AppendText(self.getDateAndTime() + "Número de empleo:" + str(self.knew_id[i])
                                                 + " Nombre:" + self.knew_name[i] + " de los datos faciales ya existe\r\n")
                        self.flag_registed = True
                        self.OnFinishRegister()
                        _thread.exit()

                face_height = biggest_face.bottom() - biggest_face.top()
                face_width = biggest_face.right() - biggest_face.left()
                im_blank = np.zeros((face_height, face_width, 3), np.uint8)
                try:
                    for ii in range(face_height):
                        for jj in range(face_width):
                            im_blank[ii][jj] = im_rd[biggest_face.top() + ii][biggest_face.left() + jj]
                    if len(self.name) > 0:
                        cv2.imencode('.jpg', im_blank)[1].tofile(
                            PATH_FACE + self.name + "/img_face_" + str(self.pic_num) + ".jpg") # Manera correcta
                        self.pic_num += 1
                        print("escribir a locales：", str(PATH_FACE + self.name) + "/img_face_" + str(self.pic_num) + ".jpg")
                        self.infoText.AppendText(
                            self.getDateAndTime() + "imagen:" + str(PATH_FACE + self.name) + "/img_face_" + str(
                                self.pic_num) + ".jpg Guardado exitosamente\r\n")
                except:
                    print("La foto guardada es anormal, apunte a la cámara")

                if self.new_register.IsEnabled():
                    _thread.exit()
                if self.pic_num == 30:
                    self.OnFinishRegister()
                    _thread.exit()

    def OnNewRegisterClicked(self, event):
        self.new_register.Enable(False)
        self.finish_register.Enable(True)
        self.loadDataBase(1)
        while self.id == ID_WORKER_UNAVIABLE:
            self.id = wx.GetNumberFromUser(message="Ingrese su número de trabajo (-1 no está disponible)",
                                           prompt="Número de empleo", caption="Consejos",
                                           value=ID_WORKER_UNAVIABLE,
                                           parent=self.bmp, max=100000000, min=ID_WORKER_UNAVIABLE)
            for knew_id in self.knew_id:
                if knew_id == self.id:
                    self.id = ID_WORKER_UNAVIABLE
                    wx.MessageBox(message="El número de trabajo ya existe, vuelva a ingresar", caption="advertir")

        while self.name == '':
            self.name = wx.GetTextFromUser(message="Por favor ingrese su nombre para crear una carpeta de nombre",
                                           caption="Consejos",
                                           default_value="", parent=self.bmp)

            # Comprobar si hay nombres duplicados
            for exsit_name in (os.listdir(PATH_FACE)):
                if self.name == exsit_name:
                    wx.MessageBox(message="La carpeta de nombre ya existe, vuelva a ingresar", caption="advertir")
                    self.name = ''
                    break
        os.makedirs(PATH_FACE + self.name)
        _thread.start_new_thread(self.register_cap, (event,))
        pass

    def OnFinishRegister(self):

        self.new_register.Enable(True)
        self.finish_register.Enable(False)
        self.cap.release()

        self.bmp.SetBitmap(wx.Bitmap(self.pic_index))
        if self.flag_registed == True:
            dir = PATH_FACE + self.name
            for file in os.listdir(dir):
                os.remove(dir + "/" + file)
                print("La imagen con la cara grabada ha sido eliminada", dir + "/" + file)
            os.rmdir(PATH_FACE + self.name)
            print("Carpeta de nombre eliminada con caras grabadas", dir)
            self.initData()
            return
        if self.pic_num > 0:
            pics = os.listdir(PATH_FACE + self.name)
            feature_list = []
            feature_average = []
            for i in range(len(pics)):
                pic_path = PATH_FACE + self.name + "/" + pics[i]
                print("La imagen de la cara que se lee:", pic_path)
                img = iio.imread(pic_path)
                img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                dets = detector(img_gray, 1)
                if len(dets) != 0:
                    shape = predictor(img_gray, dets[0])
                    face_descriptor = facerec.compute_face_descriptor(img_gray, shape)
                    feature_list.append(face_descriptor)
                else:
                    face_descriptor = 0
                    print("No se reconoce rostro en la foto")
            if len(feature_list) > 0:
                for j in range(128):
                    # prevenir fuera de los límites
                    feature_average.append(0)
                    for i in range(len(feature_list)):
                        feature_average[j] += feature_list[i][j]
                    feature_average[j] = (feature_average[j]) / len(feature_list)
                self.insertARow([self.id, self.name, feature_average], 1)
                con = pymysql.connect(db='baseinf', user='root', passwd='', host='localhost', port=3306,
                                      autocommit=True)
                cur = con.cursor()
                sql = "INSERT INTO `worker_info1` (`id`, `name`, `face_feature`) VALUES (%s, %s, %s)"
                cur.execute(sql, (self.id, self.name, self.adapt_array(feature_average)))
                self.infoText.AppendText(self.getDateAndTime() + "Número de empleo:" + str(self.id)
                                         + " Nombre:" + self.name + " de los datos faciales se ha almacenado correctamente\r\n")
            pass

        else:
            os.rmdir(PATH_FACE + self.name)
            print("Carpeta vacía eliminada", PATH_FACE + self.name)
        self.initData()

    def OnFinishRegisterClicked(self, event):
        self.OnFinishRegister()
        pass

    def punchcard_cap(self, event):

        # Llame a la función que establece el tiempo de trabajo y juzgue si llega tarde de acuerdo con la hora actual y el tiempo de trabajo

        self.cap = cv2.VideoCapture(0)

        self.loadDataBase(5)
        print("la longitud es")
        print(len(working_times))
        if len(working_times) == 0:
            win32api.MessageBox(0, "No ha configurado el tiempo de trabajo, configure el tiempo de trabajo primero y luego configure el tiempo de trabajo", "recordar", win32con.MB_ICONWARNING)
            self.start_punchcard.Enable(True)
            self.end_puncard.Enable(False)
        else:
            working = working_times[0]
            print("-----------")
            print(working)
            offworking = offworking_times[0]
            print("-----------")
            print(offworking)
            while self.cap.isOpened():

                flag, im_rd = self.cap.read()

                kk = cv2.waitKey(1)

                dets = detector(im_rd, 1)


                if len(dets) != 0:
                    biggest_face = dets[0]

                    maxArea = 0
                    for det in dets:
                        w = det.right() - det.left()
                        h = det.top() - det.bottom()
                        if w * h > maxArea:
                            biggest_face = det
                            maxArea = w * h


                    cv2.rectangle(im_rd, tuple([biggest_face.left(), biggest_face.top()]),
                                  tuple([biggest_face.right(), biggest_face.bottom()]),
                                  (255, 0, 255), 2)
                    img_height, img_width = im_rd.shape[:2]
                    image1 = cv2.cvtColor(im_rd, cv2.COLOR_BGR2RGB)
                    pic = wx.Bitmap.FromBuffer(img_width, img_height, image1)

                    self.bmp.SetBitmap(pic)


                    shape = predictor(im_rd, biggest_face)
                    features_cap = facerec.compute_face_descriptor(im_rd, shape)

                    # Para una cara, recorre todas las características faciales almacenadas
                    for i, knew_face_feature in enumerate(self.knew_face_feature):

                        compare = return_euclidean_distance(features_cap, knew_face_feature)
                        if compare == "similar":  # encontrado caras similares
                            print("similar")
                            flag = 0
                            nowdt = self.getDateAndTime()
                            for j, logcat_name in enumerate(self.logcat_name):
                                if logcat_name == self.knew_name[i] and nowdt[0:nowdt.index(" ")] == \
                                        self.logcat_datetime[
                                            j][
                                        0:self.logcat_datetime[
                                            j].index(" ")]:
                                    self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                             + " Nombre:" + self.knew_name[i] + " Error al iniciar sesión, inicio de sesión repetido\r\n")
                                    speak_info(self.knew_name[i] + "Error al iniciar sesión, inicio de sesión repetido ")
                                    flag = 1
                                    break

                            if flag == 1:
                                break

                            if nowdt[nowdt.index(" ") + 1:-1] <= working:
                                self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                         + " Nombre:" + self.knew_name[i] + " Ingresó correctamente y no llegó tarde\r\n")
                                speak_info(self.knew_name[i] + " Iniciar sesión con éxito ")
                                self.insertARow([self.knew_id[i], self.knew_name[i], nowdt, "No"], 2)
                                con = pymysql.connect(db='baseinf', user='root', passwd='', host='localhost', port=3306,
                                                      autocommit=True)
                                cur = con.cursor()
                                sql = "INSERT INTO `logcat` (`datetime`, `id`, `name`, `late`) VALUES (%s, %s, %s, %s)"
                                cur.execute(sql, (nowdt,self.knew_id[i], self.knew_name[i], "No"))
                            elif offworking >= nowdt[nowdt.index(" ") + 1:-1] >= working:
                                self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                         + " Nombre:" + self.knew_name[i] + " Ingresó con éxito, pero llegó tarde\r\n")
                                speak_info(self.knew_name[i] + " Ingresó con éxito, pero llegó tarde")
                                self.insertARow([self.knew_id[i], self.knew_name[i], nowdt, "Si"], 2)
                                con = pymysql.connect(db='baseinf', user='root', passwd='', host='localhost', port=3306,
                                                      autocommit=True)
                                cur = con.cursor()
                                sql = "INSERT INTO `logcat` (`datetime`, `id`, `name`, `late`) VALUES (%s, %s, %s, %s)"
                                cur.execute(sql, (nowdt, self.knew_id[i], self.knew_name[i], "Si"))
                            elif nowdt[nowdt.index(" ") + 1:-1] > offworking:
                                self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                         + " Nombre:" + self.knew_name[i] + " Error al iniciar sesión, excediendo el tiempo de inicio de sesión\r\n")
                                speak_info(self.knew_name[i] + " Error al iniciar sesión, horas extra ")
                            self.loadDataBase(2)
                            break

                    if self.start_punchcard.IsEnabled():
                        self.bmp.SetBitmap(wx.Bitmap(self.pic_index))
                        _thread.exit()


    def OnStartPunchCardClicked(self, event):
        self.start_punchcard.Enable(False)
        self.end_puncard.Enable(True)
        self.loadDataBase(2)
        threading.Thread(target=self.punchcard_cap, args=(event,)).start()
        pass

    def OnEndPunchCardClicked(self, event):
        self.start_punchcard.Enable(True)
        self.end_puncard.Enable(False)
        pass

    #Prueba
    def salida_cap(self, event):

        # Llame a la función que establece el tiempo de trabajo y juzgue si llega tarde de acuerdo con la hora actual y el tiempo de trabajo

        self.cap = cv2.VideoCapture(0)

        self.loadDataBase(5)
        print("la longitud es")
        print(len(working_times))
        if len(working_times) == 0:
            win32api.MessageBox(0, "No ha configurado el tiempo de trabajo, configure el tiempo de trabajo primero y luego configure el tiempo de trabajo", "recordar", win32con.MB_ICONWARNING)
            self.salida_punchcard.Enable(True)
            self.salida2_puncard.Enable(False)
        else:
            working = working_times[0]
            print("-----------")
            print(working)
            offworking = offworking_times[0]
            print("-----------")
            print(offworking)
            while self.cap.isOpened():

                flag, im_rd = self.cap.read()

                kk = cv2.waitKey(1)

                dets = detector(im_rd, 1)


                if len(dets) != 0:
                    biggest_face = dets[0]

                    maxArea = 0
                    for det in dets:
                        w = det.right() - det.left()
                        h = det.top() - det.bottom()
                        if w * h > maxArea:
                            biggest_face = det
                            maxArea = w * h


                    cv2.rectangle(im_rd, tuple([biggest_face.left(), biggest_face.top()]),
                                  tuple([biggest_face.right(), biggest_face.bottom()]),
                                  (255, 0, 255), 2)
                    img_height, img_width = im_rd.shape[:2]
                    image1 = cv2.cvtColor(im_rd, cv2.COLOR_BGR2RGB)
                    pic = wx.Bitmap.FromBuffer(img_width, img_height, image1)

                    self.bmp.SetBitmap(pic)


                    shape = predictor(im_rd, biggest_face)
                    features_cap = facerec.compute_face_descriptor(im_rd, shape)

                    # Para una cara, recorre todas las características faciales almacenadas
                    for i, knew_face_feature in enumerate(self.knew_face_feature):

                        compare = return_euclidean_distance(features_cap, knew_face_feature)
                        if compare == "similar":  # encontrado caras similares
                            print("similar")
                            flag = 0
                            nowdt = self.getDateAndTime()
                            for j, logcat_name in enumerate(self.logcat_name):
                                if logcat_name == self.knew_name[i] and nowdt[0:nowdt.index(" ")] is \
                                        self.logcat_datetime[
                                            j][
                                        0:self.logcat_datetime[
                                            j].index(" ")]:
                                    self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                             + " Nombre:" + self.knew_name[i] + " No a iniciado sesion\r\n")
                                    speak_info(self.knew_name[i] + "No a iniciado sesion ")
                                    flag = 1
                                    break

                            if flag == 1:
                                break

                            if nowdt[nowdt.index(" ") + 1:-1] <= working:
                                self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                         + " Nombre:" + self.knew_name[i] + " Ingresó correctamente y no llegó tarde\r\n")
                                speak_info(self.knew_name[i] + " Iniciar sesión con éxito ")
                                self.insertARow([self.knew_id[i], self.knew_name[i], nowdt, "No"], 3)
                                con = pymysql.connect(db='baseinf', user='root', passwd='', host='localhost', port=3306,
                                                      autocommit=True)
                                cur = con.cursor()
                                sql = "INSERT INTO `logsalida` (`datetime`, `id`, `name`, `late`) VALUES (%s, %s, %s, %s)"
                                cur.execute(sql, (nowdt,self.knew_id[i], self.knew_name[i], "No"))
                            elif offworking >= nowdt[nowdt.index(" ") + 1:-1] >= working:
                                self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                         + " Nombre:" + self.knew_name[i] + " Ingresó con éxito, pero llegó tarde\r\n")
                                speak_info(self.knew_name[i] + " Ingresó con éxito, pero llegó tarde")
                                self.insertARow([self.knew_id[i], self.knew_name[i], nowdt, "Si"], 3)
                                con = pymysql.connect(db='baseinf', user='root', passwd='', host='localhost', port=3306,
                                                      autocommit=True)
                                cur = con.cursor()
                                sql = "INSERT INTO `logsalida` (`datetime`, `id`, `name`, `late`) VALUES (%s, %s, %s, %s)"
                                cur.execute(sql, (nowdt, self.knew_id[i], self.knew_name[i], "Si"))
                            elif nowdt[nowdt.index(" ") + 1:-1] > offworking:
                                self.infoText.AppendText(nowdt + "Número de empleo:" + str(self.knew_id[i])
                                                         + " Nombre:" + self.knew_name[i] + " Error al iniciar sesión, excediendo el tiempo de inicio de sesión\r\n")
                                speak_info(self.knew_name[i] + " Error al iniciar sesión, horas extra ")
                            self.loadDataBase(2)
                            break
                    if self.start_punchcard.IsEnabled():
                        self.bmp.SetBitmap(wx.Bitmap(self.pic_index))
                        _thread.exit()


    def OnSalidaPunchCardClicked(self, event):
        self.salida_punchcard.Enable(False)
        self.salida2_puncard.Enable(True)
        self.loadDataBase(2)
        threading.Thread(target=self.salida_cap, args=(event,)).start()
        pass

    def OnSalida2PunchCardClicked(self, event):
        self.salida_punchcard.Enable(True)
        self.salida2_puncard.Enable(False)
        pass
    #Fin PRueba

    def initInfoText(self):

        resultText = wx.StaticText(parent=self, pos=(10, 20), size=(90, 60))
        resultText.SetBackgroundColour(wx.GREEN)

        self.info = "\r\n" + self.getDateAndTime() + "Inicialización del programa exitosa\r\n"

        self.infoText = wx.TextCtrl(parent=self, size=(420, 500),
                                    style=(wx.TE_MULTILINE | wx.HSCROLL | wx.TE_READONLY))

        self.infoText.SetForegroundColour('Blue')
        self.infoText.SetLabel(self.info)
        font = wx.Font()
        font.SetPointSize(10)
        font.SetWeight(wx.BOLD)
        font.SetUnderlined(True)

        self.infoText.SetFont(font)
        self.infoText.SetBackgroundColour('WHITE')
        pass

    def initGallery(self):
        self.pic_index = wx.Image("drawable/inicio.jpg", wx.BITMAP_TYPE_ANY).Scale(600, 500)
        self.bmp = wx.StaticBitmap(parent=self, pos=(420, 0), bitmap=wx.Bitmap(self.pic_index))
        pass

    def getDateAndTime(self):
        dateandtime = strftime("%Y-%m-%d %H:%M:%S", localtime())
        return "[" + dateandtime + "]"


    def initDatabase(self):
        conn = sqlite3.connect("inspurer.db")
        cur = conn.cursor()
        cur.execute('''create table if not exists worker_info
        (name text not null,
        id int not null primary key,
        face_feature array not null)''')
        cur.execute('''create table if not exists logcat
         (datetime text not null,
         id int not null,
         name text not null,
         late text not null)''')
        cur.execute('''create table if not exists time
         (id int
		constraint table_name_pk
			primary key,
         working_time time not null,
         offwork_time time not null)''')
        cur.close()
        conn.commit()
        conn.close()

    def adapt_array(self, arr):
        out = io.BytesIO()
        np.save(out, arr)
        out.seek(0)

        dataa = out.read()
        # 压缩数据流
        return sqlite3.Binary(zlib.compress(dataa, zlib.Z_BEST_COMPRESSION))
        #return pymysql.Binary(zlib.compress(dataa, zlib.Z_BEST_COMPRESSION))

    def adapt_array_prueba(self, arr):
        out = io.BytesIO()
        np.save(out, arr)
        out.seek(0)

        dataa = out.read()
        # 压缩数据流
        #return sqlite3.Binary(zlib.compress(dataa, zlib.Z_BEST_COMPRESSION))
        return pymysql.Binary(zlib.compress(dataa, zlib.Z_BEST_COMPRESSION))


    def convert_array(self, text):
        out = io.BytesIO(text)
        out.seek(0)

        dataa = out.read()
        # 解压缩数据流
        out = io.BytesIO(zlib.decompress(dataa))
        return np.load(out)

    def insertARow(self, Row, type):
        conn = sqlite3.connect("inspurer.db")
        cur = conn.cursor()
        if type == 1:
            cur.execute("insert into worker_info (id,name,face_feature) values(?,?,?)",
                        (Row[0], Row[1], self.adapt_array(Row[2])))
            print("Escribir datos faciales con éxito")
        if type == 2:
            cur.execute("insert into logcat (id,name,datetime,late) values(?,?,?,?)",
                        (Row[0], Row[1], Row[2], Row[3]))
            print("registro de escritura exitoso")
        if type == 3:
            cur.execute("insert into logsalida (id,name,datetime,late) values(?,?,?,?)",
                        (Row[0], Row[1], Row[2], Row[3]))
            print("registro de escritura exitoso")
            pass
        cur.close()
        conn.commit()
        conn.close()
        pass

    def loadDataBase(self, type):
        nowday = self.getDateAndTime()
        day = nowday[0:nowday.index(" ")]
        print(day)
        global logcat_id, logcat_name, logcat_datetime, logcat_datetimeSa, logcat_late, working_times, offworking_times
        conn = sqlite3.connect("inspurer.db")

        cur = conn.cursor()  

        if type == 1:
            self.knew_id = []
            self.knew_name = []
            self.knew_face_feature = []
            cur.execute('select id,name,face_feature from worker_info')
            origin = cur.fetchall()
            for row in origin:
                print(row[0])
                self.knew_id.append(row[0])
                print(row[1])
                self.knew_name.append(row[1])
                print(self.convert_array(row[2]))
                self.knew_face_feature.append(self.convert_array(row[2]))
        if type == 2:
            self.logcat_id = []
            self.logcat_name = []
            self.logcat_datetime = []
            self.logcat_datetimeSa = []
            self.logcat_late = []
            #cur.execute('select id,name,datetime,late from logcat')
            cur.execute('select logcat.id,logcat.name,logcat.datetime,logsalida.datetime,logcat.late from logcat left join logsalida on logcat.id = logsalida.id')
            origin = cur.fetchall()
            for row in origin:
                print(row[0])
                self.logcat_id.append(row[0])
                print(row[1])
                self.logcat_name.append(row[1])
                print(row[2])
                self.logcat_datetime.append(row[2])
                print(row[3])
                self.logcat_datetimeSa.append(row[3])
                print(row[4])
                self.logcat_late.append(row[4])
        if type == 3:
            logcat_id = []
            logcat_name = []
            logcat_datetime = []
            logcat_late = []
            s = "'"
            sql = 'select w.id,w.name,l.datetime,l.late from worker_info w left join logcat l  on  w.id=l.id and l.datetime like' + ' ' + s + day + '%' + s + ' ' + 'order by datetime desc'
            print(sql)
            cur.execute(sql)
            origin = cur.fetchall()
            for row in origin:
                print(row[0])
                logcat_id.append(row[0])
                print(row[1])
                logcat_name.append(row[1])
                print(row[2])
                logcat_datetime.append(row[2])
                print(row[3])
                logcat_late.append(row[3])
        if type == 4:
            sql = 'select working_time from time'
            cur.execute(sql)
            countResult = (cur.fetchall())
            print(countResult)
            str = "'"
            if not countResult:
                sql = 'insert into time (id,working_time,offworking_time) values (1,' + str + working + str + ',' + str + offworking + str + ')'
                cur.execute(sql)
                print(sql)
                conn.commit()
                print("Hora de inserción exitosa")
            else:
                str="'"
                sql = 'update time set working_time=' + str + working + str + ',offworking_time=' + str + offworking + str + ' where id=1'
                cur.execute(sql)
                conn.commit()
                print(sql)
                print("Hora de actualización exitosa")

        if type==5:
            sql = 'select working_time,offworking_time from time'
            cur.execute(sql)
            print(sql)
            origin = cur.fetchall()
            print(origin)
            working_times = []
            offworking_times = []
            for row in origin:
                print("Este es el tiempo de trabajo cuando se recupera la base de datos")
                print(row[0])
                working_times.append(row[0])
                print("Este es el horario de descanso cuando se recupera la base de datos")
                print(row[1])
                offworking_times.append(row[1])
        cur.close()
        conn.commit()
        conn.close()
        pass


app = wx.App()
frame = WAS()
frame.Show()
app.MainLoop()
