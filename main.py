import subprocess
from tkinter import *
import os
from urllib.request import urlopen
import wmi
from ping3 import ping
import winreg
import threading
import speedtest

font = ("Arial", 12)
label_style = {"font": font, "width": 80, "height": 1}
btn_style_1 = {"font": font, "width": 50, "height": 1}
btn_style_2 = {"font": font, "width": 30, "height": 2}
txt_style = {"font": font, "width": 50, "height": 1, "bg": "white"}

label_pack = {"side": "top", "padx": (5, 5), "pady": (5, 5)}
btn_pack_1 = {"side": "left", "padx": (5, 5), "pady": (5, 5)}
btn_pack_2 = {"side": "top", "padx": (5, 5), "pady": (5, 5)}
txt_pack = {"side": "left", "padx": (5, 5), "pady": (5, 5)}



class Window(Tk):
    def __init__(self):
        super(Window, self).__init__()
        self.title("Программная проверка информационной безопасности")
        self.geometry("1200x800")
        self.resizable(False, False)
        top_frame = Frame(self.master)
        Label(top_frame, text="Проверка межсетевого экрана", **label_style).pack(**label_pack)

        frame = Frame(top_frame)
        Button(frame, text="Проверка подключения к Интернету",
               command=self.check_connect, **btn_style_1).pack(**btn_pack_1)
        self.txt_ping = Label(frame, **txt_style)
        # поле для вывода аднных о подключение к интернету.
        # словарь txt_style распаковывается как набор аргументов в функцию.
        # Так сделано, чтоб стили задавать в начале файла и задавать ихх всем элементах формы,
        # а не каждому по отдельности
        self.txt_ping.pack(**txt_pack)  # добавить элемент на форму
        frame.pack(side="top")

        internet_speed_frame = Frame(top_frame)
        Button(internet_speed_frame, text="Проверить скорость интернета",
               command=self.check_internet_speed, **btn_style_1).pack(**btn_pack_1)
        self.txt_internet_speed = Label(internet_speed_frame, **txt_style)
        self.txt_internet_speed.pack(**txt_pack)
        internet_speed_frame.pack(side="top")



        frame = Frame(top_frame)
        Button(frame, text="Проверка наличия установленного межсетевого экрана",
               command=self.check_current_connect, **btn_style_1).pack(**btn_pack_1)
        self.txt_furewall = Label(frame, **txt_style)
        self.txt_furewall.pack(**txt_pack)
        frame.pack(side="top")

        frame = Frame(top_frame)
        Button(frame, text="Проверка работоспособности межсетевого экрана",
               command=self.check_work_connect, **btn_style_1).pack(**btn_pack_1)
        self.txt_work_connect = Label(frame, **txt_style)
        self.txt_work_connect.pack(**txt_pack)
        frame.pack(side="top")

        top_frame.pack(side="top")

        middle_frame = Frame(self.master)
        Label(top_frame, text="Проверка межсетевого экрана", **label_style).pack(**label_pack)

        frame = Frame(middle_frame)
        Button(frame, text="Проверка наличия установленного антивируса",
               command=self.check_install_antivirus, **btn_style_1).pack(**btn_pack_1)
        self.txt_install_anti = Label(frame, **txt_style)
        self.txt_install_anti.pack(**txt_pack)
        frame.pack(side="top")

        frame = Frame(middle_frame)
        Button(frame, text="Проверка работоспособности антивирусного ПО",
               command=self.check_work_antivirus, **btn_style_1).pack(**btn_pack_1)
        self.txt_work_anti = Label(frame, **txt_style)
        self.txt_work_anti.pack(**txt_pack)
        frame.pack(side="top")

        frame_updates = Frame(middle_frame)
        Button(frame_updates, text="Проверить обновления Windows",
               command=self._check_windows_updates, **btn_style_1).pack(**btn_pack_1)
        self.txt_windows_update = Text(frame_updates, font=font, height=10, width=50)  # Label to display the updates status
        self.txt_windows_update.pack(**txt_pack)
        frame_updates.pack(side="left")


        middle_frame.pack(side="top")

        bottom_frame = Frame(self.master)
        frame = Frame(bottom_frame)
        Label(frame, text="Результат проверок и рекомендаций", **label_style).pack(**label_pack)
        self.txt_out = Text(frame, font=font, height=10, width=75)
        self.txt_out.pack(**label_pack)
        frame.pack(side="left")

        frame = Frame(bottom_frame)
        Button(frame, text="Вывести результаты",
                command=self.get_result, **btn_style_2).pack(**btn_pack_2)
        Button(frame, text="Сохранить результаты в файл",
               command=self.save_to_file, **btn_style_2).pack(**btn_pack_2)
        Button(frame, text="Выход",
               command=self.exit, **btn_style_2).pack(**btn_pack_2)
        frame.pack(side="left")
        bottom_frame.pack(side="top")

    def check_connect(self):   # функция для проверки подключения к интернету
        hostname = "ya.ru"
        response = ping(hostname)
        result = ""
        if response == False:
            result = "Данный компьютер не подключен к интернету"
        else:
            result = "Данный компьютер подключен к интернету"
        self.txt_ping["text"] = result  # задаем результат
        return result

    def _check_windows_updates(self):
        try:
            script = """
            $UpdateSession = New-Object -ComObject Microsoft.Update.Session
            $UpdateSearcher = $UpdateSession.CreateupdateSearcher()
            $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
            $resultText = @()
            foreach ($update in $SearchResult.Updates) {
                $title = $update.Title
                $desc = $update.Description
                $isDownloaded = $update.IsDownloaded
                $isMandatory = $update.IsMandatory
                $updateInfo = "$title`n$desc`nDownloaded: $isDownloaded, Mandatory: $isMandatory`n`n"
                $resultText += $updateInfo
            }
            $resultText -join "`n"
            """
            process = subprocess.Popen(["powershell", script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            result, error = process.communicate()

            # Используйте CP1251 для декодирования
            if error:
                self.txt_windows_update.delete(1.0, "end")  # Очистить текст в Text виджете
                self.txt_windows_update.insert("end", f"Ошибка: {error.decode('utf-8')}")
                return result
            self.txt_windows_update.delete(1.0, "end")  # Очистить текст в Text виджете
            self.txt_windows_update.insert("end", result.decode('utf-8'))
            return result
        except Exception as e:
            self.txt_windows_update.delete(1.0, "end")  # Очистить текст в Text виджете
            self.txt_windows_update.insert("end", f"Ошибка при проверке обновлений: {str(e)}")

    def get_text_from_txt_windows_update(self):
        self._check_windows_updates()
        text_contents = self.txt_windows_update.get(1.0, "end-1c")
        return text_contents
    def check_internet_speed(self):
        # Отображение всплывающего окна с сообщением
        self.popup = Toplevel(self)
        Label(self.popup, text="Идет измерение скорости интернета. Пожалуйста, подождите...").pack()
        self.popup.geometry("500x200")

        # Запуск измерения скорости в отдельном потоке
        threading.Thread(target=self.perform_speed_test, daemon=True).start()

    def perform_speed_test(self):
        try:
            st = speedtest.Speedtest()
            st.get_best_server()
            download_speed = st.download()
            upload_speed = st.upload()
            result = f"Скорость загрузки: {download_speed / 1024 / 1024:.2f} Mbps "
            result += f"Скорость выгрузки: {upload_speed / 1024 / 1024:.2f} Mbps"
            self.txt_internet_speed["text"] = result
        except Exception as e:
            result = f"Ошибка при проверке скорости интернета: {str(e)}\n"
            self.txt_internet_speed["text"] = result
        return result
    def check_current_connect(self):
        path = "C:\\Program Files\\DrWeb\\frwl_svc.exe"
        result = ""
        if os.path.exists(path):
            result = "Фаервол установлен!"
        else:
            result = "Фаервол не установлен!"
        self.txt_furewall["text"] = result
        return result

    def check_work_connect(self):   # проверка работы фаервола
        result = ""
        try:
            urlopen("http://google.com").read().decode()       # пробуем подключиться и получить ответ
            result = "Межсетевой экран функционирует неверно, или не функционирует вовсе!" # фаервол не сработал и не запретил соединение
        except Exception as ex:
            print(ex)
            result = "Межсетевой экран функционирует правильно!" # программа вылетела в ошибку, значит фаервол сработал и запретил запрос
        self.txt_work_connect["text"] = result
        return result

    def check_install_antivirus(self):
        res = "Антивирус не установлен!"
        antiv = ['Dr.Web Security Space', 'Kaspersky Internet Security', 'Kaspersky Total Security', 'Avast Internet Security', 'Avast! Pro Antivirus']
        aReg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
        aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                              0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        count_subkey = winreg.QueryInfoKey(aKey)[0]
        software_list = list()
        for i in range(count_subkey):
            software = {}
            try:
                asubkey_name = winreg.EnumKey(aKey, i)
                asubkey = winreg.OpenKey(aKey, asubkey_name)
                software['name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]
                software_list.append(software)
                if software['name'] in antiv:
                    res = 'Антивирус установлен!'
                    break
            except EnvironmentError:
                continue
        self.txt_install_anti["text"] = res
        return res


    def check_work_antivirus(self):
        f = wmi.WMI()   # получаем доступ ко всем процессорами на компьютере
        result = "Антивирус не работает!"
        for process in f.Win32_Process():   # перебираем все запущенные процессы на компьютере
            if process.Name == "spideragent.exe":   # в примере на C# искали именно этот процессс
                result = "Антивирус работает!"
                break
        self.txt_work_anti["text"] = result
        return result


    def get_result(self):
        # проверить все тесты
        result = ""
        result += "1. " + self.check_connect() + "\n"
        result += "2. " + self.perform_speed_test() + "\n"
        result += "3. " + self.check_current_connect() + "\n"
        result += "4. " + self.check_work_connect() + "\n"
        result += "5. " + self.check_install_antivirus() + "\n"
        result += "6. " + self.check_work_antivirus() + "\n"
        result += "7. " + self.get_text_from_txt_windows_update() + "\n"
        self.txt_out.delete(1.0, END)   # необходимо для очистки поля перед выводом нового отчета.
        self.txt_out.insert(1.0, result)

    def save_to_file(self):
        result = ""
        result += "1. " + self.check_connect() + "\n"
        result += "2. " + self.perform_speed_test() + "\n"
        result += "3. " + self.check_current_connect() + "\n"
        result += "4. " + self.check_work_connect() + "\n"
        result += "5. " + self.check_install_antivirus() + "\n"
        result += "6. " + self.check_work_antivirus() + "\n"
        result += "7. " + self.get_text_from_txt_windows_update() + "\n"
        with open("Курсовая работа.txt", "w") as file:
            file.write(result)

    def exit(self):
        self.destroy()


wnd = Window()
wnd.mainloop()
