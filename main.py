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

class Tester:
    last = "тест не выполнялся"
    
    def do_test(self):
        pass
    
    def on_update(self, setter):
        def f(msg):
            self.last = msg
            setter(msg)
        self.updater = f
    
    def get_last(self):
        return self.last

class ConnectionTester(Tester):
    def do_test(self):
        hostname = "ya.ru"
        response = ping(hostname)
        self.updater("Данный компьютер " + ( response if "" else "НЕ" ) + " подключен к Интернету")

class SpeedTester(Tester):
    def test_worker(self):
        msg = ""
        try:
            st = speedtest.Speedtest()
            st.get_best_server()
            download_speed = st.download()
            upload_speed = st.upload()
            msg = f"Скорость загрузки: {download_speed / 1024 / 1024:.2f} Mbps "
            msg += f"Скорость выгрузки: {upload_speed / 1024 / 1024:.2f} Mbps"
        except Exception as e:
            msg = f"Ошибка при проверке скорости интернета: {str(e)}\n"
        self.updater(msg)
        self.popup.destroy()
    
    def do_test(self):
        self.popup = Toplevel()
        self.popup.title("Измерение скорости Интернет-соединения")
        Label(self.popup, text="Идет измерение скорости Интернет-соединения. Пожалуйста, подождите...").pack()
        self.popup.geometry("500x200")
        
        # Запуск измерения скорости в отдельном потоке
        threading.Thread(target=self.test_worker, daemon=True).start()

class FirewallPresenceTester(Tester):
    def do_test(self):
        path = "C:\\Program Files\\DrWeb\\frwl_svc.exe"
        self.updater("Межсетевой экран " + (os.path.exists(path) if "" else "НЕ") + " установлен" )

class FirewallWorkingTester(Tester):
    def do_test(self):
        try:
            urlopen("http://google.com").read().decode()
        except Exception as ex:
            self.updater("Межсетевой экран функционирует правильно!")
        else:
            self.updater("Межсетевой экран не функционирует")

class AntivirusPresenceTester(Tester):
    def do_test(self):
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
        self.updater(res)

class AntivirusWorkingTester(Tester):
    def do_test(self):
        f = wmi.WMI()   # получаем доступ ко всем процессорами на компьютере
        result = "Антивирус не работает!"
        for process in f.Win32_Process():   # перебираем все запущенные процессы на компьютере
            if process.Name == "spideragent.exe":   # в примере на C# искали именно этот процессс
                result = "Антивирус работает!"
                break
        self.updater(result)

class WindowsUpdateTester(Tester):
    def do_test(self):
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
                self.updater(f"Ошибка: {error.decode('utf-8')}")
            self.updater(str(result))
        except Exception as e:
            self.updater(f"Ошибка при проверке обновлений: {str(e)}")

class Window(Tk):
    def set_connect_status(self, status):
        self.txt_ping["text"] = status
    
    def set_internet_speed(self, speed):
        self.txt_internet_speed["text"] = speed
    
    def set_firewall_presence(self, present):
        self.txt_firewall_present["text"] = present
    
    def set_firewall_working(self, working):
        self.txt_firewall_working["text"] = working
    
    def set_antivirus_presence(self, present):
        self.txt_install_anti["text"] = present
    
    def set_antivirus_working(self, working):
        self.txt_work_anti["text"] = working
    
    def set_windows_update(self, value):
        self.txt_windows_update.delete("1.0", END)
        self.txt_windows_update.insert("1.0", value)
    
    def __init__(self):
        super(Window, self).__init__()
        
        self.conn_tester = ConnectionTester()
        self.speed_tester = SpeedTester()
        self.firewall_presence_tester = FirewallPresenceTester()
        self.firewall_working_tester = FirewallWorkingTester()
        self.antivirus_presence_tester = AntivirusPresenceTester()
        self.antivirus_working_tester = AntivirusWorkingTester()
        self.windows_update_tester = WindowsUpdateTester()
        
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
        self.conn_tester.on_update(self.set_connect_status)

        internet_speed_frame = Frame(top_frame)
        Button(internet_speed_frame, text="Проверить скорость интернета",
               command=self.check_internet_speed, **btn_style_1).pack(**btn_pack_1)
        self.txt_internet_speed = Label(internet_speed_frame, **txt_style)
        self.txt_internet_speed.pack(**txt_pack)
        internet_speed_frame.pack(side="top")
        self.speed_tester.on_update(self.set_internet_speed)

        frame = Frame(top_frame)
        Button(frame, text="Проверка наличия установленного межсетевого экрана",
               command=self.check_current_connect, **btn_style_1).pack(**btn_pack_1)
        self.txt_firewall_present = Label(frame, **txt_style)
        self.txt_firewall_present.pack(**txt_pack)
        frame.pack(side="top")
        self.firewall_presence_tester.on_update(self.set_firewall_presence)

        frame = Frame(top_frame)
        Button(frame, text="Проверка работоспособности межсетевого экрана",
               command=self.check_work_connect, **btn_style_1).pack(**btn_pack_1)
        self.txt_firewall_working = Label(frame, **txt_style)
        self.txt_firewall_working.pack(**txt_pack)
        frame.pack(side="top")
        self.firewall_working_tester.on_update(self.set_firewall_working)

        top_frame.pack(side="top")

        middle_frame = Frame(self.master)
        Label(top_frame, text="Проверка межсетевого экрана", **label_style).pack(**label_pack)

        frame = Frame(middle_frame)
        Button(frame, text="Проверка наличия установленного антивируса",
               command=self.check_install_antivirus, **btn_style_1).pack(**btn_pack_1)
        self.txt_install_anti = Label(frame, **txt_style)
        self.txt_install_anti.pack(**txt_pack)
        frame.pack(side="top")
        self.antivirus_presence_tester.on_update(self.set_antivirus_presence)

        frame = Frame(middle_frame)
        Button(frame, text="Проверка работоспособности антивирусного ПО",
               command=self.check_work_antivirus, **btn_style_1).pack(**btn_pack_1)
        self.txt_work_anti = Label(frame, **txt_style)
        self.txt_work_anti.pack(**txt_pack)
        frame.pack(side="top")
        self.antivirus_working_tester.on_update(self.set_antivirus_working)

        frame_updates = Frame(middle_frame)
        Button(frame_updates, text="Проверить обновления Windows",
               command=self._check_windows_updates, **btn_style_1).pack(**btn_pack_1)
        self.txt_windows_update = Text(frame_updates, font=font, height=10, width=50)  # Label to display the updates status
        self.txt_windows_update.pack(**txt_pack)
        frame_updates.pack(side="left")
        self.windows_update_tester.on_update(self.set_windows_update)

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
        self.conn_tester.do_test()

    def _check_windows_updates(self):
        self.windows_update_tester.do_test()
    
    def check_internet_speed(self):
        self.speed_tester.do_test()
    
    def check_current_connect(self):
        self.firewall_presence_tester.do_test()

    def check_work_connect(self):
        self.firewall_working_tester.do_test()

    def check_install_antivirus(self):
        self.antivirus_presence_tester.do_test()

    def check_work_antivirus(self):
        self.antivirus_working_tester.do_test()
    
    def get_total(self):
        result = ""
        result += "1. " + self.conn_tester.get_last() + "\n"
        result += "2. " + self.speed_tester.get_last() + "\n"
        result += "3. " + self.firewall_presence_tester.get_last() + "\n"
        result += "4. " + self.firewall_working_tester.get_last() + "\n"
        result += "5. " + self.antivirus_presence_tester.get_last() + "\n"
        result += "6. " + self.antivirus_working_tester.get_last() + "\n"
        result += "7. " + self.windows_update_tester.get_last() + "\n"
        return result
    
    def get_result(self):
        # проверить все тесты
        result = self.get_total()
        self.txt_out.delete("1.0", END)   # необходимо для очистки поля перед выводом нового отчета.
        self.txt_out.insert("1.0", result)

    def save_to_file(self):
        result = self.get_total()
        with open("Курсовая работа.txt", "w") as file:
            file.write(result)

    def exit(self):
        self.destroy()

wnd = Window()
wnd.mainloop()