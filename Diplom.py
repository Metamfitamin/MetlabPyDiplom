import pyaudio, time, wmi, random, winreg
import speech_recognition as sr

trusted_devices = [] # пис
def get_trusted_microphones(): #создание whitelist
    p = pyaudio.PyAudio()
    num_microphones = p.get_device_count()

    for i in range(num_microphones):
        device_info = p.get_device_info_by_index(i)
        if device_info["maxInputChannels"] > 0 and device_info["hostApi"] == 1:
            trusted_devices.append(device_info["name"])

    p.terminate()
    print("Список доверенных устройств:")
    for device in trusted_devices:
        print(device)

    return trusted_devices

def compare_audiolists(trusted_devices): #Проверка usb-audio девайса на то что он микрофон и его нет в whitelist
    audiodevices = []
    p = pyaudio.PyAudio()
    num_microphones = p.get_device_count()
    for i in range(num_microphones):
        device_info = p.get_device_info_by_index(i)
        if device_info["maxInputChannels"] > 0 and device_info["maxInputChannels"] > 0 and device_info["hostApi"] == 1:
            audiodevices.append(device_info["name"])

    for item in audiodevices:
        if item not in trusted_devices:
            print("Обнаружено новое USB устройство с входным каналом ", item)
            return item
    return None

def recognize_spoken_word(): #проверка микрофона
    WORDS = ["яблоко", "банан", "груша", "апельсин", "слива"]
    word = random.choice(WORDS)
    recognizer = sr.Recognizer()
    WAIT_TIME = 30  # в секундах
    print(f"Пожалуйста, произнесите слово '{word}'.")
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source, timeout=WAIT_TIME)
    try:
        spoken_word = recognizer.recognize_google(audio, language="ru-RU").lower()
        if spoken_word == word:
            print("Вы правильно произнесли слово!")
            return True
        else:
            print(f"Вы произнесли неправильное слово. Запрашиваемое слово было '{word}', а вы произнесли '{spoken_word}'. Пожалуйста, попробуйте снова.")
            return False
    except sr.UnknownValueError:
        print("Не удалось распознать произнесенное слово. Пожалуйста, попробуйте снова.")
        return False


def checkusbaudio(): # функция непрерывного сканирования событий wmi на ивенты подключения usb-audio устройства
    raw_wql = "select * from __instancecreationevent within 2 where targetinstance isa 'win32_usbhub'"
    c = wmi.WMI()
    watcherusb = c.watch_for(raw_wql=raw_wql)
    wql = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_SoundDevice'"
    c = wmi.WMI()
    watchermic = c.watch_for(raw_wql=wql)
    while True:
        usb = watcherusb()
        mic = watchermic()
        if usb is not None and mic is not None:
            new_mic_name = compare_audiolists(trusted_devices)
            if new_mic_name is not None:
                #block_cmd_powershell()
                #block_task_manager()
                #block_regedit()
                for i in range(4):
                    if recognize_spoken_word() == True:
                        trusted_devices.append(new_mic_name)
                        print("Устройство добавлено в доверенный список\nВывожу информацию о устройстве\n", mic)
                        unblock_regedit()
                        unblock_task_manager()
                        unblock_cmd_powershell()
                        break
                    elif i == 3:
                        print("Устройство проверку не прошло. Блокирую машину.")
                    
#для тестов
#блокировка разблокировка реестра
def block_regedit():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Policies\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.SetValueEx(key, 'DisableRegistryTools', 0, winreg.REG_DWORD, 1)
    winreg.CloseKey(key)
def unblock_regedit():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Policies\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.DeleteValue(key, 'DisableRegistryTools')
    winreg.CloseKey(key)

#блокировка разблокировка диспетчера задач
def block_task_manager():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Policies\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.SetValueEx(key, 'DisableTaskMgr', 0, winreg.REG_DWORD, 1)
    winreg.CloseKey(key)
def unblock_task_manager():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Policies\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.DeleteValue(key, 'DisableTaskMgr')
    winreg.CloseKey(key)

#для финала
#блокировка cmd и powershell
def block_cmd_powershell():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Policies\Microsoft\Windows\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.SetValueEx(key, 'DisableCMD', 0, winreg.REG_DWORD, 1)
    winreg.SetValueEx(key, 'DisablePowerShell', 0, winreg.REG_DWORD, 1)
    winreg.CloseKey(key)
def unblock_cmd_powershell():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Policies\Microsoft\Windows\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.DeleteValue(key, 'DisableCMD')
    winreg.DeleteValue(key, 'DisablePowerShell')
    winreg.CloseKey(key)

def block_regedit_taskmgr():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Policies\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.SetValueEx(key, 'DisableRegistryTools', 0, winreg.REG_DWORD, 1)
    winreg.SetValueEx(key, 'DisableTaskMgr', 0, winreg.REG_DWORD, 1)
    winreg.CloseKey(key)
def unblock_regedit_taskmgr():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Policies\System', 0, winreg.KEY_ALL_ACCESS)
    winreg.DeleteValue(key, 'DisableRegistryTools')
    winreg.DeleteValue(key, 'DisableTaskMgr')
    winreg.CloseKey(key)


if __name__ == '__main__':
    get_trusted_microphones()
    checkusbaudio()