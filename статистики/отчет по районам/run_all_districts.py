import subprocess
import os

folder_path = os.path.dirname(os.path.abspath(__file__))

files_to_run = [
    "abay_oblast.py",
    "akmola_oblast.py",
    "aktobe_oblast.py",
    "almaty_city.py",
    "almaty_oblast.py",
    "astana_city.py",
    "atyrau_oblast.py",
    "kostanay_oblast.py",
    "karaganda_oblast.py",
    "kyzylorda_oblast.py",
    "mangistau_oblast.py",
    "pavlodar_oblast.py",
    "north_kazakhstan_oblast.py",
    "shymkent_city.py",
    "turkestan_oblast.py",
    "ulytau_oblast.py",
    "east_kazakhstan_oblast.py",
    "west_kazakhstan_oblast.py",
    "zhambyl_oblast.py",
    "zhetysu_oblast.py",
]

for file_name in files_to_run:
    file_path = os.path.join(folder_path, file_name)
    print(f"Запуск {file_name}...")
    subprocess.run(["python", file_path], check=True)
    print(f"{file_name} выполнен\n")

print("Все скрипты выполнены!")
