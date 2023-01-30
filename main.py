import os
import openpyxl

def get_files():
    for root, dirnames, files in os.walk("."):
        for file in files:
            yield [root, file]
class FileInfo:
    def __init__(self, root, file):
        self.root = root
        self.file_name, self.extension = os.path.splitext(file)

def save_to_excel(file_infos):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Номер строки", "Папка в которой лежит файл", "Название файла", "Расширение файла"])
        for i, file_info in enumerate(file_infos):
            sheet.append([i+1, file_info.root, file_info.file_name, file_info.extension])
        workbook.save("result.xlsx")
    except Exception as e:
        print("При сохранении файла Excel произошла ошибка:", e)

def main():
    try:
        file_infos = [FileInfo(root, file) for root, file in get_files()]
        save_to_excel(file_infos)
    except Exception as e:
        print("Произошла ошибка:", e)

if __name__ == '__main__':
    main()