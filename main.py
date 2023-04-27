import openpyxl
import pandas as pd
import os


# Функция для чтения файла xlsx и создания объектов файлов
def read_excel_file(file_path):
    files = []
    # читаем файл xlsx
    df = pd.read_excel(file_path)
    for i in range(len(df)):
        # отбрасываем ненужные поля
        if df.loc[i, 'File'] != 'Fill (Заполнение)':
            name = df.loc[i, 'File']
            size = df.loc[i, 'File Size']
            start_sector = df.loc[i, 'Start Sector']
            end_sector = df.loc[i, 'End Sector']
            # формируем файл
            files.append(
                {'name': name, 'size': int(size.replace(',', '')), 'start_sector': int(start_sector.replace(',', '')),
                 'end_sector': int(end_sector.replace(',', '')),
                 'file_name': file_path.replace('.xlsx', '.dd')})
    return files


# Функция для сбора файлов вместе
def merge_files(df, out_file):
    # Отсортировать DataFrame по названию и начальному сектору
    df = df.sort_values(by=['file_name', 'start_sector'])
    # Открыть выходной файл для записи
    with open(out_file, 'wb') as out:
        for i, row in df.iterrows():
            if out_file in row['name']:
                # Открыть файл на чтение
                with open(row['file_name'], 'rb') as f:
                    # Поставить позицию на сектор
                    f.seek((row['start_sector']) * 512)
                    # Прочитать и записать содержимое файла
                    out.write(f.read(row['size']))


# Главный метод
if __name__ == '__main__':
    # Прочитать xlsx файлы в DataFrame
    files = []
    for filename in os.listdir('.'):
        if filename.endswith('.xlsx'):
            files += read_excel_file(filename)
    df = pd.DataFrame(files)

    # Извлекает названия файлов, в которые будем сохранять
    name_file = []
    for file_ot in df['name']:
        naming_file = file_ot.split('(')[0].strip()
        name_file.append(naming_file)

    name_file = list(set(name_file))
    # Собрать файлы вместе
    for name_filee in name_file:
        merge_files(df, name_filee)
