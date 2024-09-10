import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def load_and_filter_data(file_path):

    df = pd.read_excel(file_path)

    print(f"Заголовки столбцов в файле {file_path}: {df.columns.tolist()}")

    if 'Unnamed: 2' not in df.columns:
        raise KeyError(f"Столбец 'Unnamed: 2' не найден в файле {file_path}")

    df['Unnamed: 2'] = pd.to_numeric(df['Unnamed: 2'], errors='coerce')

    df = df.dropna(subset=['Unnamed: 2'])

    filtered_values = df[df['Unnamed: 2'] < -0.1]['Unnamed: 2']

    return filtered_values


def process_directory(directory_path, pdf_path):

    pdf = PdfPages(pdf_path)

    txt_file_path = os.path.splitext(pdf_path)[0] + '.txt'
    txt_file = open(txt_file_path, 'w')

    subfolders = [f.path for f in os.scandir(directory_path) if f.is_dir()]

    for folder in subfolders:
        data = []
        min_indexes = []
        file_paths = [os.path.join(folder, f"{i}.xlsx") for i in range(1, 4)]

        for file_path in file_paths:
            if os.path.exists(file_path):
                try:
                    filtered_data = load_and_filter_data(file_path)
                    data.append(filtered_data)
                except KeyError as e:
                    print(e)
            else:
                print(f"Файл не найден: {file_path}")

        if len(data) >= 2:
            for values in data:
                min_index = values.idxmin()
                min_indexes.append(min_index)

            min_index_range = f"({min(min_indexes)/1000} ; {max(min_indexes)/1000})"
            minudl = f" отн. удлинение = {min(min_indexes)/1000/37.2} ; {max(min_indexes)/1000/37.2})"
            txt_file.write(f"Диапазон для папки {os.path.basename(folder)}: {min_index_range} мм\n{minudl}%\n")

            plt.figure()

            for i, values in enumerate(data):
                x_values = [(index - values.index[0]) * 0.001666667 for index in values.index]
                y_values = values.values

                min_index = values.idxmin()
                min_value = values[min_index]

                plt.plot(x_values, y_values, label=f'Файл {i + 1} ({min_index/1000}) мм')
                plt.scatter([(min_index - values.index[0]) * 0.001666667], [min_value],
                            color='red')

                print(f"Папка: {folder}, Файл {i + 1}, Номер наименьшего значения: {min_index}")

            plt.legend(loc='lower left')

            plt.title(f"График для папки: {os.path.basename(folder)}")
            plt.xlabel('Растяжение, мм')
            plt.ylabel('Прилагаемая сила, H')

            pdf.savefig()
            plt.close()
        else:
            print(f"Недостаточно файлов для построения графика в папке: {folder}")

    pdf.close()
    print(f"Графики сохранены в файл: {pdf_path}")

    txt_file.close()
    print(f"Файл с диапазоном сохранен: {txt_file_path}")


if __name__ == "__main__":
    directory_path = r'C:\Users\Сергей\Documents\23.04.2024'
    pdf_path = os.path.join(directory_path, 'graphs.pdf')
    process_directory(directory_path, pdf_path)
