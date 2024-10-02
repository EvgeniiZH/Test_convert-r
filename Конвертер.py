import csv
import tkinter as tk
from tkinter import filedialog
import openpyxl


def оpen_window():
    # Инициализируем tkinter
    root = tk.Tk()
    root.withdraw()  # Прячем основное окно

    # Открываем диалог выбора файла
    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=(("Excel файлы", "*.xlsx;*.xls"), ("Все файлы", "*.*"))
    )

    # Открываем файл Excel
    workbook = openpyxl.load_workbook(file_path)

    # Для примера, работаем с активным листом
    sheet = workbook.active
    return sheet


sheet = оpen_window()
lighthouse = 0
# Проход по строкам и столбцам
text_list = []
for row in sheet.iter_rows(values_only=True):
    text_list.append(row)

questions = []
answers = []
explanations = []

current_question = None
current_answers = []
current_explanations = []

# Перебор каждой строки в списке text_list.
for row in text_list:
    if row[0] == 'Вопрос':
        if current_question is not None:
            # Сохранение текущего набора вопросов, ответов и пояснений
            questions.append(current_question)
            answers.append(current_answers)
            explanations.append(current_explanations)

        # Инициализация новой группы вопросов, ответов и пояснений
        current_question = row[1]
        current_answers = []
        current_explanations = []

    elif row[0] == 'Ответ':
        if row[2] == 1:
            current_answers.append(row[1] + '*')
        else:
            current_answers.append(row[1])

    elif row[0] == 'Пояснение':
        current_explanations.append(row[1])

# Добавление последнего набора, если цикл завершился
if current_question is not None:
    questions.append(current_question)
    answers.append(current_answers)
    explanations.append(current_explanations)

format_quest = []
count = 0

for i in answers:
    for j in i:
        if j[-1] == '*':
            count += 1
    if count >= 2:
        format_quest.append('multiple_choice')
    else:
        format_quest.append('single_choice')
    count = 0

teg_explanations = []
for text_list in explanations:
    text = ''.join(text_list)  # Конкатенирует элементы списка в одну строку
    teg_explanations.append('<details><summary class=coursqestion>Подсказка</summary>' + text + '</details>')

list_of_dicts = []
for question, format_quest, answers, teg_explanationsе in zip(questions, format_quest, answers, teg_explanations):
    dictionary = {'question': question, 'answer': answers, 'Правила': teg_explanations, 'Формат': format_quest}
    list_of_dicts.append(dictionary)
# print(list_of_dicts, sep="\n")
# Открываем файл для записи
with open(f'file_name.csv', mode='w', encoding='utf-8', newline='') as csv_file:
    writer = csv.writer(csv_file, delimiter=',')
    if lighthouse == 0:
        # Получаем список ключей из первого словаря
        key = ["settings", "Экзамен", "", "0", "minutes", "", "10", "80", "1000",
               "", "question_below_each_other", "asc", "1", "200"]
    else:
        key = ["settings", "Подготовка", "", "0", "minutes", "", "10", "80", "1000",
               "", "question_below_each_other", "asc", "1", "200"]
    seting_quest = ["1.00", "0", "1", "", "", "<p><br data-mce-bogus=\"1\"></p>"]
    seting_answer_p = ["text", "1"]
    seting_answer_n = ["text"]
    # Записываем ключи на отдельной строке
    writer.writerow(key)
    # Записываем значения словарей на отдельные строки
    for dict in list_of_dicts:
        queston = ['question', dict['question'], *dict['Правила'], dict['Формат'], *seting_quest]
        writer.writerow(queston)
        for ans in dict['answer']:
            if lighthouse == 0:
                if ans[-1] == '*':
                    answer = ["answer", ans[:-1], *seting_answer_p]
                else:
                    answer = ["answer", ans, *seting_answer_n]
            else:
                if ans[-1] == '*':
                    answer = ["answer", ans, *seting_answer_p]
                else:
                    answer = ["answer", ans, *seting_answer_n]
            writer.writerow(answer)
# Вывод собранных вопросов, ответов и пояснений
# print(len(questions))
# print(len(answers))
# print(len(explanations))
# print("Вопросы:", questions)
# print("Ответы:", answers)
# print("Пояснения:", explanations)
# print(len(format_quest))
print(teg_explanations)
