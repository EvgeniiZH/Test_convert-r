import csv
import tkinter as tk
from tkinter import filedialog
import openpyxl


def open_window():
    """Инициализируем tkinter и возвращаем активный лист Excel."""
    root = tk.Tk()
    root.withdraw()  # Прячем основное окно
    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=(("Excel файлы", "*.xlsx;*.xls"), ("Все файлы", "*.*"))
    )
    if not file_path:
        print("Файл не выбран.")
        return None
    workbook = openpyxl.load_workbook(file_path)
    return workbook.active


def process_sheet(sheet):
    """
    Обрабатывает данные из листа Excel и возвращает списки вопросов, ответов и пояснений.
    """
    text_list = [row for row in sheet.iter_rows(values_only=True)]
    questions, answers, explanations = [], [], []
    current_question, current_answers, current_explanations = None, [], []
    for row in text_list:
        if row[0] == 'Вопрос':
            if current_question is not None:
                questions.append(current_question)
                answers.append(current_answers)
                explanations.append(" ".join(current_explanations))
            current_question = row[1]
            current_answers, current_explanations = [], []
        elif row[0] == 'Ответ':
            answer = str(row[1]) + '*' if row[2] == 1 else str(row[1])
            current_answers.append(answer)
        elif row[0] == 'Пояснение':
            current_explanations.append(row[1])
    if current_question is not None:
        questions.append(current_question)
        answers.append(current_answers)
        explanations.append(" ".join(current_explanations))
    return questions, answers, explanations


def determine_format(answers):
    """Определяет формат вопроса (одиночный или множественный выбор)."""
    return [
        'multiple_choice' if sum(1 for ans in answer_set if ans.endswith('*')) >= 2 else 'single_choice'
        for answer_set in answers
    ]


def generate_tegs(explanations):
    """Создает HTML-теги для пояснений."""
    return [
        f'<details><summary class="coursqestion">Подсказка</summary>{"".join(explanation)}</details>'
        for explanation in explanations
    ]

def save_to_csv(lighthouse, list_of_dicts):
    """Сохраняет обработанные данные в CSV-файл."""
    file_name = 'file_name.csv'
    with open(file_name, mode='w', encoding='utf-8', newline='') as csv_file:
        writer = csv.writer(csv_file, delimiter=',')
        key = ["settings", "Экзамен" if lighthouse == 0 else "Подготовка", "", "0", "minutes", "",
               "10", "80", "1000", "", "question_below_each_other", "asc", "1", "200"]
        seting_quest = ["1.00", "0", "1", "", ""]
        seting_answer_p = ["text", "1"]
        seting_answer_n = ["text"]
        writer.writerow(key)
        for dictionary in list_of_dicts:
            question = ['question', dictionary['question'], dictionary['Правила'], dictionary['Формат'], *seting_quest,
                        dictionary['Правила без тега']]

            writer.writerow(question)
            for ans in dictionary['answer']:
                v = 1
                correct = ans.endswith('*')
                if v == 1:
                    answer_text = ans
                    answer = ["answer", answer_text, *seting_answer_p] if correct else ["answer", answer_text,
                                                                                        *seting_answer_n]
                    writer.writerow(answer)
                else:
                    answer_text = ans.rstrip('*')
                    answer = ["answer", answer_text, *seting_answer_p] if correct else ["answer", answer_text,
                                                                                        *seting_answer_n]
                    writer.writerow(answer)


def main():
    sheet = open_window()
    if not sheet:
        return
    questions, answers, explanations = process_sheet(sheet)
    format_quest = determine_format(answers)
    teg_explanations = generate_tegs(explanations)
    list_of_dicts = [
        {'question': question, 'answer': ans, 'Правила': tegs, 'Формат': fmt, 'Правила без тега': pr_n_tg}
        for question, ans, tegs, fmt, pr_n_tg in zip(questions, answers, teg_explanations, format_quest, explanations)
    ]
    lighthouse = 0  # Временный флаг, можете заменить его другим значением, если требуется
    save_to_csv(lighthouse, list_of_dicts)


if __name__ == "__main__":
    main()
