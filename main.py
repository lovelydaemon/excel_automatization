import pandas as pd
from pandas import Series, DataFrame
import numpy as np


DOC_FILEPATH = 'sample.xlsx'
SHEET_NAME = 'Лист2'

def df_get(*columns: str) -> Series | DataFrame:
    """Возвращает серию, либо фрейм"""
    columns = list(columns)
    if len(columns) == 1:
        columns = columns[0]

    try:
        data = df[columns]
    except KeyError as e:
        print(f'KeyError: key {e} not found')
        exit()
    else:
        return data

def df_split_fullname(column):
    """Возвращает n колонок, где n - количество слов через пробел"""
    return column.str.split(expand=True)

# TODO
def df_gender():
    """Возвращает колонку с полом"""
    conds = [
        df['отчество'].str.endswith('ВИЧ'),
        df['отчество'].str.endswith('ВНА'),
    ]
    choices = ['мужской', 'женский']
    return np.select(conds, choices, default='не определен')


def df_timedelta(start, end, format: str):
    """Возвращает массив разницы во времени в формате указанного времени"""
    time = np.timedelta64(1, format)
    return ((df_get(start) - df_get(end)) / time).astype(int)


def df_on_pension():
    """Возвращает колонку сотрудников пенсионного возраста"""
    conds = [
        df['пол'].eq('мужской') & df['возраст на дату приема'].ge(60),
        df['пол'].eq('женский') & df['возраст на дату приема'].ge(55),
    ]
    choices = ['да', 'да']
    return np.select(conds, choices, default='нет')

def df_count_employees(column: str, startwith: str) -> int:
    """Возвращает кол-во сотрудников, где значение начинается с startwith"""
    return np.count_nonzero(df_get(column).str.startswith(startwith))

# TODO
def df_max_fired() -> str:
    """Возвращает пол с максимальным кол-вом уволенных сотрудников"""
    new_df = df_get('дата увольнения', 'пол').dropna()
    genders, counts = np.unique(new_df['пол'], return_counts=True)
    return genders[counts == counts.max()][0]


def df_max_count_name(column: str) -> str:
    """Возвращает максимально встречающееся значение колонки"""
    uniq, counts = np.unique(df_get(column).dropna(), return_counts=True)
    return uniq[counts == counts.max()][0]

def df_unique_len(column: str) -> str:
    """Возвращает кол-во различных вариантов колонки"""
    uniq = np.unique(df_get(column).dropna())
    return len(uniq)

if __name__ == '__main__':
    df = pd.read_excel(DOC_FILEPATH, SHEET_NAME)
    df.columns = [x.lower() for x in df.columns]

    # разбор по фио
    df[['фамилия', 'имя', 'отчество']] = df_split_fullname(df_get('фамилия'))
    # нахождение пола
    df['пол'] = df_gender()
    # кол-во лет на момент устройства
    df['возраст на дату приема'] = df_timedelta('дата приема', 'дата рождения', 'Y')
    # на пенсии
    df['пенсионного возраста'] = df_on_pension()

    # кол-во сотрудников по описанию
    # print(df_count_employees('фамилия', 'М'))
    # какой пол больше увольняли
    # print(df_max_fired())


    # самая частая причина увольнения
    # print(df_max_count_name('причина увольнения'))
    # кол-во разных вариантов увольнения
    # print(df_unique_len('причина увольнения'))

    df.columns = [x.capitalize() for x in df.columns]
    df.to_excel('result.xlsx', sheet_name='NewSheet', index=False)
