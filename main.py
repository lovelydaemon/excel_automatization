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


def df_split(column: str):
    """Возвращает n колонок, где n - количество слов через пробел"""
    return df_get(column).str.split(expand=True)


def df_gender():
    """Возвращает колонку с полом"""
    conds = [
        df_get('отчество').str.endswith('ВИЧ'),
        df_get('отчество').str.endswith('ВНА'),
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
        df_get('пол').eq('мужской') & df_get('возраст на дату приема').ge(60),
        df_get('пол').eq('женский') & df_get('возраст на дату приема').ge(55),
    ]
    choices = ['да', 'да']
    return np.select(conds, choices, default='нет')


def df_count_employees(column: str, startwith: str) -> int:
    """Возвращает кол-во сотрудников, где значение начинается с startwith"""
    return df_get(column).str.startswith(startwith).sum()


def df_max_fired() -> str:
    """Возвращает пол с максимальным кол-вом уволенных сотрудников"""
    new_df = df_get('дата увольнения', 'пол').dropna()
    genders, c = np.unique(new_df['пол'], return_counts=True)
    return genders[c.argmax()]


def df_max_count_value(column: str) -> str:
    """Возвращает максимально встречающееся значение колонки"""
    uniq, c = np.unique(df_get(column).dropna(), return_counts=True)
    return uniq[c.argmax()]


def df_num_working(at_time: str) -> int:
    new_df = df_get('дата увольнения')
    return np.logical_or(new_df.isna(), new_df.gt(at_time)).sum()


def df_num_fired(at_time: str) -> int:
    return np.sum(df_get('дата увольнения').le(at_time))


def df_mean_age_fired():
    new = df_get('дата рождения', 'дата увольнения').dropna()
    time = np.timedelta64(1, 'Y')
    return (np.subtract(new['дата увольнения'], new['дата рождения']) / time).mean().astype(int)


if __name__ == '__main__':
    df = pd.read_excel(DOC_FILEPATH, SHEET_NAME)
    df.columns = [x.lower() for x in df.columns]

    # разбор по фио
    df[['фамилия', 'имя', 'отчество']] = df_split('фамилия')
    # нахождение пола
    df['пол'] = df_gender()
    # кол-во лет на момент устройства
    df['возраст на дату приема'] = df_timedelta('дата приема', 'дата рождения', 'Y')
    # на пенсии
    df['пенсионного возраста'] = df_on_pension()

    # кол-во работающих на 1.1.17
    # print(df_num_working('2017-01-01'))

    # кол-во уволенных на 1.1.17
    # print(df_num_fired('2017-01-01'))

    # кол-во уволенных на 1.1.09
    # print(df_num_fired('2009-01-01'))

    # средний возраст уволенных
    # print(df_mean_age_fired())

    # какой пол больше уволили
    # print(df_max_fired())

    # кол-во сотрудников по описанию
    # print(df_count_employees('фамилия', 'М'))

    # кол-во разных вариантов увольнения
    # print(df_get('причина увольнения').nunique())

    # самая частая причина увольнения
    # print(df_max_count_value('причина увольнения'))

    df.columns = [x.capitalize() for x in df.columns]
    df.to_excel('result.xlsx', sheet_name='NewSheet', index=False)
