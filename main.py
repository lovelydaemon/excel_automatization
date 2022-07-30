import pandas as pd


DOC_FILEPATH = 'sample.xlsx'
SHEET_NAME = 'Лист2'

df = pd.read_excel(DOC_FILEPATH, SHEET_NAME)
df.columns = [x.lower() for x in df.columns]

def pd_get_data(*columns: str):
    columns = list(columns)
    if len(columns) == 1:
        columns = columns[0]

    data = df.get(columns)
    if data is None:
        raise ValueError('columns not found')
    return data



df[['фамилия', 'имя', 'отчество']] = df['фамилия'].str.split(expand=True)


df.to_excel('result.xlsx', sheet_name='NewSheet', index=False)
