import calendar
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
import ReadData
from io import BytesIO
import openpyxl

excel_path = ''
writer = pd.ExcelWriter('VerdelingKlanten.xlsx', engine='openpyxl', mode='a')


def convert_int_to_month(number: int) -> calendar.month:
    return calendar.month[number]


def return_monthly_avg(month_df: pd.DataFrame) -> float:
    amount_clicks = month_df['Aantal'].sum()
    amount_companies = len(month_df['Klant'])
    try:
        return round(amount_clicks / amount_companies, 2)
    except ZeroDivisionError:
        return 0


def return_annual_result(df: pd.DataFrame, year: str) -> list[float]:
    result: list[float] = [year]
    month_list = df['Datum'].unique()
    for month in month_list:
        month_df = df[df['Datum'] == month]
        result.append(return_monthly_avg(month_df))

    return result


def get_column_names(annual_df: pd.DataFrame) -> list[str]:
    columns: list[str] = ['Year']
    df = annual_df.copy(deep=True)
    df['Datum'] = df['Datum'].apply(lambda x: convert_int_to_month(x.month))
    month_list = df['Datum'].unique()
    columns.extend(month_list)
    
    return columns

def write_df_to_excel(df: pd.DataFrame, sheet_name: str, start_col: float) -> None:
    df.to_excel(writer, sheet_name=sheet_name, startcol=start_col)

class DataSetAllYears:
    df: pd.DataFrame
    years: list[datetime]

    def __init__(self) -> None:
        self.df = pd.read_excel(excel_path)
        self.set_years()
        self.set_df_to_needed_data()

    def set_years(self) -> None:
        for i in range(18, 21):
            self.years.append(datetime(int(f'20{i}'), 1, 1))

    def set_df_to_needed_data(self) -> None:
        # only the amount of clicks are necessary, corresponding id's are 28, 38, 50
        self.df = self.df[(self.df['Kosten component id'] == 28) |
                          (self.df['Kosten component id'] == 38) |
                          (self.df['Kosten component id'] == 50)]

    def get_df_previous_years(self) -> list[pd.DataFrame]:
        df_list: list[pd.DataFrame] = []
        for i in range(len(self.years) - 1):
            annual_df = self.df[self.df['Datum'] < self.years[i + 1]]
            annual_df = annual_df[annual_df['Datum'] > self.years[i]]
            df_list.append(annual_df)

        return df_list

    def get_df_current_year(self) -> pd.DataFrame:
        return self.df[self.df['Datum'] > self.years[-1]]

    def get_result_previous_years(self) -> list[list[float]]:
        result_list: list[list[float]] = []
        df_previous_years = self.get_df_previous_years()
        for index, year in enumerate(range(18, 21)):
            annual_result = return_annual_result(df=df_previous_years[index], year=f'20{year}')
            result_list.append(annual_result)

        return result_list

    def get_result_current_year(self) -> list[float]:
        df_current_year = self.get_df_current_year()
        
        return return_annual_result(df=df_current_year, year='2021')

    def get_overall_result(self) -> list[list[float]]:
        overall_result: list[list[float]] = []
        result_previous_years = self.get_result_previous_years()
        result_current_year = self.get_result_current_year()
        # adjust the length of the list 'result_current_year' to the length of an annual result list
        for _ in range(len(result_previous_years[-1]) - len(result_current_year)):
            result_current_year.append(0)

        overall_result.extend(result_previous_years)
        overall_result.append(result_current_year)
        
        return overall_result

    def get_result_df(self) -> pd.DataFrame:
        annual_df = self.get_df_previous_years()[-1]
        columns = get_column_names(annual_df=annual_df)
        result_df: pd.DataFrame = pd.DataFrame(columns=columns)
        overall_result = self.get_overall_result()
        for result in overall_result:
            result_df.loc[len(result_df)] = result

        return result_df

    def write_result_df_to_excel(self) -> None:
        result_df = self.get_result_df()
        result_df.set_index('Year')
        write_df_to_excel(df=result_df, sheet_name='Data_6', start_col=1)


class DataCurrentSituation:
    app: ReadData.ReadData

    def __init__(self) -> None:
        self.app = ReadData.ReadData(path_file=excel_path)
        self.app.process_customers()
        self.df = self.app.get_dataframe_from_customer_list()
        self.df = self.df[self.df['Abonnement'].notna()]

    def get_df_active_customers(self) -> pd.DataFrame:
        df = self.df.sort_values(by='Laatst gefactureerde datum')
        
        return df[df['Laatst gefactureerde datum'] == df['Laatst gefactureerde datum'].unique()[-1]]

    def write_active_df_to_excel(self) -> None:
        df = self.get_df_active_customers()
        df.to_excel(writer, sheet_name='Data_punt_3_HuidigeVerdeling', startcol=4)


