import statistics
import pandas as pd
from Subscription import Subscription
from Settings import settings


def get_clicks(data_df: pd.DataFrame, statement: str) -> int:
    code = settings.code_id['Requests within subscription']
    amount = 0
    for c in code:
        amount += data_df[(data_df['Kosten component id'] == c) & (data_df['Is extra'] == statement)]['Aantal'].sum()
    return amount


def process_monthly_df(data_df: pd.DataFrame) -> Subscription:
    date = data_df.iloc[0]['Factuur Export Datum'].date()

    free_requests = get_clicks(data_df, 'Nee')
    extra_requests = get_clicks(data_df, 'Ja')

    try:
        fee = data_df[data_df['Kosten component id'] == settings.code_id['Subscription form']].iloc[0]['Totaal, €']
    except IndexError:
        fee = 0

    sub = Subscription(price=fee, free_clicks=free_requests, extra_clicks=extra_requests, date=date)

    return sub


class Customer:

    def __init__(self) -> None:
        self.subscriptions: list[Subscription] = []
        self.debtor_number: int = 0
        self.name: str = ''
        self.subscription_type: str = ''
        self.month_highest_above: int = 0
        self.last_paid_month: str = ''
        self.total_costs_paid = 0
        self.months_above: int = 0
        self.avg_month_price: float = 0
        self.avg_clicks: float = 0
        self.median_clicks: int = 0
        self.mode_clicks: int = 0
        self.highest_clicks: int = 0

    def process_dataframe(self, df: pd.DataFrame) -> None:
        self.debtor_number = df.iloc[0]['Debiteur nr.']
        self.name = df.iloc[0]['Klant']
        df = df.sort_values(by='Factuur Export Datum')
        unique_date_list = df['Factuur Export Datum'].unique()
        for u in unique_date_list:
            month_df = df[df['Factuur Export Datum'] == u]
            subscription = process_monthly_df(month_df)
            self.subscriptions.append(subscription)

    def get_subscription_type(self):
        result = []
        for s in self.subscriptions:
            try:
                result.append(s.subscription_type)
            except AttributeError:
                pass
        try:
            return max(result) if len(result) > 0 else None
        except TypeError:
            return None

    def count_total_costs(self) -> int:
        result = 0
        for s in self.subscriptions:
            result += s.price
            try:
                if s.total_clicks > s.allowed_clicks:
                    result += (s.total_clicks - s.allowed_clicks) * s.extra_fee
            except AttributeError:
                pass
        return result

    def count_months_above(self) -> int:
        result = 0
        for s in self.subscriptions:
            try:
                if s.total_clicks > s.allowed_clicks:
                    result += 1
            except AttributeError:
                pass
        return result

    def count_avg_month_price(self) -> float:
        total_costs = self.count_total_costs()
        try:
            result = round(total_costs / len(self.subscriptions), 2)
        except ZeroDivisionError:
            result = 0
        return result

    def count_avg_clicks(self) -> float:
        try:
            result = round(sum([s.total_clicks for s in self.subscriptions]) / len(self.subscriptions), 2)
        except ZeroDivisionError:
            result = 0
        return result

    def count_median_clicks(self) -> int:
        try:
            result = statistics.median([s.total_clicks for s in self.subscriptions])
        except statistics.StatisticsError:
            result = 0
        return result

    def count_mode_clicks(self) -> int:
        sub_list = [s.total_clicks for s in self.subscriptions if s.total_clicks > 0]
        try:
            result = max(set(sub_list), key=sub_list.count)
        except ValueError:
            result = 0
        return result

    def count_highest_clicks(self) -> int:
        result = 0
        for s in self.subscriptions:
            if s.total_clicks >= result:
                result = s.total_clicks
        return result

    def set_last_paid_month(self) -> None:
        self.last_paid_month = self.subscriptions[-1].date

    def perform_calculations(self) -> None:
        self.subscription_type = self.get_subscription_type()
        self.total_costs_paid = self.count_total_costs()
        self.months_above = self.count_months_above()
        self.avg_month_price = self.count_avg_month_price()
        self.avg_clicks = self.count_avg_clicks()
        self.median_clicks = self.count_median_clicks()
        self.mode_clicks = self.count_mode_clicks()
        self.highest_clicks = self.count_highest_clicks()
        self.set_last_paid_month()

    def change_subscriptions(self, data_dict: dict) -> None:
        for s in self.subscriptions:
            s.change_single_subscription_with_parameters(data_dict)

    def find_and_change_subscription_with_parameters(self, data: list) -> None:
        data_dict = {
            'Type': str,
            'Price': float,
            'Allowed clicks': int,
            'Extra fee': float
        }
        for d in data:
            subscription_dict = dict(zip(data_dict.keys(), d))
            if subscription_dict['Type'] == self.subscription_type:
                self.change_subscriptions(subscription_dict)

