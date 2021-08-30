import matplotlib.pyplot as plt
from Customer import Customer
from Settings import settings
import pandas as pd


def update_subscriptions(data_df: pd.DataFrame, light_threshold: str, standard_threshold: str) -> pd.DataFrame:
    data_df = data_df.copy(deep=True)

    data_df.loc[(data_df['Gem. Gebruik'] < light_threshold), 'Abonnement'] = 'Light'

    data_df.loc[(data_df['Gem. Gebruik'] > light_threshold) &
                (data_df['Gem. Gebruik'] < standard_threshold), 'Abonnement'] = 'Standard'

    data_df.loc[(data_df['Gem. Gebruik'] > standard_threshold), 'Abonnement'] = 'Premium'

    return data_df


def create_graph(column_names: list, values: list, title: str) -> None:
    plt.figure(figsize=(9, 6))
    plt.bar(column_names, values)
    plt.title(title)
    for n, v in zip(column_names, values):
        plt.text(n, v + 0.75, str(v))
    plt.show()


class ReadData:
    def __init__(self, path_file: str) -> None:
        self.df = pd.read_excel(path_file, parse_dates=[11, 12])
        self.df = self.df[self.df['Debiteur nr.'].notna()]
        self.columns_data_df = ['Debiteur nr.', 'Klant', 'Abonnement', 'Totaal betaald', 'Gem. Prijs',
                                'Maanden boven bundel', 'Aantal maanden actief', 'Gem. Gebruik', 'Mediaan gebruik',
                                'Modus gebruik', 'Hoogste gebruik', 'Laatst gefactureerde datum', 'Subscription Object']
        self.customers: list[Customer] = []

    def process_customers(self) -> None:
        unique_debtor_list = self.df['Debiteur nr.'].unique()
        for u in unique_debtor_list:
            customer_df = self.df[self.df['Debiteur nr.'] == u]
            customer = Customer()
            customer.process_dataframe(customer_df)
            customer.perform_calculations()
            self.customers.append(customer)

    def get_dataframe_from_customer_list(self) -> pd.DataFrame:
        df = pd.DataFrame(columns=self.columns_data_df)
        for c in self.customers:
            values = [
                c.debtor_number, c.name, c.subscription_type, c.total_costs_paid, c.avg_month_price,
                c.months_above, len(c.subscriptions), c.avg_clicks, c.median_clicks,
                c.mode_clicks, c.highest_clicks, c.last_paid_month, c.subscriptions
            ]
            df.loc[len(df)] = values

        return df

    def update_customer_list(self, data_df: pd.DataFrame) -> pd.DataFrame:
        for index, row in data_df.iterrows():
            for subscription in row['Subscription Object']:
                subscription.change_single_subscription(row['Abonnement'])

        for c in self.customers:
            c.perform_calculations()

        return data_df


# if __name__ == '__main__':
#     path = settings.path_file
#     app = ReadData(path_file=path)
