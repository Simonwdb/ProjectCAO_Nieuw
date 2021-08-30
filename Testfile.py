import pandas as pd
import matplotlib.pyplot as plt


def get_subscriptions_length(string: str, df: pd.DataFrame) -> list:
    values = [string]
    subscriptions = ['Light', 'Standard', 'Premium']
    for sub in subscriptions:
        values.append(len(df[df['Abonnement'] == sub]))
    return values


def create_result_df(first_df: pd.DataFrame, second_df: pd.DataFrame) -> pd.DataFrame:
    values_normal = get_subscriptions_length('Huidig', first_df)
    values_updated = get_subscriptions_length('Na verdeling', second_df)
    cols = ['Situatie', 'Light', 'Standard', 'Premium']
    result = pd.DataFrame(columns=cols)
    result.loc[len(result)] = values_normal
    result.loc[len(result)] = values_updated
    result = result.set_index('Situatie')
    return result


def create_plot(first_df: pd.DataFrame, second_df: pd.DataFrame, thresholds: list[float]):
    df = create_result_df(first_df, second_df)
    ax = df.T.plot(kind='bar', rot=360)
    for p in ax.patches:
        ax.annotate(str(p.get_height()), (p.get_x() + 0.05, p.get_height() + 0.75))
    plt.title(f'Met snijpunten:\n Light:{thresholds[0]} - Standard: {thresholds[1]}')
    plt.show()











