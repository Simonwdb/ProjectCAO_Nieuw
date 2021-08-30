import matplotlib.pyplot as plt
from CAOWijzer2.ReadData import ReadData


def create_plot(column_names: list, values: list, title: str) -> None:
    plt.figure(figsize=(9, 6))
    plt.bar(column_names, values)
    plt.title(title)
    for n, v in zip(column_names, values):
        plt.text(n, v + 0.75, str(v))
    plt.show()


path_file = 'Kopie van Facturen 14-06-2021 10_39_06.xlsx'

app = ReadData(path_file=path_file)
df = app.get_dataframe_from_customer_list()
print(df)
