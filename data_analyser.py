import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook

NUMBER_OF_RESPONSES = 148
NUMBER_OF_ROWS = NUMBER_OF_RESPONSES + 1

wb = load_workbook(filename='./../results_iterasjon_1.xlsx')

sheet = wb.active

def get_column_data(column='A'):
    """
        get all data in the specified column, returns list of data
    """
    title = sheet[f'{column}1'].value
    cells = sheet[f'{column}2':f'{column}{NUMBER_OF_RESPONSES}']
    data = [int(cell[0].value) for cell in cells]

    return data, title


def plot_hours_spent(*columns):
    """ Plot a bar diagram of answers related to hours spent on subject """
    for column in columns:
        values, title = get_column_data(column)

        bar_values = [values.count(i) for i in range(0,21)]

        bars = np.arange(0, 21, 1)
        
        mean = round(sum(values) / len(values), 3)
        plt.title(f'{title}\nGjennomsnitt: {str(mean)}')

        plt.bar(bars, bar_values, align='center')

        plt.xlabel('Antall timer')
        plt.ylabel('Antall svar')
        plt.grid(axis='y')

        plt.savefig(f'./../diagrams/hours/{title.replace(" ", "_").lower()}_bar.png')
        plt.close()

def plot_satisfaction(*columns):
    """ Questions related to scale of 1=Ikke tilfreds to 5=Svært tilfreds """
    for column in columns:
        values, title = get_column_data(column)

        bar_values = [values.count(1), values.count(2), values.count(3), values.count(4), values.count(5)]

        bars = np.arange(1, 5 + 1)
        
        mean = round(sum(values) / len(values), 3)
        plt.title(f'{title}\nGjennomsnitt: {str(mean)}')

        plt.bar(bars, bar_values, align='center')

        plt.xlabel('1=Ikke tilfreds | 5=Svært tilfreds')
        plt.ylabel('Antall svar')
        plt.grid(axis='y')

        plt.savefig(f'./../diagrams/satisfaction/{title.replace(" ", "_").lower()}_bar.png')
        plt.close()


def plot_agreement(*columns):
    """ Questions related to scale of 1=Helt uenig to 5=Helt enig """
    for column in columns:
        values, title = get_column_data(column)

        bar_values = [values.count(i) for i in range(1,5+1)]

        bars = np.arange(1, 5 + 1)
        
        mean = round(sum(values) / len(values), 3)
        plt.title(f'{title}\nGjennomsnitt: {str(mean)}')

        plt.bar(bars, bar_values, align='center')

        plt.xlabel('1=Helt uenig | 5=Helt enig')
        plt.ylabel('Antall svar')
        plt.grid(axis='y')

        plt.savefig(f'./../diagrams/agreement/{title.replace(" ", "_").lower()}_bar.png')
        plt.close()



if __name__ == '__main__':
     
    # Questions about Hours spent
    columns_for_hours = ['D', 'E']
    # plot_hours_spent(*columns_for_hours)
    plot_hours_spent(*columns_for_hours)
    
    # Questions about Satisfaction
    columns_for_satisfaction = ['F', 'G']
    # plot_satisfaction(*columns_for_satisfaction)
    plot_satisfaction(*columns_for_satisfaction)

    # Questions about Agreement
    columns_for_agreement = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']
    # plot_agreement(*columns_for_agreement)
    plot_agreement(*columns_for_agreement)