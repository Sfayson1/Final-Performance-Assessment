import csv
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
import matplotlib.pyplot as plt

# File paths
csv_file = "final.csv"
excel_file = "final.xlsx"


student_id = "SHFAY9008"


def askUser():
    total = 0

    # This loop runs 5 times to ask the user for a number each time.
    # Each number is added to the running total so the program can display
    # the sum after all 5 numbers have been entered.
    for _ in range(5):
        while True:
            try:
                number = int(input("Please enter a number: "))
                break
            except ValueError:
                print("Invalid input. Please enter a whole number.")
        total += number

    print("The sum for the 5 numbers entered is:", total)


def askIncome():
    # This loop runs 5 times to collect 5 names and 5 annual incomes.
    # Each set of values is immediately appended to the existing CSV file
    # so the original data remains and the new data is added to the end.
    with open(csv_file, "a", newline="") as file:
        writer = csv.writer(file)
        for _ in range(5):
            name = input("Please enter a name: ")
            while True:
                try:
                    income = float(input("Please enter their income: "))
                    break
                except ValueError:
                    print("Invalid input. Please enter a number.")
            writer.writerow([name, income])


def excelPie():
    data_rows = []

    with open(csv_file, "r", newline="") as file:
        reader = csv.reader(file)
        for row in reader:
            if len(row) == 2:
                name = row[0]
                income = int(float(row[1]))
                data_rows.append([name, income])

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet"

    for row in data_rows:
        ws.append(row)

    chart_title = f"{student_id} {datetime.today().strftime('%B %d, %Y')}"

    # Create a PieChart object for the worksheet.
    pie = PieChart()

    # Select the income values from column B to use as chart data.
    data = Reference(ws, min_col=2, min_row=1, max_row=len(data_rows))

    # Select the names from column A to use as category labels.
    labels = Reference(ws, min_col=1, min_row=1, max_row=len(data_rows))

    # Add the income data to the pie chart.
    pie.add_data(data)

    # Add the names as labels for each slice of the pie chart.
    pie.set_categories(labels)

    # Set the title of the pie chart to StudentID and today's date.
    pie.title = chart_title

    # Position the pie chart on the worksheet starting at cell D10.
    ws.add_chart(pie, "D10")

    # Save the workbook as final.xlsx in the FinalExam folder.
    wb.save(excel_file)


def verticalBar():
    names = []
    incomes = []

    with open(csv_file, "r", newline="") as file:
        reader = csv.reader(file)
        for row in reader:
            if len(row) == 2:
                names.append(row[0])
                incomes.append(int(float(row[1])))

    chart_title = f"{student_id} {datetime.today().strftime('%B %d, %Y')}"

    plt.bar(names, incomes, color="green", label="Income")
    plt.title(chart_title)
    plt.xlabel("Name")
    plt.ylabel("Income")
    plt.legend()
    plt.show()


def main():
    askUser()
    askIncome()
    excelPie()
    verticalBar()


if __name__ == "__main__":
    main()
