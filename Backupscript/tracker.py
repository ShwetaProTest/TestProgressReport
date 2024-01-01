from flask import Flask, render_template, request
from datetime import datetime
import csv

app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        expense_date = request.form['date']
        expense_amount = request.form['amount']

        with open('../Reports/expenses.csv', 'a') as f:
            writer = csv.writer(f)
            writer.writerow([expense_date, expense_amount])

    return render_template('index.html')

@app.route('/monthly_average', methods=['GET'])
def monthly_average():
    monthly_expenses = {}

    with open('expenses.csv', 'r') as f:
        reader = csv.reader(f)
        next(reader)  # skip header row
        for row in reader:
            expense_date = datetime.strptime(row[0], '%Y-%m-%d')
            year_month = expense_date.strftime('%Y-%m')
            if year_month not in monthly_expenses:
                monthly_expenses[year_month] = []
            monthly_expenses[year_month].append(float(row[1]))

    monthly_averages = {}
    for year_month, expenses in monthly_expenses.items():
        monthly_averages[year_month] = sum(expenses) / len(expenses)

    return render_template('monthly_average.html', monthly_averages=monthly_averages)

if __name__ == '__main__':
    app.run(debug=True)
