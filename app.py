from flask import Flask, render_template, request
import pandas as pd

from accounting_analysis import classify_entries, compute_kpi, demo_dataframe

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST' and 'file' in request.files:
        uploaded = request.files['file']
        if uploaded and uploaded.filename:
            df = pd.read_excel(uploaded, engine='openpyxl')
        else:
            df = demo_dataframe()
    else:
        df = demo_dataframe()

    classified = classify_entries(df)
    kpi = compute_kpi(classified)
    table_html = classified.to_html(index=False)
    return render_template('index.html', kpi=kpi, table=table_html)

if __name__ == '__main__':
    app.run(debug=True)
