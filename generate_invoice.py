import numpy as np
import pandas as pd
import os
from datetime import datetime as dt
import openpyxl
import click

@click.command()
@click.argument('invoice_template', type=click.Path(exists=True))
@click.argument('time_file', type=click.Path(exists=True))
@click.argument('invoice_id', type=int)
@click.option('-o', '--out_path', type=click.Path(), default=None)
def generate_invoice(invoice_template: str, time_file: str, invoice_id: int, out_path: str = None):
    fname = f'Invoice_{invoice_id}.xlsx'
    if out_path is None:
        out_path = fname
    elif out_path == ":template":
        out_path = os.path.dirname(invoice_template)
        out_path = os.path.join(out_path, fname)
    elif os.path.isdir(out_path):
        out_path = os.path.join(out_path, fname)

    df = pd.read_csv(time_file)[["Description", "Start date", "Duration", "Project"]].rename(columns={'Start date': 'Date'})
    df["Duration"] = pd.to_timedelta(df["Duration"])
    # df = df.groupby(["Description", "Date"]).sum().reset_index()
    df = df.groupby(["Date", "Project"]).apply(
        lambda x: pd.Series({'Description': '; '.join(sorted(x.Description.unique())), 'Duration': x.Duration.sum()})
    ).reset_index()
    df['Duration'] = np.round(df.Duration.dt.seconds / 60 / 60 * 10) / 10
    df['Description'] = df['Project'] + ': ' + df['Description']
    del df['Project']

    time_report = df.sort_values("Date")
    workbook = openpyxl.load_workbook(invoice_template)
    sheet = workbook['Invoice Template']

    sheet['B3'] = f'#{int(invoice_id):03d}'
    sheet['E10'] = dt.now().strftime('%m/%d/%Y')

    for i, row in time_report.iterrows():
        ri = i + 16
        sheet['C' + str(ri)] = row['Description']
        sheet['D' + str(ri)] = row['Date']
        sheet['E' + str(ri)] = row['Duration']

    workbook.save(out_path)


if __name__ == '__main__':
    generate_invoice()
