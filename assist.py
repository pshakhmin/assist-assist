#!/bin/python3

import os.path

import click
import pandas as pd
from openpyxl import load_workbook


def is_solved(problem_standing):
    return 1 if '+' in str(problem_standing) else 0


def work_sheet(sheet, standings, contest):
    contests = {}

    for col in range(5, sheet.max_column):
        col_name = sheet.cell(1, col).value
        if col_name is not None:
            contests[col_name] = col

    cur_contest = list(contests.keys())[-1]

    if contest == 'last':
        print('Contest is not specified, processing the last in the table')

    if cur_contest not in contests:
        print(list(contests))
        print(f'Contest {cur_contest} not found (possible options:{",".join(list(contests))}), aborting..')
        exit(-1)

    print(f'Processing {cur_contest} - {sheet.title}')

    for row in range(1, sheet.max_row):
        student_name = sheet.cell(row, 1).value
        if student_name is None:
            continue
        if student_name == 'Student':
            continue

        student_login = str(sheet.cell(row, 2).value)

        student_row = standings[
            (standings['user_name'] == str(student_name)) | (standings['login'] == str(student_login))]
        if student_row.shape[0] == 0:
            print(f'Student {student_name} is not found. Putting zeros...')
            for ind, problem in enumerate(student_row.iloc[:, 3:-1]):
                sheet.cell(row, contests[cur_contest] + ind).value = 0
            continue

        for ind, problem in enumerate(student_row.iloc[:, 3:-1]):
            sheet.cell(row, contests[cur_contest] + ind).value = is_solved(student_row[problem])


@click.command()
@click.option("-csv", "--standings_csv", prompt=True, required=True, type=click.Path())
@click.option("-xlsx", "--grades_xlsx", prompt=True, required=True)
@click.option("-c", "--contest", prompt=True, required=True, default='last')
def run(standings_csv, grades_xlsx, contest):
    if not os.path.exists(standings_csv):
        print(f'File not found: {standings_csv}')
        exit(-1)

    if not os.path.exists(grades_xlsx):
        print(f'File not found: {grades_xlsx}')
        exit(-1)

    df = pd.read_csv(standings_csv)

    print(f'Done reading csv: {df.shape[0]} participants.')
    print(f'Mean score: {df["Score"].mean().round(2)}/{df["Score"].max()}')

    workbook = load_workbook(grades_xlsx)
    for sheet in workbook.worksheets:
        if sheet.title == 'All':
            continue
        work_sheet(sheet, df, contest)

    workbook.save(grades_xlsx)


if __name__ == '__main__':
    run()
