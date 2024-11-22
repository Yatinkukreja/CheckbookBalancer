import pdfplumber
from collections import defaultdict
import pandas as pd


def learn_boundaries(line):
    starts = []
    ends = []
    lastX = 0
    for c in line:
        if c['x0'] > lastX + 5:
            starts.append(c['x0'])
            ends.append(lastX)
        lastX = int(c['x1'])
    ends.append(lastX)
    return [
        starts[0], ends[1] + 10, (ends[2] + starts[2]) / 2, starts[3] - 25, ends[4], ends[5], ends[6]
    ]


def read_line(line):
    builder = ""
    for c in line:
        builder += c['text']
    return builder


def organize_line(line, boundaries):
    assert boundaries
    assert line
    row = [""]
    i = 1
    for c in line:
        while i < len(boundaries) and c['x0'] >= boundaries[i]:
            i += 1
            row.append("")
        row[-1] += c['text']
    return row


def organize_text(lines):
    table = []
    collectBoundaries = False
    tableStarted = False
    boundaries = None
    for line in lines:
        lineText = read_line(line)
        if tableStarted and lineText.strip().split(" ")[0] == "Ending":
            return table
        if collectBoundaries:
            boundaries = learn_boundaries(line)
            collectBoundaries = False
            tableStarted = True
            continue
        elif tableStarted:
            new_row = organize_line(line, boundaries)
            if table and new_row[0] == "":
                table[-1][2] += new_row[2]
            elif new_row[0] != "":
                table.append(new_row)
        if lineText.strip().split(" ")[0] == "Transaction":
            collectBoundaries = True
    return table


def group_and_sort_chars(chars, y_tolerance):
    lines = defaultdict(list)
    for char in chars:
        matchedLine = None
        for y, line_chars in lines.items():
            if abs(char['y0'] - y) < y_tolerance:
                matchedLine = y
                break
        if matchedLine:
            lines[matchedLine].append(char)
        else:
            lines[char['y0']].append(char)
    return [
        sorted(line_chars, key=lambda c: c['x0'])
        for line_chars in lines.values()
    ]


def parse_page(file, page_number, table):
    chars = file.pages[page_number].chars
    sorted_lines = group_and_sort_chars(chars, 2)
    extra = organize_text(sorted_lines)
    table += extra


def format_table(table):
    balance = None
    first_balance_instance = -1
    for rowCount, row in enumerate(table):
        row.pop()
        for i in range(len(row)):
            row[i] = row[i].strip()
        temp = row[0]
        row[0] = row[1]
        row[1] = temp
        if not balance:
            if row[5] != "":
                balance = round(string_to_float(row[5]), 2)
                first_balance_instance = rowCount
            continue
        if row[3]:
            balance += string_to_float(row[3])
        if row[4]:
            balance -= string_to_float(row[4])
        balance = round(balance, 2)
        if row[5] != "":
            assert balance == string_to_float(row[5])
        else:
            row[5] = balance
    balance_backwards(table, first_balance_instance)


def balance_backwards(table, start):
    if start < 1:
        return
    balance = string_to_float(table[start][5])
    for i in range(start - 1, -1, -1):
        if table[i + 1][3]:
            balance -= string_to_float(table[i + 1][3])
        if table[i + 1][4]:
            balance += string_to_float(table[i + 1][4])
        balance = round(balance, 2)
        table[i][5] = balance


def string_to_float(number):
    builder = ""
    for n in number:
        if n != ",":
            builder += n
    return float(builder)


def write_to_excel(file_name, table):
    df = pd.DataFrame(table, columns=["Check No.", "Date", "Description", "Credit", "Debit", "Balance"])
    df.to_excel(file_name, index=False, engine='openpyxl')


if __name__ == '__main__':
    # replace with absolute filepath of pdf, using '//' instead of '/'
    file = "C:\\Users\\Yatin\\Downloads\\073124 WellsFargo.pdf"
    # replace with name you would like for the excel sheet
    excel_path = "transactions.xlsx"
    with pdfplumber.open(file) as pdf:
        table = []
        for i in range(1, len(pdf.pages)):
            parse_page(pdf, i, table)
    format_table(table)
    write_to_excel(excel_path, table)
