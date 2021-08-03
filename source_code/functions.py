import datetime
import re


def get_rate(num, den, row, output_box):
    try:
        if num == -1:
            print(f'CHECK ROW #{row} in hsog.xlsx!!!!!')
            output_box.append(f'\n\nCHECK ROW #{row} in hsog.xlsx!!!!!\n\n')
        return round((num / den), 2)
    except Exception:
        return 0


def copy_col_same_row(rows, sheet, col1, col2):
    for x in range(1, rows):
        sheet[col2 % x] = sheet[col1 % x].value


# copy value from cell below to the cell on the right
def copy_col_one_up(rows, sheet, col1, col2):
    for x in range(1, rows):
        sheet[col2 % x] = sheet[col1 % (x + 1)].value


def move_columns(sheet, rows):
    for x in range(1, rows):
        sheet['U%s' % x] = sheet['A%s' % x].value
        sheet['W%s' % x] = sheet['A%s' % (x + 1)].value
        sheet['X%s' % x] = sheet['R%s' % x].value
        sheet['Y%s' % x] = sheet['R%s' % (x + 1)].value
        sheet['Z%s' % x] = sheet['R%s' % (x + 2)].value


def find_missing(sheet, col, row):
    cols = ['Q%s', 'P%s', 'O%s', 'N%s', 'M%s', 'L%s']
    val = sheet[col % row].value
    adjusted_row = row
    if col == 'Y%s':
        adjusted_row = row + 1

    if col == 'Z%s':
        adjusted_row = row + 2

    if val == '-1':
        x = 0
        while val is None or val == '-1':
            try:
                val = sheet[cols[x] % adjusted_row].value.split(' ')[-1]
                sheet[col % row] = val
                # print(f'row={row} col={cols[x]}  fixed value={val}')
            except Exception:
                pass
            if x > 4:
                break
            x += 1


def get_tos(sheet, row):
    cols = ['B%s', 'C%s', 'D%s', 'E%s', 'F%s', 'G%s', 'H%s', 'I%s', 'J%s', 'K%s', 'L%s', 'M%s', 'N%s', 'O%s', 'P%s',
            'Q%s']
    tos_count = 0

    for i in range(len(cols)):
        val = sheet[cols[i] % row].value
        if val is not None:
            sku_tos = str(val).strip().split(' ')
            if (len(sku_tos) > 1 and sku_tos[1] == '@') or (len(sku_tos) == 1 and sku_tos[0] == '@'):
                tos_count += 1
    return tos_count


def filter_for_skus(sheet, rows):
    cols = ['X%s', 'Y%s', 'Z%s']
    for x in range(1, rows):
        # copy sku to new column
        a = str(sheet['B%s' % (x + 1)].value)
        if a == None:
            sheet['T%s' % x] = '0'
        else:
            sheet['T%s' % x] = a[:6]

        # look for empty cells
        for col in cols:
            c = sheet[col % x].value
            if c == None or c == '':
                if col == 'X%s':
                    sheet[col % x] = '0'
                else:
                    sheet[col % x] = '-1'

        # filter for skus
        b = sheet['T%s' % x].value
        try:
            val = int(str(b).strip())
            if val > 99:
                sheet['T%s' % x] = str(val)
        except Exception:
            continue

        # move to new column same row SKU
        sheet['V%s' % x] = sheet['T%s' % x].value


# copy data from one column to another, moving the cells up two index positions
def copy_col_two_up(rows, sheet, col1, col2):
    for x in range(1, rows):
        sheet[col2 % x] = sheet[col1 % (x + 2)].value


# Function to retrieve and fill report date range
def get_report_date(sheet, col1):
    for cell in range(1, 5):
        c = str(sheet[col1 % cell].value)
        if 'From' in c:
            return datetime.datetime.strptime(c[-10:], "%m/%d/%Y").date()


def decInput():
    while True:
        user_input = input("Enter projected increase (e.g 10% = .1): ")
        try:
            return float(user_input)
        except ValueError:
            print("Not a valid entry")


def isFloat(s):
    search_return = re.search(r"\.", s)
    if not search_return:
        return False

    return s.replace('.', '', 1).isdigit()


def checkFloat(s):
    s = s.strip()
    search_return = re.search(r"\.", s)
    if not search_return:
        return False

    # handle the 6 for x.xx price string
    if s.count(' ') == 2:
        t = s.split(' ')
        if t[1] == 'for':
            t = t[2]
            return float(t[1:])

    # handle new item
    if s[0] == 'N' and s[3] == '$':
        return float(s[3:])

    if s.count('.') == 1:
        if s[0] == '@':
            s = s[2:]
        if s[0] == '$':
            s = s[1:]

        if s == '0.00':
            return float('0.01')
        return float(s)


def checkSize(s):
    s = s.strip()
    if s.count('/') == 2:
        s = s.split('/')
        if s[2] == '2021':
            return None
        # print(f'string={s} int={int(s[0])}')
        return int(s[0])
    return None


def getPrices(row, sheet, col):
    cell = f'{col}%s' % row
    val = sheet[cell].value

    if val is not None:

        try:
            a = checkFloat(str(val))
            # a = float(val)
            return a
        except Exception:
            return False

    return False


def getSkus(row, sheet, col):
    cell = f'{col}%s' % row
    val = sheet[cell].value
    val = str(val).strip()

    if val is not None:
        try:
            return int(val)
        except Exception:
            return None

    return None


def getSize(row, sheet, col):
    cell = f'{col}%s' % row
    val = sheet[cell].value

    if val is not None:
        try:
            return checkSize(str(val))

        except Exception:
            return False

    return False


def output_msg(msg, val, output_box, progress_bar):
    print(msg)
    output_box.append(msg)
    progress_bar.setValue(val)
