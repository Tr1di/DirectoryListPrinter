import tempfile
import glob
import os
import sys
import xlsxwriter

TEXT_FORMAT = {
    'font_color': 'black',
    'underline': 0,
    'font_size': 12,
    'border': 1,
    'border_color': '#000000'
}

DIRECTORY = sys.argv[1]
SAVE_DIRECTORY = os.path.expanduser("~/Documents") + "/Оглавления/"

PRINT_LIST = False
DELETE_LIST = False


def is_a_slash(char):
    return char == "\\" or char == "/"


def get_files(files_dir):
    files = {}

    for file in glob.glob(files_dir + "*.*"):
        filename = file.replace(files_dir, "")
        files[filename] = file

    return files


def generate_list(files_dir, save_dir):
    book_name = save_dir + files_dir.split("\\")[-2] + ".xlsx"

    if DELETE_LIST:
        book_name = tempfile.mktemp(".xlsx")

    with xlsxwriter.Workbook(book_name) as book:
        sheet = book.add_worksheet("Sheet 1")

        t_format = book.add_format(TEXT_FORMAT)

        sheet.write(0, 0, "№", t_format)
        sheet.write(0, 1, "Название", t_format)

        files = get_files(files_dir)
        i = 1

        for file in files.keys():
            sheet.write(i, 0, i, t_format)
            sheet.write_url(i, 1, files[file], t_format, file)
            i += 1

        sheet.set_column("A:A", 5)
        sheet.set_column("B:B", 83.29)

    if PRINT_LIST:
        os.startfile(book_name, 'print')
    else:
        os.startfile(book_name)


if __name__ == "__main__":
    if not is_a_slash(DIRECTORY[-1]):
        DIRECTORY += "\\"

    if not os.path.exists(SAVE_DIRECTORY):
        os.makedirs(SAVE_DIRECTORY)

    if len(sys.argv) > 2:
        PRINT_LIST = bool(sys.argv[2].lower())

    if len(sys.argv) > 3:
        DELETE_LIST = bool(sys.argv[3].lower())

    if len(sys.argv) > 4:
        SAVE_DIRECTORY = sys.argv[4]
        if not is_a_slash(SAVE_DIRECTORY[-1]):
            SAVE_DIRECTORY += "\\"

    generate_list(DIRECTORY, SAVE_DIRECTORY)
