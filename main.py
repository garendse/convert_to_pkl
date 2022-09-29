import PySimpleGUI as sg
import pandas as pd


def _get_excel_filenames():
    layout = [
        [sg.Text("Asset Data File: "), sg.Input(), sg.FileBrowse(key="-XLS-", file_types=(("XLSX Files", "*.xlsx"),))],
        [sg.Text("Worksheet: "), sg.Input(key='-WSN-', size=(6, 1), default_text='FAR')],
        [sg.Text("Pickled Name: "), sg.Input(key='-PKL-', size=(40, 1), default_text='c:\\tmp\\output.pkl')],
        [sg.Button("Submit"), sg.Cancel()]]

    window = sg.Window("Pickle Data Converter", layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Exit":
            break
        if event == "Submit":
            return (values["-XLS-"], values["-WSN-"], values["-PKL-"])


if __name__ == '__main__':
    xls_name, ws_name, pkl_name = _get_excel_filenames()

    if len(xls_name) > 0:
        try:
            print("Reading", xls_name)
            df = pd.read_excel(xls_name, ws_name)

            print("Saving", pkl_name)
            df.to_pickle(pkl_name)

            print("Done")

        except Exception as e:
            print("Error:", e)
