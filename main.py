import os, os.path
import win32com.client

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    if os.path.exists("D:\\excle_macro_temp\\prem_macro.xlsm"):
        xl=win32com.client.Dispatch("Excel.Application")
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(os.path.abspath("D:\\excle_macro_temp\\x7.xlsx"))
        try:
            xl.Application.Run("m4.xlsm!Module1.Macro_3")
            # NOT WORK BELOW!!!!
            # xl.Application.Run("D:\\excle_macro_temp\\m7.xlsm!Module1.Macro_3")
        except:
            xl.Application.Quit()  # Comment this out if your excel script closes
            del xl
            exit

        wb.SaveAs(Filename="D:\\haohao.xlsx")
        wb.Close()
        xl.Quit() # Comment this out if your excel script closes
        del xl