import os, os.path
import win32com.client
import sys

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    if os.path.exists("D:\\excle_macro_temp\\prem_macro.xlsm"):
        xl=win32com.client.Dispatch("Excel.Application")
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(os.path.abspath(sys.argv[1]))
        excel_name = os.path.basename(sys.argv[1])
        try:
            # Note: Macro_1 should not be run since it was fixed to file path of Andrea
            #       real macro from 2 to 7
            # xl.Application.Run("m4.xlsm!Module1.Macro_1")

            xl.Application.Run("m4.xlsm!Module1.Macro_2")
            xl.Application.Run("m4.xlsm!Module1.Macro_3")
            xl.Application.Run("m4.xlsm!Module1.Macro_4")
            xl.Application.Run("m4.xlsm!Module1.Macro_5")
            xl.Application.Run("m4.xlsm!Module1.Macro_6")
            xl.Application.Run("m4.xlsm!Module1.Macro_7")

            # NOT WORK BELOW!!!!
            # xl.Application.Run("D:\\excle_macro_temp\\m4.xlsm!Module1.Macro_2")
            # xl.Application.Run("D:\\excle_macro_temp\\m4.xlsm!Module1.Macro_3")
            # xl.Application.Run("D:\\excle_macro_temp\\m4.xlsm!Module1.Macro_4")
            # xl.Application.Run("D:\\excle_macro_temp\\m4.xlsm!Module1.Macro_5")
            # xl.Application.Run("D:\\excle_macro_temp\\m4.xlsm!Module1.Macro_6")
            # xl.Application.Run("D:\\excle_macro_temp\\m4.xlsm!Module1.Macro_7")
        except:
            xl.Application.Quit()  # Comment this out if your excel script closes
            del xl
            exit

        wb.SaveAs(Filename="D:\\excle_macro_temp\\" + "macro_" + excel_name)
        # wb.SaveAs(Filename="D:\\excle_macro_temp\\ciao.xlsx")

        wb.Close()
        xl.Quit() # Comment this out if your excel script closes
        del xl