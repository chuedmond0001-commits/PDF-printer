import os
import sys
import time
import win32print
import win32api
import win32com.client

def start_printing_task():
    # 1. 取得資料夾路徑
    raw_path = input("請輸入或拖入資料夾路徑: ").strip().replace('"', '')
    folder_path = os.path.abspath(raw_path)

    if not os.path.exists(folder_path):
        print(f"❌ 錯誤：找不到路徑 -> {folder_path}")
        return

    # 2. 選擇印表機
    try:
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        printer_list = [p[2] for p in printers]
    except Exception as e:
        print(f"❌ 無法取得印表機清單: {e}")
        return

    print("\n--- 🖨️ 可用的印表機清單 ---")
    for i, name in enumerate(printer_list):
        print(f"[{i}] {name}")

    try:
        printer_choice = int(input("\n請輸入數字選擇印表機: "))
        selected_printer = printer_list[printer_choice]
        print(f"✅ 已選擇印表機: {selected_printer}")
    except (ValueError, IndexError):
        print("❌ 印表機選擇無效。")
        return

    # 3. 選擇列印模式 (新增多樣化選項)
    print("\n--- 📂 請選擇列印模式 ---")
    print("[1] 列印所有支援格式 (Word, Excel, RTF, PDF)")
    print("[2] 僅列印 PDF 檔案")
    print("[3] 僅列印 Word 檔案 (.docx, .doc)")
    print("[4] 僅列印 Excel 檔案 (.xlsx, .xls)")
    print("[5] 僅列印 RTF 檔案")
    
    mode_choice = input("\n請輸入數字選擇模式 (預設為 1): ").strip()
    
    if mode_choice == "2":
        valid_exts = ('.pdf',)
    elif mode_choice == "3":
        valid_exts = ('.docx', '.doc')
    elif mode_choice == "4":
        valid_exts = ('.xlsx', '.xls')
    elif mode_choice == "5":
        valid_exts = ('.rtf',)
    else:
        valid_exts = ('.docx', '.doc', '.rtf', '.xlsx', '.xls', '.pdf')

    # 4. 掃描檔案
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(valid_exts)]

    if not files:
        print(f"ℹ️ 在該資料夾中沒有發現符合條件的檔案。")
        return

    # 5. 排除檔案功能
    print("\n--- 📋 待列印檔案清單 ---")
    for i, f in enumerate(files):
        print(f"[{i}] {f}")
    
    exclude_input = input("\n請輸入想【排除】的檔案編號 (例如: 1,3,5)，若不排除請直接按 Enter: ").strip()
    
    if exclude_input:
        try:
            exclude_indices = {int(x.strip()) for x in exclude_input.split(',') if x.strip().isdigit()}
            files = [f for i, f in enumerate(files) if i not in exclude_indices]
            print(f"✅ 已排除指定檔案，剩餘 {len(files)} 個檔案。")
        except ValueError:
            print("⚠️ 輸入格式有誤，將處理清單中所有檔案。")

    if not files:
        print("ℹ️ 排除後已無剩餘檔案。")
        return

    # 6. 設定印表機模式
    is_pdf_mode = "PDF" in selected_printer.upper()
    if not is_pdf_mode:
        win32print.SetDefaultPrinter(selected_printer)

    word_app = None
    excel_app = None

    print(f"\n🚀 開始執行任務...")

    try:
        for file_name in files:
            full_path = os.path.join(folder_path, file_name)
            ext = os.path.splitext(file_name)[1].lower()
            output_pdf = os.path.splitext(full_path)[0] + ".pdf"
            
            print(f"正在處理: {file_name}")

            try:
                # Word 處理 (.docx, .doc, .rtf)
                if ext in ('.docx', '.doc', '.rtf'):
                    if not word_app:
                        word_app = win32com.client.Dispatch("Word.Application")
                        word_app.Visible = False
                    
                    doc = word_app.Documents.Open(full_path)
                    if is_pdf_mode:
                        doc.ExportAsFixedFormat(output_pdf, 17)
                        print(f"  ∟ [PDF] 已自動存檔/取代")
                    else:
                        word_app.ActivePrinter = selected_printer
                        doc.PrintOut()
                        print(f"  ∟ [OK] 已送出")
                    doc.Close(False)

                # Excel 處理 (.xlsx, .xls)
                elif ext in ('.xlsx', '.xls'):
                    if not excel_app:
                        excel_app = win32com.client.Dispatch("Excel.Application")
                        excel_app.Visible = False
                        excel_app.DisplayAlerts = False
                    
                    wb = excel_app.Workbooks.Open(full_path)
                    if is_pdf_mode:
                        wb.ExportAsFixedFormat(0, output_pdf)
                        print(f"  ∟ [PDF] 已自動存檔/取代")
                    else:
                        wb.PrintOut()
                        print(f"  ∟ [OK] 已送出")
                    wb.Close(False)

                # PDF 處理
                elif ext == '.pdf':
                    if is_pdf_mode:
                        print(f"  ∟ [SKIP] 本身即是 PDF")
                    else:
                        win32api.ShellExecute(0, "printto", full_path, f'"{selected_printer}"', ".", 0)
                        print(f"  ∟ [OK] 已送出")
                        time.sleep(1.5)

            except Exception as e:
                print(f"  ∟ [FAILED] 錯誤: {e}")

    finally:
        if word_app: word_app.Quit()
        if excel_app: excel_app.Quit()

    print("\n✨ 任務圓滿完成！")

if __name__ == "__main__":
    start_printing_task()