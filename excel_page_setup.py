import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox

def adjust_page_setup(file_path):
    print(f"Processing file: {file_path}")
    try:
        # Excel 애플리케이션 객체 생성
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Excel 창을 보이지 않게 설정

        # Excel 파일 열기
        workbook = excel.Workbooks.Open(file_path)
        print("Excel file opened successfully.")

        # 모든 워크시트에 대해 페이지 설정 조정
        for sheet in workbook.Sheets:
            print(f"Processing sheet: {sheet.Name}")
            page_setup = sheet.PageSetup

            # 페이지 설정 코드 (여백, 용지 방향 등)

            page_setup.LeftMargin = excel.Application.InchesToPoints(0.5)
            page_setup.RightMargin = excel.Application.InchesToPoints(0.5)
            page_setup.TopMargin = excel.Application.InchesToPoints(0.5)
            page_setup.BottomMargin = excel.Application.InchesToPoints(0.5)

            # 용지 방향 설정 (1: 세로, 2: 가로)
            page_setup.Orientation = 2  # 세로 방향

            # 용지 크기 설정 (9: A4 용지)
            page_setup.PaperSize = 9  # A4 용지

            # 페이지에 맞추기 설정
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1  # 너비를 한 페이지로 맞춤
            page_setup.FitToPagesTall = False  # 높이는 제한 없음

        # 변경 사항 저장
        workbook.Save()
        print("Changes saved successfully.")

        # 인쇄 실행 및 에러 처리 강화
        # **인쇄 실행 및 에러 처리 부분 제거 또는 주석 처리**
        # 자동 인쇄를 원하지 않으시면 아래 코드를 주석 처리하거나 삭제하세요
        #try:
        #    workbook.PrintOut()
        #    print("Document has been sent to the printer.")
        #except Exception as e:
        #    print(f"Printing error: {e}")
        #    messagebox.showerror("인쇄 에러", f"인쇄 중 에러가 발생했습니다:\n{e}")

    except Exception as e:
        print(f"An error occurred: {e}")
        messagebox.showerror("에러", f"에러가 발생했습니다:\n{e}")

    finally:
        # Excel 애플리케이션 종료
        workbook.Close(SaveChanges=False)
        excel.Quit()
        print("Excel application closed.")

def select_excel_files():
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기
    file_paths = filedialog.askopenfilenames(
        title="Excel 파일 선택",
        filetypes=[("Excel 파일", "*.xls *.xlsx *.xlsm *.xlsb")]
    )
    return file_paths

if __name__ == "__main__":
    file_paths = select_excel_files()
    if file_paths:
        for file_path in file_paths:
            adjust_page_setup(file_path)
        print("모든 작업이 완료되었습니다.")
        # 작업 완료 메시지 박스 표시
        messagebox.showinfo("완료", "모든 작업이 완료되었습니다.")
    else:
        print("파일을 선택하지 않았습니다.")
        messagebox.showwarning("경고", "파일을 선택하지 않았습니다.")
