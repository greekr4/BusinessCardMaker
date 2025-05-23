import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm  # mm 단위 추가
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from tkinter import *
from tkinter import filedialog, messagebox
from fontTools.ttLib import TTFont as FontToolsTTFont
from PIL import Image, ImageTk
import fitz  # PyMuPDF 라이브러리
import os
import tkinter.ttk as ttk
import pandas as pd
import sys

class BusinessCardMaker:
    def __init__(self):
        self.window = Tk()
        self.window.title("BusinessCardEditor")
        self.window.geometry("650x800+100+50") 
        self.window.attributes('-topmost', True)  # 항상 위
        self.window.update() 
        self.window.attributes('-topmost', False)  # 다시 일반 상태로 변경
        # 실행 파일 위치 기준으로 경로 설정
        self.base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.temp_dir = os.path.join(self.base_path, "temp")
        self.font_dir = os.path.join(self.base_path, "font")
        
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        else:
            # 기존 임시 파일 삭제
            self.clear_temp_folder()
        # 사용 가능한 폰트 목록 가져오기
        self.available_fonts = self.get_available_fonts()
        # 입력 필드 리스트 초기화
        self.input_fields = []
        # 템플릿 페이지
        self.template_page = 0
        self.selected_page = 0
        ############## GUI START ##############
        # 템플릿 선택 프레임
        template_frame = Frame(self.window)
        template_frame.pack(pady=10)
        self.template_btn = Button(template_frame, text="템플릿 선택", command=self.select_template)
        self.template_btn.pack(padx=5)
        self.template_path = None  # 템플릿 경로 초기화
        self.template_label = Label(template_frame, text="선택된 템플릿: 없음")
        self.template_label.pack(pady=0)
        # 미리보기 프레임
        preview_frame = LabelFrame(self.window, text="미리보기")
        preview_frame.pack(pady=10, padx=10, fill='both', expand=True)
        self.preview_label = Label(preview_frame, text="템플릿을 선택하면 미리보기가 표시됩니다")
        self.preview_label.pack(pady=5)
        # 미리보기 이미지를 표시할 Label 추가
        self.preview_image_label = Label(preview_frame)
        self.preview_image_label.pack(pady=5)
        # 미리보기 Next/Prev 버튼
        button_frame = Frame(preview_frame)
        button_frame.pack(pady=5)
        self.prev_btn = Button(button_frame, text="<", command=self.toggle_preview)
        self.prev_btn.pack(side=LEFT, padx=5)
        self.label_side = Label(button_frame, text=f"페이지: 0/0")
        self.label_side.pack(side=LEFT, padx=5)
        self.next_btn = Button(button_frame, text=">", command=self.toggle_preview)
        self.next_btn.pack(side=LEFT, padx=5)
        # 엑셀 양식 관련 프레임
        excel_frame = Frame(self.window)
        excel_frame.pack(pady=5)
        Button(excel_frame, text="엑셀 양식 다운로드", command=self.download_excel_template, width=15).pack(side=LEFT, padx=2)
        Button(excel_frame, text="엑셀 양식 불러오기", command=self.upload_excel, width=15).pack(side=LEFT, padx=2)
        # 입력 필드 추가/삭제 버튼 프레임
        field_control_frame = Frame(self.window)
        field_control_frame.pack(pady=5)
        Button(field_control_frame, text="입력 필드 추가", command=self.add_input_field,width=15).pack(side=LEFT, padx=2)
        Button(field_control_frame, text="마지막 필드 삭제", command=self.remove_input_field,width=15).pack(side=LEFT, padx=2)
        # 정보 입력 프레임
        front_frame = LabelFrame(self.window, text="정보 입력")
        front_frame.pack(pady=0, padx=10, fill='both', expand=True)
        # 입력 필드를 담을 스크롤 가능한 프레임
        self.scroll_canvas = Canvas(front_frame)
        scrollbar = Scrollbar(front_frame, orient="vertical", command=self.scroll_canvas.yview)
        self.scrollable_frame = Frame(self.scroll_canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))
        )
        self.scroll_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")
        # 하단 버튼 프레임
        button_frame = Frame(self.window)
        button_frame.pack(pady=5)
        Button(button_frame, text="미리보기", command=self.show_preview).pack(side=LEFT, padx=5)
        Button(button_frame, text="다운로드", command=self.create_namecard).pack(side=LEFT, padx=5)
        ############## GUI END ##############
        # 기본 입력 필드 추가
        self.add_input_field()

    def download_excel_template(self):
        """엑셀 양식 다운로드"""
        try:
            data = {
                'Text': [],
                'X(mm)': [],
                'Y(mm)': [],
                'Font': [],
                'Size(Pt)': [],
                'Page': []
            }
            for field in self.input_fields:
                data['Text'].append(field['text'].get())
                data['X(mm)'].append(field['x_coord'].get())
                data['Y(mm)'].append(field['y_coord'].get())
                data['Font'].append(field['font_var'].get())
                data['Size(Pt)'].append(field['font_size'].get())
                data['Page'].append(field['draw_page'].get())
            df = pd.DataFrame(data)
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("성공", "엑셀 양식이 다운로드되었습니다!")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 생성 중 오류가 발생했습니다: {str(e)}")
        
    def upload_excel(self):
        """엑셀 양식 업로드"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(file_path)
            self.remove_all_input_field()
            for index, row in df.iterrows():
                text_value = "" if pd.isna(row['Text']) else row['Text']
                self.add_input_field(text=text_value, x_coord=row['X(mm)'], y_coord=row['Y(mm)'], font=row['Font'], font_size=row['Size(Pt)'], draw_page=row['Page'])

    def select_template(self): 
        """템플릿 선택"""
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.template_path = path
        template_name = self.template_path.split('/')[-1] if self.template_path else "없음"
        self.template_label.config(text=f"선택된 템플릿: {template_name}")
        self.preview_label.config(text="")
        if self.template_path:
            doc = fitz.open(self.template_path)
            self.template_page = len(doc)
            self.selected_page = 1

        if self.template_path:
            # ./temp/preview.pdf 파일이 이미 존재하면 삭제
            preview_path = "./temp/preview.pdf"
            if os.path.exists(preview_path):
                try:
                    os.remove(preview_path)
                except Exception as e:
                    messagebox.showerror("오류", f"미리보기 파일 삭제 중 오류가 발생했습니다: {str(e)}")
            self.update_preview()  # 템플릿 선택 시 미리보기 업데이트

    def create_namecard(self):
        """명함 생성 (show_preview() 호출 => ./temp/preview.pdf 파일 생성 => 최종 PDF 저장)"""
        if not self.template_path:
            messagebox.showerror("오류", "템플릿을 선택해주세요!")
            return
            
        try:
            self.show_preview()
            # 최종 PDF 저장
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if output_path:
                # ./temp/preview.pdf 파일을 최종 출력 파일로 복사
                import shutil
                try:
                    shutil.copy2("./temp/preview.pdf", output_path)
                    messagebox.showinfo("성공", "명함이 생성되었습니다!")
                except Exception as e:
                    messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {str(e)}")
                
        except Exception as e:
            messagebox.showerror("오류", f"명함 생성 중 오류가 발생했습니다: {str(e)}")

    def update_preview(self):
        """ 미리보기 업데이트 """
        if not self.template_path:
            return
        try:
            # PDF 파일 열기
            # 미리보기 PDF 파일이 없으면 원본 템플릿을 복제
            # 원본 템플릿 파일을 임시 폴더에 복사
            import shutil
            if not os.path.exists("./temp/preview.pdf") and self.template_path:
                try:
                    shutil.copy2(self.template_path, "./temp/preview.pdf")
                except Exception as e:
                    messagebox.showerror("오류", f"미리보기 파일 생성 중 오류가 발생했습니다: {str(e)}")
                    return
            self.preview_path = "./temp/preview.pdf" 
            doc = fitz.open(self.preview_path)
            # 페이지 인덱스 계산 (0부터 시작)
            page_idx = self.selected_page - 1
            # 해당 페이지 객체 가져오기
            page = doc[page_idx]
            # PDF 페이지를 이미지로 변환 (고정 크기로 설정)
            fixed_width = 400  # 고정 너비 픽셀
            fixed_height = 150  # 고정 높이 픽셀
            
            # 원본 페이지 크기 가져오기
            original_rect = page.rect
            original_width = original_rect.width
            original_height = original_rect.height
            
            # 적절한 스케일 계산
            width_scale = fixed_width / original_width
            height_scale = fixed_height / original_height
            scale = min(width_scale, height_scale)  # 비율 유지를 위해 작은 값 사용
            
            # 스케일 적용하여 이미지 생성
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Tkinter에서 표시할 수 있는 형식으로 변환
            photo = ImageTk.PhotoImage(img)
            # 미리보기 이미지 레이블에 이미지 설정 및 크기 고정
            self.preview_image_label.configure(image=photo, width=fixed_width, height=fixed_height)
            self.preview_image_label.image = photo
            # 현재 페이지 정보 업데이트
            self.label_side.config(text=f"페이지: {self.selected_page}/{doc.page_count}")
            # PDF 문서 닫기
            doc.close()
        except Exception as e:
            messagebox.showerror("오류", f"미리보기 생성 중 오류가 발생했습니다: {str(e)}")

    def toggle_preview(self):
        if not self.template_path:
            return
        # 페이지 전환 로직
        if self.selected_page < self.template_page:
            self.selected_page += 1
        else:
            self.selected_page = 1
        self.update_preview()

    def show_preview(self):
        """ 미리보기 업데이트 """
        if not self.template_path:
            messagebox.showerror("오류", "템플릿을 먼저 선택해주세요!")
            return
            
        try:
            # 로딩 중 표시
            self.preview_label.config(text="생성 중...")
            self.window.update() 
            # 원본 템플릿 경로 저장
            original_template = self.template_path
            # 템플릿 PDF 열기
            template = PyPDF2.PdfReader(self.template_path)
            output = PyPDF2.PdfWriter()
            # 모든 페이지에 대해 처리
            for page_idx in range(len(template.pages)):
                # 각 페이지에 해당하는 필드가 있는지 확인
                has_fields_for_page = False
                for field in self.input_fields:
                    if field['draw_page'].get() == str(page_idx+1):
                        has_fields_for_page = True
                        break
                
                if has_fields_for_page:
                    # 각 페이지별 임시 PDF 생성
                    # 템플릿 PDF에서 페이지 크기 가져오기
                    template_page = template.pages[page_idx]
                    template_width = float(template_page.mediabox.width)
                    template_height = float(template_page.mediabox.height)
                    c = canvas.Canvas(f"./temp/temp_page_{page_idx}.pdf", pagesize=(template_width, template_height))
                    # 모든 입력 필드 처리
                    for field in self.input_fields:
                        if field['draw_page'].get() == str(page_idx+1):
                            font_name = field['font_var'].get()
                            font_size = float(field['font_size'].get())
                            x_coord = field['x_coord'].get()
                            y_coord = field['y_coord'].get()
                            text = field['text'].get()
                            pdfmetrics.registerFont(TTFont(font_name, f'./font/{font_name}.ttf'))
                            c.setFont(font_name, font_size)
                            c.drawString(float(x_coord)*mm, float(y_coord)*mm, text)
                    c.save()
                    # 템플릿과 합치기
                    overlay = PyPDF2.PdfReader(f"./temp/temp_page_{page_idx}.pdf")
                    template_page = template.pages[page_idx]
                    template_page.merge_page(overlay.pages[0])
                    output.add_page(template_page)
                else:
                    # 해당 페이지에 필드가 없으면 원본 페이지만 추가
                    template_page = template.pages[page_idx]
                    output.add_page(template_page)
            # 미리보기용 임시 파일 저장
            with open("./temp/preview.pdf", 'wb') as f:
                output.write(f)
            # 미리보기 업데이트
            self.preview_path = "./temp/preview.pdf"  # 임시로 경로 변경
            self.preview_label.config(text="")  # 로딩 메시지 제거
            self.update_preview()
            # 원본 템플릿 경로 복원
            self.template_path = original_template
        except Exception as e:
            messagebox.showerror("오류", f"미리보기 생성 중 오류가 발생했습니다: {str(e)}")
            self.preview_label.config(text="미리보기 생성 실패")

    def add_input_field(self, text="", x_coord="0", y_coord="0", font="", font_size="8", draw_page="1"):
        """ 입력 필드 추가 """
        field_frame = Frame(self.scrollable_frame)
        field_frame.pack(anchor='w', padx=10, pady=3)
        # 값 입력창
        text_entry = Entry(field_frame, width=20)
        text_entry.pack(side=LEFT, padx=(0, 10))
        text_entry.insert(0, text)
        # X좌표 입력
        Label(field_frame, text="X", width=2).pack(side=LEFT)
        x_coord_entry = Entry(field_frame, width=5)
        x_coord_entry.pack(side=LEFT, padx=(0, 10))
        x_coord_entry.insert(0, x_coord)
        # Y좌표 입력
        Label(field_frame, text="Y", width=2).pack(side=LEFT)
        y_coord_entry = Entry(field_frame, width=5)
        y_coord_entry.pack(side=LEFT, padx=(0, 10))
        y_coord_entry.insert(0, y_coord)
        # 폰트 선택 콤보박스
        font_var = StringVar(value=font if font else self.available_fonts[0])
        font_combo = ttk.Combobox(field_frame, textvariable=font_var, values=self.available_fonts, width=15)
        font_combo.pack(side=LEFT, padx=(0, 10))
        # 폰트 크기 입력
        Label(field_frame, text="크기").pack(side=LEFT)
        font_size_entry = Entry(field_frame, width=3)
        font_size_entry.pack(side=LEFT, padx=(0, 10))
        font_size_entry.insert(0, font_size)
        # 페이지 입력
        Label(field_frame, text="페이지").pack(side=LEFT)
        draw_page_entry = Entry(field_frame, width=3)
        draw_page_entry.pack(side=LEFT)
        draw_page_entry.insert(0, draw_page)
        
        # 입력 필드 정보를 리스트에 저장
        field_info = {
            'frame': field_frame,
            'text': text_entry,
            'x_coord': x_coord_entry,
            'y_coord': y_coord_entry,
            'font_var': font_var,
            'font_combo': font_combo,
            'font_size': font_size_entry,
            'draw_page': draw_page_entry
        }
        self.input_fields.append(field_info)
        
        # 스크롤 영역 업데이트
        self.scrollable_frame.update_idletasks()
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def remove_input_field(self):
        """ 마지막 필드 삭제 """
        if self.input_fields:
            field_info = self.input_fields.pop()
            field_info['frame'].destroy()
            
            # 스크롤 영역 업데이트
            self.scrollable_frame.update_idletasks()
            self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def remove_all_input_field(self):
        """ 모든 필드 삭제 """
        for field in self.input_fields:
            field['frame'].destroy()
        self.input_fields = []
        self.scrollable_frame.update_idletasks()
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def clear_temp_folder(self):
        """ ./temp 폴더의 모든 파일을 삭제"""
        for file in os.listdir(self.temp_dir):
            file_path = os.path.join(self.temp_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"임시 파일 삭제 중 오류 발생: {e}")

    def get_available_fonts(self):
        """font 폴더에 있는 모든 폰트 불러오기"""
        fonts = []
        if os.path.exists(self.font_dir):
            for file in os.listdir(self.font_dir):
                if file.lower().endswith('.ttf'):
                    # 파일 이름에서 확장자 제거
                    font_name = os.path.splitext(file)[0]
                    if font_name not in fonts:
                        fonts.append(font_name)
        return fonts

if __name__ == "__main__":
    app = BusinessCardMaker()
    app.window.mainloop()

