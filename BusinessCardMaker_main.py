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

class BusinessCardMaker:
    def __init__(self):
        self.window = Tk()
        self.window.title("명함 제작기")
        self.window.geometry("650x800+100+50")  # 창 크기 증가, X,Y 위치 100으로 설정
        self.window.attributes('-topmost', True)  # 창을 항상 위에 표시
        self.window.update()  # 창 업데이트
        self.window.attributes('-topmost', False)  # 다시 일반 상태로 변경
        # 임시 폴더 생성 및 관리
        self.temp_dir = "./temp"
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
        
        # 템플릿 선택 프레임
        template_frame = Frame(self.window)
        template_frame.pack(pady=10)
        
        self.template_btn = Button(template_frame, text="템플릿 선택", command=self.select_template)
        self.template_btn.pack(padx=5)
        
        self.template_path = None  # 템플릿 경로 초기화
        self.template_label = Label(template_frame, text="선택된 템플릿: 없음")
        self.template_label.pack(pady=0)
        
        
        # 미리보기 프레임 추가
        preview_frame = LabelFrame(self.window, text="미리보기")
        preview_frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        self.preview_label = Label(preview_frame, text="템플릿을 선택하면 미리보기가 표시됩니다")
        self.preview_label.pack(pady=5)
        
        # 미리보기 이미지를 표시할 Label 추가
        self.preview_image_label = Label(preview_frame)
        self.preview_image_label.pack(pady=5)
        
        # 앞면/뒷면 전환 버튼
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
        

        # 엑셀 관련 버튼들
        Button(excel_frame, text="엑셀 양식 다운로드", command=self.download_excel_template, width=15).pack(side=LEFT, padx=2)
        Button(excel_frame, text="엑셀 파일 업로드", command=self.upload_excel, width=15).pack(side=LEFT, padx=2)

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
        
        # 입력 필드 관리를 위한 리스트
        self.input_fields = []
        
        
        # 기본 입력 필드 추가
        self.add_input_field()
        
        # 버튼들을 담을 프레임 생성
        button_frame = Frame(self.window)
        button_frame.pack(pady=5)

        # 테스트 버튼
        self.test_btn = Button(button_frame, text="샘플 불러오기", command=self.test_input)
        self.test_btn.pack(side=LEFT, padx=5)
        
        # 미리보기 버튼
        Button(button_frame, text="미리보기", command=self.show_preview).pack(side=LEFT, padx=5)
        
        # 생성 버튼
        self.create_btn = Button(button_frame, text="명함 생성", command=self.create_namecard)
        self.create_btn.pack(side=LEFT, padx=5)
        
        self.template_path = None  # 단일 템플릿 경로로 변경
        

    def download_excel_template(self):
        # 엑셀 파일 다운로드
        try:
            import pandas as pd
            # 엑셀 파일 생성
            df = pd.DataFrame({
                '한글직책': ['관리팀 / 대리'],
                '한글이름': ['노       하       은'],
                '영문이름': ['ROH HA EUN'],
                '영문직책1': ['Management Team/'],
                '영문직책2': ['Assistant Manager'],
                '전화번호': ['(031) 8077-4567'],
                '팩스': ['(031) 8055-8599'],
                '휴대폰': ['010-7727-9972'],
                '영문전화': ['+82-(0)31-8077-8002'],
                '영문팩스': ['+82-(0)31-8055-8599'],
                '영문휴대폰': ['+82-(0)10-7727-9972'],
                '이메일': ['example@email.com'],
                '한글직책X': [44],
                '한글직책Y': [47],
                '한글이름X': [44],
                '한글이름Y': [37],
                '영문이름X': [44],
                '영문이름Y': [37],
                '영문직책1X': [44],
                '영문직책1Y': [37],
                '영문직책2X': [44],
                '영문직책2Y': [37],
                '전화번호X': [55],
                '전화번호Y': [11.5],
                '팩스X': [55],
                '팩스Y': [11.5],
                '휴대폰X': [55],
                '휴대폰Y': [8.5],
                '영문전화X': [55],
                '영문전화Y': [11.5],
                '영문팩스X': [55],
                '영문팩스Y': [11.5],
                '영문휴대폰X': [55],
                '영문휴대폰Y': [8.5],
                '이메일X': [55],
                '이메일Y': [5.5]
            })
            
            # 파일 저장 대화상자 열기
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
        print('test ')


    def select_template(self):  # 'side' 매개변수 제거
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.template_path = path

        template_name = self.template_path.split('/')[-1] if self.template_path else "없음"
        self.template_label.config(text=f"선택된 템플릿: {template_name}")

        if self.template_path:
            # PDF 파일 열기
            doc = fitz.open(self.template_path)
            self.template_page = len(doc)
            self.selected_page = 1
            self.label_side.config(text=f"페이지: {self.selected_page}/{self.template_page}")

            print(self.template_page)
        
        if self.template_path:
            self.update_preview()  # 템플릿 선택 시 미리보기 업데이트

    def test_input(self):
        # 테스트용 샘플 데이터 입력
        # self.korean_position_entry.delete(0, END)
        # self.korean_position_entry.insert(0, "관리팀 / 대리")
        
        # self.korean_name_entry.delete(0, END) 
        # self.korean_name_entry.insert(0, "노       하       은")
        
        # self.english_name_entry.delete(0, END)
        # self.english_name_entry.insert(0, "ROH HA EUN")
        
        # self.english_position_entry.delete(0, END)
        # self.english_position_entry.insert(0, "Management Team/")

        # self.english_position2_entry.delete(0, END)
        # self.english_position2_entry.insert(0, "Assistant Manager")
        
        # self.tel_entry.delete(0, END)
        # self.tel_entry.insert(0, "(031) 8077-4567")
        
        # self.fax_entry.delete(0, END)
        # self.fax_entry.insert(0, "(031) 8055-8599")
        
        # self.hp_entry.delete(0, END)
        # self.hp_entry.insert(0, "010-7727-9972")

        # self.tel2_entry.delete(0, END)
        # self.tel2_entry.insert(0, "+82-(0)31-8077-8002")

        # self.fax2_entry.delete(0, END)
        # self.fax2_entry.insert(0, "+82-(0)31-8055-8599")
        
        # self.hp2_entry.delete(0, END)
        # self.hp2_entry.insert(0, "+82-(0)10-7727-9972")

        # self.email_entry.delete(0, END)
        # self.email_entry.insert(0, "sypark@doowoncorp.com")
        print('test')

    def create_namecard(self):
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
        if not self.template_path:
            return
            
        try:
            # PDF 파일 열기
            # 미리보기 PDF 파일이 없으면 원본 템플릿을 복제
            if not os.path.exists("./temp/preview.pdf") and self.template_path:
                # 원본 템플릿 파일을 임시 폴더에 복사
                import shutil
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
            
            # PDF 페이지를 이미지로 변환
            pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Tkinter에서 표시할 수 있는 형식으로 변환
            photo = ImageTk.PhotoImage(img)
            
            # 미리보기 이미지 레이블에 이미지 설정
            self.preview_image_label.configure(image=photo)
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
        if not self.template_path:
            messagebox.showerror("오류", "템플릿을 먼저 선택해주세요!")
            return
            
        try:
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
                    c = canvas.Canvas(f"./temp/temp_page_{page_idx}.pdf", pagesize=letter)
                    
                    # 모든 입력 필드 처리
                    for field in self.input_fields:
                        if field['draw_page'].get() == str(page_idx+1):
                            font_name = field['font_var'].get()
                            font_size = float(field['font_size'].get())
                            x_coord = field['x_coord'].get()
                            y_coord = field['y_coord'].get()
                            value = field['value'].get()
                            pdfmetrics.registerFont(TTFont(font_name, f'./font/{font_name}.ttf'))
                            c.setFont(font_name, font_size)
                            c.drawString(float(x_coord)*mm, float(y_coord)*mm, value)
                    
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
            self.update_preview()
            
            # 원본 템플릿 경로 복원
            self.template_path = original_template
            messagebox.showinfo("성공", "미리보기가 업데이트 되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"미리보기 생성 중 오류가 발생했습니다: {str(e)}")

    def add_input_field(self):
        field_frame = Frame(self.scrollable_frame)
        field_frame.pack(anchor='w', padx=10, pady=3)
        
        # 값 입력창
        value = Entry(field_frame, width=20)
        value.pack(side=LEFT, padx=(0, 10))
        
        # X, Y 좌표 입력
        Label(field_frame, text="X", width=2).pack(side=LEFT)
        x_coord = Entry(field_frame, width=5)
        x_coord.insert(0, "0")
        x_coord.pack(side=LEFT, padx=(0, 10))
        
        Label(field_frame, text="Y", width=2).pack(side=LEFT)
        y_coord = Entry(field_frame, width=5)
        y_coord.insert(0, "0")
        y_coord.pack(side=LEFT, padx=(0, 10))
        
        # 폰트 선택 콤보박스
        font_var = StringVar(value=self.available_fonts[0])
        font_combo = ttk.Combobox(field_frame, textvariable=font_var, values=self.available_fonts, width=15)
        font_combo.pack(side=LEFT, padx=(0, 10))
        
        # 폰트 크기 입력
        Label(field_frame, text="크기").pack(side=LEFT)
        font_size = Entry(field_frame, width=3)
        font_size.insert(0, "8")
        font_size.pack(side=LEFT, padx=(0, 10))

        # 앞/뒷면
        Label(field_frame, text="페이지").pack(side=LEFT)
        draw_page = Entry(field_frame, width=3)
        draw_page.insert(0, "1")
        draw_page.pack(side=LEFT)
        
        # 입력 필드 정보를 리스트에 저장
        field_info = {
            'frame': field_frame,
            'value': value,
            'x_coord': x_coord,
            'y_coord': y_coord,
            'font_var': font_var,
            'font_combo': font_combo,
            'font_size': font_size,
            'draw_page': draw_page
        }
        self.input_fields.append(field_info)
        
        # 스크롤 영역 업데이트
        self.scrollable_frame.update_idletasks()
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def remove_input_field(self):
        """마지막 입력 필드 세트를 삭제하는 메소드"""
        if self.input_fields:
            field_info = self.input_fields.pop()
            field_info['frame'].destroy()
            
            # 스크롤 영역 업데이트
            self.scrollable_frame.update_idletasks()
            self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def get_dynamic_fields(self):
        """동적 입력 필드의 값들을 반환하는 메소드"""
        fields = []
        for field in self.input_fields:
            fields.append({
                'value': field['value'].get(),
                'x': field['x_coord'].get(),
                'y': field['y_coord'].get(),
                'font': field['font_var'].get(),
                'font_size': field['font_size'].get()
            })
        return fields

    def clear_temp_folder(self):
        """임시 폴더의 모든 파일을 삭제하는 메소드"""
        for file in os.listdir(self.temp_dir):
            file_path = os.path.join(self.temp_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"임시 파일 삭제 중 오류 발생: {e}")

    def get_available_fonts(self):
        """사용 가능한 폰트 목록을 반환하는 메소드"""
        fonts = []
        font_dir = './font'
        if os.path.exists(font_dir):
            for file in os.listdir(font_dir):
                if file.lower().endswith('.ttf'):
                    # 파일 이름에서 확장자 제거
                    font_name = os.path.splitext(file)[0]
                    if font_name not in fonts:
                        fonts.append(font_name)
        return fonts if fonts else ["Arial"]  # 기본 폰트 제공

if __name__ == "__main__":
    app = BusinessCardMaker()
    app.window.mainloop()
