import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm  # mm 단위 추가
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.ttfonts import TTFOpenFile
from tkinter import *
from tkinter import filedialog, messagebox
from fontTools.ttLib import TTFont as FontToolsTTFont
from PIL import Image, ImageTk
import fitz  # PyMuPDF 라이브러리
import os

class NameCardMaker:
    def __init__(self):
        self.window = Tk()
        self.window.title("명함 제작기")
        self.window.geometry("500x800+100+50")  # 창 크기 증가, X,Y 위치 100으로 설정
        self.window.attributes('-topmost', True)  # 창을 항상 위에 표시
        self.window.update()  # 창 업데이트
        self.window.attributes('-topmost', False)  # 다시 일반 상태로 변경
        
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
        self.preview_side = StringVar(value="front")
        self.toggle_btn = Button(preview_frame, text="앞면/뒷면 전환", command=self.toggle_preview)
        self.toggle_btn.pack(pady=5)
        
        # 엑셀 양식 관련 프레임
        excel_frame = Frame(self.window)
        excel_frame.pack(pady=5)
        

        # 엑셀 양식 다운로드 버튼
        self.download_btn = Button(excel_frame, text="엑셀 양식 다운로드", command=self.download_excel_template)
        self.download_btn.pack(side=LEFT, padx=5)
        
        # 엑셀 파일 업로드 버튼 
        self.upload_btn = Button(excel_frame, text="엑셀 파일 업로드", command=self.upload_excel)
        self.upload_btn.pack(side=LEFT, padx=5)

        # 프레임
        front_frame = LabelFrame(self.window, text="정보 입력")
        front_frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        position_coord_frame = Frame(front_frame)
        position_coord_frame.pack(anchor='w', padx=10, pady=3)  # anchor='w'로 왼쪽 정렬

        position_frame = Frame(position_coord_frame)
        position_frame.pack(fill=X)
        
        # 직책 입력 필드
        Label(position_frame, text="직책", width=8).pack(side=LEFT)  # 레이블 너비 고정
        self.korean_position_entry = Entry(position_frame, width=20)  # 입력 필드 너비 증가
        self.korean_position_entry.pack(side=LEFT, padx=(0, 10))  # 오른쪽에 여백 추가
        Label(position_frame, text="X", width=2).pack(side=LEFT)
        self.position_x_entry = Entry(position_frame, width=5)
        self.position_x_entry.pack(side=LEFT, padx=(0, 10))
        self.position_x_entry.insert(0, "44")
        Label(position_frame, text="Y", width=2).pack(side=LEFT)
        self.position_y_entry = Entry(position_frame, width=5)
        self.position_y_entry.pack(side=LEFT)
        self.position_y_entry.insert(0, "47")

        # 이름 좌표 입력 프레임
        name_coord_frame = Frame(front_frame)
        name_coord_frame.pack(anchor='w', padx=10, pady=3)  # anchor='w'로 왼쪽 정렬
        Label(name_coord_frame, text="한글이름", width=8).pack(side=LEFT)
        self.korean_name_entry = Entry(name_coord_frame, width=20)
        self.korean_name_entry.pack(side=LEFT, padx=(0, 10))  # 오른쪽에 여백 추가
        Label(name_coord_frame, text="X", width=2).pack(side=LEFT)
        self.name_x_entry = Entry(name_coord_frame, width=5)
        self.name_x_entry.pack(side=LEFT, padx=(0, 10))  # 오른쪽에 여백 추가
        self.name_x_entry.insert(0, "44")
        Label(name_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.name_y_entry = Entry(name_coord_frame, width=5)
        self.name_y_entry.pack(side=LEFT, padx=(0, 10))  # 오른쪽에 여백 추가   
        self.name_y_entry.insert(0, "37")

        # 영문이름 입력 프레임
        eng_name_coord_frame = Frame(front_frame)
        eng_name_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(eng_name_coord_frame, text="영문이름", width=8).pack(side=LEFT)
        self.english_name_entry = Entry(eng_name_coord_frame, width=20)
        self.english_name_entry.pack(side=LEFT, padx=(0, 10))
        Label(eng_name_coord_frame, text="X", width=2).pack(side=LEFT)
        self.eng_name_x_entry = Entry(eng_name_coord_frame, width=5)
        self.eng_name_x_entry.insert(0, "44")
        self.eng_name_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(eng_name_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.eng_name_y_entry = Entry(eng_name_coord_frame, width=5)
        self.eng_name_y_entry.insert(0, "37")
        self.eng_name_y_entry.pack(side=LEFT)

        # 영문직책 입력 프레임
        eng_pos_coord_frame = Frame(front_frame)
        eng_pos_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(eng_pos_coord_frame, text="영문직책", width=8).pack(side=LEFT)
        self.english_position_entry = Entry(eng_pos_coord_frame, width=20)
        self.english_position_entry.pack(side=LEFT, padx=(0, 10))
        Label(eng_pos_coord_frame, text="X", width=2).pack(side=LEFT)
        self.eng_pos_x_entry = Entry(eng_pos_coord_frame, width=5)
        self.eng_pos_x_entry.insert(0, "44")
        self.eng_pos_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(eng_pos_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.eng_pos_y_entry = Entry(eng_pos_coord_frame, width=5)
        self.eng_pos_y_entry.insert(0, "47")
        self.eng_pos_y_entry.pack(side=LEFT)

        # 영문직책2 입력 프레임
        eng_pos2_coord_frame = Frame(front_frame)
        eng_pos2_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(eng_pos2_coord_frame, text="영문직책2", width=8).pack(side=LEFT)
        self.english_position2_entry = Entry(eng_pos2_coord_frame, width=20)
        self.english_position2_entry.pack(side=LEFT, padx=(0, 10))
        Label(eng_pos2_coord_frame, text="X", width=2).pack(side=LEFT)
        self.eng_pos2_x_entry = Entry(eng_pos2_coord_frame, width=5)
        self.eng_pos2_x_entry.insert(0, "44")
        self.eng_pos2_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(eng_pos2_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.eng_pos2_y_entry = Entry(eng_pos2_coord_frame, width=5)
        self.eng_pos2_y_entry.insert(0, "44")
        self.eng_pos2_y_entry.pack(side=LEFT)

        # TEL 입력 프레임
        tel_coord_frame = Frame(front_frame)
        tel_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(tel_coord_frame, text="TEL", width=8).pack(side=LEFT)
        self.tel_entry = Entry(tel_coord_frame, width=20)
        self.tel_entry.pack(side=LEFT, padx=(0, 10))
        Label(tel_coord_frame, text="X", width=2).pack(side=LEFT)
        self.tel_x_entry = Entry(tel_coord_frame, width=5)
        self.tel_x_entry.insert(0, "55")
        self.tel_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(tel_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.tel_y_entry = Entry(tel_coord_frame, width=5)
        self.tel_y_entry.insert(0, "14.5")
        self.tel_y_entry.pack(side=LEFT)

        # FAX 입력 프레임
        fax_coord_frame = Frame(front_frame)
        fax_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(fax_coord_frame, text="FAX", width=8).pack(side=LEFT)
        self.fax_entry = Entry(fax_coord_frame, width=20)
        self.fax_entry.pack(side=LEFT, padx=(0, 10))
        Label(fax_coord_frame, text="X", width=2).pack(side=LEFT)
        self.fax_x_entry = Entry(fax_coord_frame, width=5)
        self.fax_x_entry.insert(0, "55")
        self.fax_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(fax_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.fax_y_entry = Entry(fax_coord_frame, width=5)
        self.fax_y_entry.insert(0, "11.5")
        self.fax_y_entry.pack(side=LEFT)

        # H.P 입력 프레임
        hp_coord_frame = Frame(front_frame)
        hp_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(hp_coord_frame, text="H.P", width=8).pack(side=LEFT)
        self.hp_entry = Entry(hp_coord_frame, width=20)
        self.hp_entry.pack(side=LEFT, padx=(0, 10))
        Label(hp_coord_frame, text="X", width=2).pack(side=LEFT)
        self.hp_x_entry = Entry(hp_coord_frame, width=5)
        self.hp_x_entry.insert(0, "55")
        self.hp_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(hp_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.hp_y_entry = Entry(hp_coord_frame, width=5)
        self.hp_y_entry.insert(0, "8.5")
        self.hp_y_entry.pack(side=LEFT)

        # TEL2 입력 프레임
        tel2_coord_frame = Frame(front_frame)
        tel2_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(tel2_coord_frame, text="TEL2", width=8).pack(side=LEFT)
        self.tel2_entry = Entry(tel2_coord_frame, width=20)
        self.tel2_entry.pack(side=LEFT, padx=(0, 10))
        Label(tel2_coord_frame, text="X", width=2).pack(side=LEFT)
        self.tel2_x_entry = Entry(tel2_coord_frame, width=5)
        self.tel2_x_entry.insert(0, "55")
        self.tel2_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(tel2_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.tel2_y_entry = Entry(tel2_coord_frame, width=5)
        self.tel2_y_entry.insert(0, "14.5")
        self.tel2_y_entry.pack(side=LEFT)

        # FAX2 입력 프레임
        fax2_coord_frame = Frame(front_frame)
        fax2_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(fax2_coord_frame, text="FAX2", width=8).pack(side=LEFT)
        self.fax2_entry = Entry(fax2_coord_frame, width=20)
        self.fax2_entry.pack(side=LEFT, padx=(0, 10))
        Label(fax2_coord_frame, text="X", width=2).pack(side=LEFT)
        self.fax2_x_entry = Entry(fax2_coord_frame, width=5)
        self.fax2_x_entry.insert(0, "55")
        self.fax2_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(fax2_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.fax2_y_entry = Entry(fax2_coord_frame, width=5)
        self.fax2_y_entry.insert(0, "11.5")
        self.fax2_y_entry.pack(side=LEFT)

        # H.P2 입력 프레임
        hp2_coord_frame = Frame(front_frame)
        hp2_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(hp2_coord_frame, text="H.P2", width=8).pack(side=LEFT)
        self.hp2_entry = Entry(hp2_coord_frame, width=20)
        self.hp2_entry.pack(side=LEFT, padx=(0, 10))
        Label(hp2_coord_frame, text="X", width=2).pack(side=LEFT)
        self.hp2_x_entry = Entry(hp2_coord_frame, width=5)
        self.hp2_x_entry.insert(0, "55")
        self.hp2_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(hp2_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.hp2_y_entry = Entry(hp2_coord_frame, width=5)
        self.hp2_y_entry.insert(0, "8.5")
        self.hp2_y_entry.pack(side=LEFT)

        # E-mail 입력 프레임
        email_coord_frame = Frame(front_frame)
        email_coord_frame.pack(anchor='w', padx=10, pady=3)
        Label(email_coord_frame, text="E-mail", width=8).pack(side=LEFT)
        self.email_entry = Entry(email_coord_frame, width=20)
        self.email_entry.pack(side=LEFT, padx=(0, 10))
        Label(email_coord_frame, text="X", width=2).pack(side=LEFT)
        self.email_x_entry = Entry(email_coord_frame, width=5)
        self.email_x_entry.insert(0, "55")
        self.email_x_entry.pack(side=LEFT, padx=(0, 10))
        Label(email_coord_frame, text="Y", width=2).pack(side=LEFT)
        self.email_y_entry = Entry(email_coord_frame, width=5)
        self.email_y_entry.insert(0, "5.5")
        self.email_y_entry.pack(side=LEFT)

        Label(front_frame, text="").pack()
        
        # 버튼들을 담을 프레임 생성
        button_frame = Frame(self.window)
        button_frame.pack(pady=5)

        # 테스트 버튼
        self.test_btn = Button(button_frame, text="샘플 불러오기", command=self.test_input)
        self.test_btn.pack(side=LEFT, padx=5)
        
        # 테스트하기2 버튼 (미리보기용)
        self.test_preview_btn = Button(button_frame, text="미리보기", command=self.test_preview)
        self.test_preview_btn.pack(side=LEFT, padx=5)
        
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
            self.update_preview()  # 템플릿 선택 시 미리보기 업데이트

    def test_input(self):
        # 테스트용 샘플 데이터 입력
        self.korean_position_entry.delete(0, END)
        self.korean_position_entry.insert(0, "관리팀 / 대리")
        
        self.korean_name_entry.delete(0, END) 
        self.korean_name_entry.insert(0, "노       하       은")
        
        self.english_name_entry.delete(0, END)
        self.english_name_entry.insert(0, "ROH HA EUN")
        
        self.english_position_entry.delete(0, END)
        self.english_position_entry.insert(0, "Management Team/")

        self.english_position2_entry.delete(0, END)
        self.english_position2_entry.insert(0, "Assistant Manager")
        
        self.tel_entry.delete(0, END)
        self.tel_entry.insert(0, "(031) 8077-4567")
        
        self.fax_entry.delete(0, END)
        self.fax_entry.insert(0, "(031) 8055-8599")
        
        self.hp_entry.delete(0, END)
        self.hp_entry.insert(0, "010-7727-9972")

        self.tel2_entry.delete(0, END)
        self.tel2_entry.insert(0, "+82-(0)31-8077-8002")

        self.fax2_entry.delete(0, END)
        self.fax2_entry.insert(0, "+82-(0)31-8055-8599")
        
        self.hp2_entry.delete(0, END)
        self.hp2_entry.insert(0, "+82-(0)10-7727-9972")

        self.email_entry.delete(0, END)
        self.email_entry.insert(0, "sypark@doowoncorp.com")

    def create_namecard(self):
        if not self.template_path:
            messagebox.showerror("오류", "템플릿을 선택해주세요!")
            return
            
        try:
            # OTF를 TTF로 변환
            otf_font = FontToolsTTFont('./font/AdobeGothicStd-Bold.otf')
            otf_font.save('./font/AdobeGothicStd-Bold.ttf')
            
            output = PyPDF2.PdfWriter()
            template = PyPDF2.PdfReader(self.template_path)
            
            # 앞면 생성
            c = canvas.Canvas("./temp/temp_front.pdf", pagesize=letter)
            # 폰트 등록
            pdfmetrics.registerFont(TTFont('Franklin Gothic Demi', './font/FRADM.TTF'))
            pdfmetrics.registerFont(TTFont('나눔바른고딕', './font/NanumBarunGothic-YetHangul.TTF'))
            pdfmetrics.registerFont(TTFont('NotoSansKR_Regular', './font/NotoSansKR-Regular_0.ttf'))
            pdfmetrics.registerFont(TTFont('NotoSansKR_ExtraBold', './font/NotoSansKR-ExtraBold_0.ttf'))
            # mm 단위로 좌표 지정

            # TEL, FAX, HP, EMAIL #########
            c.setFont("Franklin Gothic Demi", 8)
            c.drawString(float(self.tel_x_entry.get())*mm, float(self.tel_y_entry.get())*mm, self.tel_entry.get())
            c.drawString(float(self.fax_x_entry.get())*mm, float(self.fax_y_entry.get())*mm, self.fax_entry.get())
            c.drawString(float(self.hp_x_entry.get())*mm, float(self.hp_y_entry.get())*mm, self.hp_entry.get())
            c.drawString(float(self.email_x_entry.get())*mm, float(self.email_y_entry.get())*mm, self.email_entry.get())
            # 한글 직책 ########
            c.setFont("나눔바른고딕", 7)
            c.drawString(float(self.position_x_entry.get())*mm, float(self.position_y_entry.get())*mm, self.korean_position_entry.get())

            # 한글 이름 ########
            c.setFont("NotoSansKR_ExtraBold", 15)
            c.drawString(float(self.name_x_entry.get())*mm, float(self.name_y_entry.get())*mm, self.korean_name_entry.get())

            c.save()
            
            # 뒷면 생성
            c = canvas.Canvas("./temp/temp_back.pdf", pagesize=letter)
            # TEL, FAX, HP, EMAIL #########
            c.setFont("Franklin Gothic Demi", 8)
            c.drawString(float(self.tel2_x_entry.get())*mm, float(self.tel2_y_entry.get())*mm, self.tel2_entry.get())
            c.drawString(float(self.fax2_x_entry.get())*mm, float(self.fax2_y_entry.get())*mm, self.fax2_entry.get())
            c.drawString(float(self.hp2_x_entry.get())*mm, float(self.hp2_y_entry.get())*mm, self.hp2_entry.get())
            c.drawString(float(self.email_x_entry.get())*mm, float(self.email_y_entry.get())*mm, self.email_entry.get())
            
            # 영문 직책
            c.setFont("Franklin Gothic Demi", 8)
            c.drawString(float(self.eng_pos_x_entry.get())*mm, float(self.eng_pos_y_entry.get())*mm, self.english_position_entry.get())
            c.drawString(float(self.eng_pos2_x_entry.get())*mm, float(self.eng_pos2_y_entry.get())*mm, self.english_position2_entry.get())

            c.setFont("Franklin Gothic Demi", 15)
            c.drawString(float(self.eng_name_x_entry.get())*mm, float(self.eng_name_y_entry.get())*mm, self.english_name_entry.get())
            c.save()
            
            # 앞면 합치기 (템플릿의 첫 페이지)
            front_overlay = PyPDF2.PdfReader("./temp/temp_front.pdf")
            front_page = template.pages[0]  # 첫 페이지가 앞면
            front_page.merge_page(front_overlay.pages[0])
            output.add_page(front_page)
            
            # 뒷면 합치기 (템플릿의 두번째 페이지)
            back_overlay = PyPDF2.PdfReader("./temp/temp_back.pdf")
            back_page = template.pages[1]  # 두번째 페이지가 뒷면
            back_page.merge_page(back_overlay.pages[0])
            output.add_page(back_page)
            
            # 최종 PDF 저장
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if output_path:
                with open(output_path, 'wb') as f:
                    output.write(f)
                messagebox.showinfo("성공", "명함이 생성되었습니다!")
                
        except Exception as e:
            messagebox.showerror("오류", f"명함 생성 중 오류가 발생했습니다: {str(e)}")

    def update_preview(self):
        if not self.template_path:
            return
            
        try:
            # PDF 파일 열기
            doc = fitz.open(self.template_path)
            page_num = 0 if self.preview_side.get() == "front" else 1
            page = doc[page_num]
            
            # PDF 페이지를 이미지로 변환
            pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8))  # 크기를  축소
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Tkinter에서 표시할 수 있는 형식으로 변환
            photo = ImageTk.PhotoImage(img)
            
            # 미리보기 이미지 업데이트
            self.preview_image_label.configure(image=photo)
            self.preview_image_label.image = photo  # 참조 유지
            
            doc.close()
            
        except Exception as e:
            messagebox.showerror("오류", f"미리보기 생성 중 오류가 발생했습니다: {str(e)}")

    def toggle_preview(self):
        # 현재 template_path 저장
        current_path = self.template_path
        
        # 미리보기 PDF가 존재하는 경우 해당 경로 사용
        if os.path.exists("./temp/preview.pdf"):
            self.template_path = "./temp/preview.pdf"
        
        if self.preview_side.get() == "front":
            self.preview_side.set("back")
        else:
            self.preview_side.set("front")
        
        self.update_preview()
        
        # 원래 template_path 복원
        self.template_path = current_path

    def test_preview(self):
        if not self.template_path:
            messagebox.showerror("오류", "템플릿을 먼저 선택해주세요!")
            return
            
        try:
            # 원본 템플릿 경로 저장
            original_template = self.template_path
            
            # 임시 PDF 생성
            c = canvas.Canvas("./temp/temp_front.pdf", pagesize=letter)
            # 폰트 등록
            pdfmetrics.registerFont(TTFont('Franklin Gothic Demi', './font/FRADM.TTF'))
            pdfmetrics.registerFont(TTFont('나눔바른고딕', './font/NanumBarunGothic-YetHangul.TTF'))
            pdfmetrics.registerFont(TTFont('NotoSansKR_Regular', './font/NotoSansKR-Regular_0.ttf'))
            pdfmetrics.registerFont(TTFont('NotoSansKR_ExtraBold', './font/NotoSansKR-ExtraBold_0.ttf'))

            # 앞면 정보 입력
            c.setFont("Franklin Gothic Demi", 8)
            c.drawString(float(self.tel_x_entry.get())*mm, float(self.tel_y_entry.get())*mm, self.tel_entry.get())
            c.drawString(float(self.fax_x_entry.get())*mm, float(self.fax_y_entry.get())*mm, self.fax_entry.get())
            c.drawString(float(self.hp_x_entry.get())*mm, float(self.hp_y_entry.get())*mm, self.hp_entry.get())
            c.drawString(float(self.email_x_entry.get())*mm, float(self.email_y_entry.get())*mm, self.email_entry.get())
            
            c.setFont("나눔바른고딕", 7)
            c.drawString(float(self.position_x_entry.get())*mm, float(self.position_y_entry.get())*mm, self.korean_position_entry.get())
            
            c.setFont("NotoSansKR_ExtraBold", 15)
            c.drawString(float(self.name_x_entry.get())*mm, float(self.name_y_entry.get())*mm, self.korean_name_entry.get())
            
            c.save()
            
            # 뒷면 생성
            c = canvas.Canvas("./temp/temp_back.pdf", pagesize=letter)
            c.setFont("Franklin Gothic Demi", 8)
            c.drawString(float(self.tel2_x_entry.get())*mm, float(self.tel2_y_entry.get())*mm, self.tel2_entry.get())
            c.drawString(float(self.fax2_x_entry.get())*mm, float(self.fax2_y_entry.get())*mm, self.fax2_entry.get())
            c.drawString(float(self.hp2_x_entry.get())*mm, float(self.hp2_y_entry.get())*mm, self.hp2_entry.get())
            c.drawString(float(self.email_x_entry.get())*mm, float(self.email_y_entry.get())*mm, self.email_entry.get())
            
            c.setFont("Franklin Gothic Demi", 8)
            c.drawString(float(self.eng_pos_x_entry.get())*mm, float(self.eng_pos_y_entry.get())*mm, self.english_position_entry.get())
            c.drawString(float(self.eng_pos2_x_entry.get())*mm, float(self.eng_pos2_y_entry.get())*mm, self.english_position2_entry.get())
            
            c.setFont("Franklin Gothic Demi", 15)
            c.drawString(float(self.eng_name_x_entry.get())*mm, float(self.eng_name_y_entry.get())*mm, self.english_name_entry.get())
            c.save()
            
            # 템플릿과 합치기
            output = PyPDF2.PdfWriter()
            template = PyPDF2.PdfReader(self.template_path)
            
            # 앞면 합치기
            front_overlay = PyPDF2.PdfReader("./temp/temp_front.pdf")
            front_page = template.pages[0]
            front_page.merge_page(front_overlay.pages[0])
            output.add_page(front_page)
            
            # 뒷면 합치기
            back_overlay = PyPDF2.PdfReader("./temp/temp_back.pdf")
            back_page = template.pages[1]
            back_page.merge_page(back_overlay.pages[0])
            output.add_page(back_page)
            
            # 미리보기용 임시 파일 저장
            with open("./temp/preview.pdf", 'wb') as f:
                output.write(f)
            
            # 미리보기 업데이트
            self.template_path = "./temp/preview.pdf"  # 임시로 경로 변경
            self.update_preview()
            
            # 원본 템플릿 경로 복원
            self.template_path = original_template
            messagebox.showinfo("성공", "미리보기가 업데이트 되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"미리보기 생성 중 오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    app = NameCardMaker()
    app.window.mainloop()
