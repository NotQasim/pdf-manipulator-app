import customtkinter as ctk

from tkinter import messagebox
from pdfmanipulator import PdfManipulator, merge_pdfs, convert_word_to_pdf


ctk.set_appearance_mode('dark')

class PdfManipulatorApp(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title('PDF Manipulator App')
        self.geometry('900x500')
        self.resizable(0, 0)

        self.titlebox = ctk.CTkLabel(self, text='PDF Manipulator', height=200, width=900, text_color='#FEE715', fg_color='#101820', font=('', 50))
        self.titlebox.grid(row=0, column=0, columnspan=5, pady=1, sticky='ew')

        self.text_btn = ctk.CTkButton(self, text='Extract Text from PDF', command=self.extract_text, text_color='#FEE715', fg_color='#101820')
        self.text_btn.grid(row=1, column=0, padx=10, pady=50, sticky='ew')

        self.images_btn = ctk.CTkButton(self, text='Extract Images from PDF', command=self.extract_images, text_color='#FEE715', fg_color='#101820')
        self.images_btn.grid(row=1, column=1, padx=10, pady=50, sticky='ew')

        self.metadata_btn = ctk.CTkButton(self, text='Extract Metadata from PDF', command=self.extract_metadata, text_color='#FEE715', fg_color='#101820')
        self.metadata_btn.grid(row=1, column=2, padx=10, pady=50, sticky='ew')

        self.stamp_btn = ctk.CTkButton(self, text='Add Stamp to PDF', command=self.add_stamp, text_color='#FEE715', fg_color='#101820')
        self.stamp_btn.grid(row=1, column=3, padx=10, pady=50, sticky='ew')

        self.watermark_btn = ctk.CTkButton(self, text='Add Watermark to PDF', command=self.add_watermark, text_color='#FEE715', fg_color='#101820')
        self.watermark_btn.grid(row=1, column=4, padx=10, pady=50, sticky='ew')

        self.trim_btn = ctk.CTkButton(self, text='Trim PDF', command=self.trim_pdf, text_color='#FEE715', fg_color='#101820')
        self.trim_btn.grid(row=2, column=0, padx=10, sticky='ew')

        self.encrypt_btn = ctk.CTkButton(self, text='Encrypt PDF', command=self.encrypt_pdf, text_color='#FEE715', fg_color='#101820')
        self.encrypt_btn.grid(row=2, column=1, padx=10, sticky='ew')

        self.compress_btn = ctk.CTkButton(self, text='Compress PDF', command=self.compress_pdf, text_color='#FEE715', fg_color='#101820')
        self.compress_btn.grid(row=2, column=2, padx=10, sticky='ew')

        self.merge_btn = ctk.CTkButton(self, text='Merge PDFs', command=self.merge_pdfs, text_color='#FEE715', fg_color='#101820')
        self.merge_btn.grid(row=2, column=3, padx=10, sticky='ew')

        self.convert_btn = ctk.CTkButton(self, text='Convert Word to PDF', command=self.convert_word_to_pdf, text_color='#FEE715', fg_color='#101820')
        self.convert_btn.grid(row=2, column=4, padx=10, sticky='ew')

        self.namebox = ctk.CTkLabel(self, text='Made By: Qasim Ali Sabiri', height=50, width=900, text_color='#FEE715', fg_color='#101820', font=('', 12))
        self.namebox.grid(row=3, column=0, columnspan=5, pady=92, sticky='ew')


    @staticmethod
    def select_filename(type: str, name: str) -> str:
        messagebox.showinfo('Select File', f'Select {name} File.')
        filename: str = ctk.filedialog.askopenfilename(filetypes=[(f'{type.upper()} files', f'*.{type.lower()}')])
        return filename


    def extract_text(self) -> None:
        try:
            pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
            if not pdf_filename: return
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.extract_text()
            messagebox.showinfo('Success', 'Text extracted successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def extract_images(self) -> None:
        try:
            pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
            if not pdf_filename: return
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.extract_images()
            messagebox.showinfo('Success', 'Images extracted successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def extract_metadata(self) -> None:
        try:
            pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
            if not pdf_filename: return
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.extract_metadata()
            messagebox.showinfo('Success', 'Metadata extracted successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def add_stamp(self) -> None:
        try:
            pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
            stamp_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='Stamp')
            if (not pdf_filename) or (not stamp_filename): return
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.add_stamp(stamp_filename)
            messagebox.showinfo('Success', 'Stamp added successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def add_watermark(self) -> None:
        try:
            pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
            watermark_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='Watermark')
            if (not pdf_filename) or (not watermark_filename): return
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.add_stamp(watermark_filename)
            messagebox.showinfo('Success', 'Watermark added successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def trim_pdf(self) -> None:
        pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
        if not pdf_filename: return

        new_window = ctk.CTkToplevel()
        new_window.title('Trim PDF')
        new_window.geometry('300x300')
        new_window.resizable(0, 0)

        first_page_label = ctk.CTkLabel(new_window, text='First Page: ', text_color='#FEE715', fg_color='#101820')
        first_page_label.grid(row=0, column=0, padx=5, pady=50, sticky='ew')
        last_page_label = ctk.CTkLabel(new_window, text='Last Page: ', text_color='#FEE715', fg_color='#101820')
        last_page_label.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

        first_page_entry = ctk.CTkEntry(new_window, text_color='#FEE715', fg_color='#101820')
        first_page_entry.grid(row=0, column=1, padx=25, pady=50, sticky='ew')
        last_page_entry = ctk.CTkEntry(new_window, text_color='#FEE715', fg_color='#101820')
        last_page_entry.grid(row=1, column=1, padx=25, pady=5, sticky='ew')

        trim_pdf_btn = ctk.CTkButton(new_window, text='Trim', command=lambda: self.trim_pdf_btn_clicked(pdf_filename, first_page_entry.get(), last_page_entry.get(), new_window), text_color='#FEE715', fg_color='#101820')
        trim_pdf_btn.grid(row=2, column=0, columnspan=2, padx=5, pady=50, sticky='ew')

        
    def trim_pdf_btn_clicked(self, pdf_filename: str, first_page: str, last_page: str, window: ctk.CTkToplevel) -> None:
        try:
            first_page = int(first_page)
            last_page = int(last_page)
        except ValueError:
            messagebox.showerror('Invalid Input', 'Enter valid numbers for the first and last page.')
            return

        try:
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.trim_pdf(first_page, last_page)
            messagebox.showinfo('Success', 'PDF trimmed successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))

        window.destroy()
        

    def encrypt_pdf(self) -> None:
        encryption_algorithms: list[str] = ['AES-256', 'AES-256-R5', 'AES-128', 'RC4-128', 'RC4-40']
        pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
        if not pdf_filename: return

        new_window = ctk.CTkToplevel()
        new_window.title('Encrypt PDF')
        new_window.geometry('360x300')
        new_window.resizable(0, 0)

        algorithm_label = ctk.CTkLabel(new_window, text='Select Encryption Algorithm: ', text_color='#FEE715', fg_color='#101820')
        algorithm_label.grid(row=0, column=0, padx=5, pady=50, sticky='ew')

        selected_algorithm = ctk.StringVar(new_window)

        dropdown = ctk.CTkOptionMenu(new_window, variable=selected_algorithm, values=encryption_algorithms, dropdown_text_color='#FFF111', text_color='#FEE715', fg_color='#101820')
        dropdown.set('AES-256')
        dropdown.grid(row=0, column=1, padx=25, pady=50, sticky='ew')

        password_label = ctk.CTkLabel(new_window, text='Enter a Password: ', text_color='#FEE715', fg_color='#101820')
        password_label.grid(row=1, column=0, padx=5, pady=10, sticky='ew')

        password_entry = ctk.CTkEntry(new_window, text_color='#FEE715', fg_color='#101820')
        password_entry.grid(row=1, column=1, padx=25, pady=10, sticky='ew')

        encrypt_pdf_btn = ctk.CTkButton(new_window, text='Encrypt', command=lambda: self.encrypt_pdf_btn_clicked(pdf_filename, selected_algorithm.get(), password_entry.get(), new_window), text_color='#FEE715', fg_color='#101820')
        encrypt_pdf_btn.grid(row=2, column=0, columnspan=2, padx=5, pady=50, sticky='ew')


    def encrypt_pdf_btn_clicked(self, pdf_filename: str, selected_algorithm: ctk.StringVar, password: str, window: ctk.CTkToplevel) -> None:
        try:
            selected_algorithm = str(selected_algorithm)
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.encrypt_pdf(selected_algorithm, password)
            messagebox.showinfo('Success', 'PDF encrypted successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))

        window.destroy()


    def compress_pdf(self) -> None:
        try:
            pdf_filename: str = PdfManipulatorApp.select_filename(type='pdf', name='PDF')
            if not pdf_filename: return
            pdf: PdfManipulator = PdfManipulator(pdf_filename)
            pdf.compress_pdf()
            messagebox.showinfo('Success', 'PDF compressed successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def merge_pdfs(self) -> None:
        try:
            messagebox.showinfo('Select Files', 'Select PDF Files.')
            pdf_names: tuple[str] = ctk.filedialog.askopenfilenames(filetypes=[('PDF Files', '*.pdf')])
            merge_pdfs(pdf_names)
            messagebox.showinfo('Success', 'PDFs merged successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))


    def convert_word_to_pdf(self) -> None:
        try:
            word_filename: str = PdfManipulatorApp.select_filename(type='docx', name='Word')
            if not word_filename: return
            convert_word_to_pdf(word_filename)
            messagebox.showinfo('Success', 'Word converted to PDF successfully.')
        except Exception as e:
            messagebox.showerror('Error', str(e))
