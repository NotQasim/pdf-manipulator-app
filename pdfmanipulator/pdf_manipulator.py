from pypdf import PdfReader, PdfWriter


class PdfManipulator:
    def __init__(self, filename_pdf: str) -> None:
        self.filename_pdf = filename_pdf
        self.reader = PdfReader(filename_pdf)
        self.writer = PdfWriter(clone_from=filename_pdf)


    def extract_text(self) -> None:
        try:
            with open(f'{self.filename_pdf.split('.pdf')[0]}_text.txt', 'w', encoding='utf-8', errors='ignore') as f:
                for page_number, page_object in enumerate(self.reader.pages):
                    f.write(f'\nPage {page_number+1}: {page_object.extract_text()}\n')
        except UnicodeEncodeError:
            pass
    

    def extract_images(self) -> None:
        for page_number, page_object in enumerate(self.reader.pages):
            for image_object in page_object.images:
                with open(str(page_number) + image_object.name, 'wb') as f:
                    f.write(image_object.data)


    def extract_metadata(self) -> None:
        raw_metadata = self.reader.metadata
        metadata = {
            'Title': raw_metadata.title,
            'Author': raw_metadata.author,
            'Creator': raw_metadata.creator,
            'Producer': raw_metadata.producer,
            'Creation Date': raw_metadata.creation_date.strftime("%Y-%m-%d %H:%M:%S %Z"),
            'Modification Date': raw_metadata.modification_date.strftime("%Y-%m-%d %H:%M:%S %Z")
        }
        
        with open(f'{self.filename_pdf.split('.pdf')[0]}_metadata.txt', 'w') as f:
            for key in metadata:
                f.write(f'\n{key}: {metadata[key]}\n')
    

    def add_stamp(self, stamp_filename: str) -> None:
        stamp = PdfReader(stamp_filename).pages[0]

        for page in self.writer.pages:
            page.merge_page(stamp, over=True)

        self.writer.write(f'{self.filename_pdf.split('.pdf')[0]}_stamped.pdf')


    def add_watermark(self, watermark_filename: str) -> None:
        watermark = PdfReader(watermark_filename).pages[0]

        for page in self.writer.pages:
            page.merge_page(watermark, over=False)

        self.writer.write(f'{self.filename_pdf.split('.pdf')[0]}_watermarked.pdf')


    def trim_pdf(self, first_page: int, last_page: int) -> None:
        self.writer = PdfWriter()
        self.writer.append(self.filename_pdf, (first_page-1, last_page))

        with open(f'{self.filename_pdf.split('.pdf')[0]}_trimmed.pdf', 'wb') as f:
            self.writer.write(f)


    def encrypt_pdf(self, encryption_alg: str, password: str) -> None:
        self.writer.encrypt(password, algorithm=encryption_alg)

        with open(f'{self.filename_pdf.split('.pdf')[0]}_encrypted.pdf', 'wb') as f:
            self.writer.write(f)


    def compress_pdf(self) -> None:
        for page in self.writer.pages:
            page.compress_content_streams()
        
        with open(f'{self.filename_pdf.split('.pdf')[0]}_compressed.pdf', 'wb') as f:
            self.writer.write(f)
