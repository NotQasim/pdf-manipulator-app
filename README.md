<h1 align="center">PDF Manipulator App</h1>

<h2>Introduction</h2>
<p>This is a GUI Application written in Python, intended for Windows users that can manipulate PDF files in various ways, including text extraction, image extraction, encryption, and conversion of DOCX files to PDF files.</p>
<br>

<h2>Features</h2>
<ul>
  <li>Extract text from PDF</li>
  <li>Extract images from PDF</li>
  <li>Extract metadata from PDF</li>
  <li>Add stamp to PDF</li>
  <li>Add watermark to PDF</li>
  <li>Trim PDF</li>
  <li>Encrypt PDF</li>
  <li>Compress PDF</li>
  <li>Merge PDFs</li>
  <li>Convert Word to PDF</li>
</ul>
<br>

<h2>Explanation</h2>
<ol type="1">
  <li>
    <h4>Extract Text from PDF</h4>
    <p>Used to extract all text from a PDF and store it in a TXT file that is created in the same directory as the PDF. May fail sometimes in properly reading the text.</p>
  </li>

  <li>
    <h4>Extract Images from PDF</h4>
    <p>Used to extract all images from a PDF and store them as JPGs and/or PNGs in the same directory as the program.</p>
  </li>

  <li>
    <h4>Extract Metadata from PDF</h4>
    <p>Used to extract information about the title, creator, author, creation date, and modification date of a PDF and store it in a TXT file in the same directory as the PDF.</p>
  </li>

  <li>
    <h4>Add Stamp to PDF</h4>
    <p>Add a solid stamp over each page of the PDF, creating a new PDF in the same directory as the PDF. The stamp must be a PDF file, with a transparent background.</p>
  </li>

  <li>
    <h4>Add Watermark to PDF</h4>
    <p>Add a non-solid watermark over each page of the PDF, creating a new PDF in the same directory as the PDF. The watermark must be a PDF file, with a transparent background.</p>
  </li>

  <li>
    <h4>Trim PDF</h4>
    <p>Trim a PDF by entering the number of the first and last page to keep, creating a new PDF in the same directory as the original PDF.</p>
  </li>

  <li>
    <h4>Encrypt PDF</h4>
    <p>Encrypt a PDF by choosing an encryption algorithm (AES-256 is recommended) and entering a password, creating a new PDF in the same directory as the original PDF.</p>
  </li>

  <li>
    <h4>Compress PDF</h4>
    <p>Reduce the file size of a PDF by compressing the contents of each page, creating a new PDF in the same directory as the original PDF. Encryption may not always work as intended, and the size of the "compressed" may actually be greater than the original PDF.</p>
  </li>

  <li>
    <h4>Merge PDFs</h4>
    <p>Merge multiple selected PDF files, joining the last page of the first to the first page of the second, and so on. A new PDF will be created in the same directory as the program. The order of the PDFs when merging will be the order that appears in the file selection dialog box (alphabetical order).</p>
  </li>

  <li>
    <h4>Convert Word to PDF</h4>
    <p>Convert a DOC or DOCX file to a PDF, creating the PDF in the same directory as the Word file.</p>
  </li>
</ol>
