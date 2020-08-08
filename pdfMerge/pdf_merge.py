import PyPDF2
import fitz
import io
from PIL import Image
import os
import random
import string
import pikepdf
from tqdm import tqdm
from zipfile import ZipFile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pathlib import Path


"""
Converts entire PDF to set of disjoint
PNG images and returns as .zip 
"""
def toPNG(pdf_path):
    pdffile = pdf_path
    doc = fitz.open(pdffile)
    N = 12
    res = ''.join(random.choices(string.ascii_uppercase +
                                 string.digits, k=N))
    zipname=res+'.zip'
    with ZipFile(zipname, 'w') as zipper :
        for i in range(len(doc)):
            page = doc.loadPage(i)  # number of page
            # to keep pdf quality preserved for image
            zoom = 2  # zoom factor
            mat = fitz.Matrix(zoom, zoom)
            pix = page.getPixmap(matrix=mat)

            # saving
            index=i+1
            tmp='page'+str(index)
            filename=tmp+'.png'
            pix.writePNG(filename)
            zipper.write(filename)
            os.remove(filename)


"""
Converts PDF to JPG images and returns as .zip 
"""
def toJPG(pdf_path):
    pdffile = pdf_path
    doc = fitz.open(pdffile)
    N = 12
    res = ''.join(random.choices(string.ascii_uppercase +
                                 string.digits, k=N))
    zipname=res+'.zip'
    with ZipFile(zipname, 'w') as zipper :
        for i in range(len(doc)):
            page = doc.loadPage(i)  # number of page
            # to keep pdf quality preserved for image
            zoom = 2  # zoom factor
            mat = fitz.Matrix(zoom, zoom)
            pix = page.getPixmap(matrix=mat)

            # saving
            index=i+1
            tmp='page'+str(index)
            filename=tmp+'.jpg'
            pix.writePNG(filename)
            zipper.write(filename)
            os.remove(filename)

"""
Converts singular image to a PDF
"""
def toPDF(image,path):
    x=image.split('.')
    filename=str(x[0])+'.pdf'
    tmp=Image.open(path+"\\"+image,mode='r')
    tmp1=tmp.convert('RGB')
    tmp1.save(path+'\\'+filename)

"""
Merges all PDFs given, if the given file is not a 
PDF, but an image, then it is converted to PDF
"""
def merge(path,filename):
    filename+='.pdf'
    if os.path.exists(path):
        for i in os.listdir(path):
            if not i.endswith('pdf'):
                toPDF(i,path)
        x=[i for i in os.listdir(path) if i.endswith('.pdf')]
        print(x)
        merger=PyPDF2.PdfFileMerger()
        for i in x:
            merger.append(open(path+'\\'+i,'rb'))
        with open("out\\"+filename,"wb") as handler:
            merger.write(handler)
    else:
        print("Invalid Path")

"""
 Parses all pages from PDF and returns a .zip of all individual files
"""
def parseAll(pdf_path):
    inputpdf = PyPDF2.PdfFileReader(open(pdf_path, "rb"))
    N = 12
    res = ''.join(random.choices(string.ascii_uppercase +string.digits, k=N))
    zipname=res+'.zip'
    with ZipFile(zipname, 'w') as zipper :
        for i in range(inputpdf.numPages) :
            output = PyPDF2.PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            index=str(i+1)
            filename="page"+index+'.pdf'
            with open(filename, "wb") as stream :
                output.write(stream)
            zipper.write(filename)
            os.remove(filename)


"""
Extracts specified PDF pages (given in 2nd argument as a list)
And returns a zip containing the specified pages
"""
def extract(pdf_path,lst):
    inputpdf = PyPDF2.PdfFileReader(open(pdf_path, "rb"))
    N = 12
    res = ''.join(random.choices(string.ascii_uppercase + string.digits, k=N))
    zipname = res + '.zip'
    with ZipFile(zipname,"w") as zipper:
        for i in range(inputpdf.numPages):
            if(i in lst):
                output=PyPDF2.PdfFileWriter()
                output.addPage(inputpdf.getPage(i))
                index=str(i+1)
                filename="page"+index+'.pdf'
                with open(filename,"wb") as handler:
                    output.write(handler)
                zipper.write(filename)
                os.remove(filename)

"""
Deletes specified pages in lst and returns a single
PDF excluding given pages
"""
def delete(pdf_path,lst):
    inputpdf = PyPDF2.PdfFileReader(open(pdf_path, "rb"))
    output = PyPDF2.PdfFileWriter()
    for i in range(inputpdf.getNumPages()) :
        if i not in lst :
            p = inputpdf.getPage(i)
            output.addPage(p)
    with open('newfile.pdf', 'wb') as f :
        output.write(f)


"""
Rearranges pages in PDF and merges them back together
by creating a new PDF and appending to it in the order 
specified in lst (same length is assumed between
lst and number of pages in pdf_path)
"""
def rearrange(pdf_path,lst):
    inputpdf=PyPDF2.PdfFileReader(open(pdf_path,"rb"))
    output=PyPDF2.PdfFileWriter()
    for i in lst:
        index=i-1    # to account for +1 offset
        p=inputpdf.getPage(index)
        output.addPage(p)
    with open("rearranged.pdf","wb") as handler:
        output.write(handler)


"""
Adds page numbers to given PDF file and returns this new file. 
There are 6 positions to choose from: BC, BL, BR, TC, TL TR 
which essentially correspond to bottom-center, bottom-left, 
bottom-right, top-center, etc...
"""
def addPageNumber(pdf_path,position):
    filename=Path(pdf_path).resolve().stem+"-numbered.pdf"
    if (position == "BC") :
        x = 300
        y = 20
    elif (position == "BL") :
        x = 30
        y = 20
    elif (position == "BR") :
        x = 580
        y = 20
    elif (position == "TL") :
        x = 30
        y = 760
    elif (position == "TC") :
        x = 300
        y = 760
    else :              # TR
        x = 580
        y = 760
    output = PyPDF2.PdfFileWriter()
    existing_pdf = PyPDF2.PdfFileReader(open(pdf_path, "rb"))
    for i in range(existing_pdf.getNumPages()):
        packet = io.BytesIO()
        index=str(i+1)

        # create a new PDF with Reportlab
        can = canvas.Canvas(packet, pagesize=letter)
        if position[0] == 'B' :
           can.drawCentredString(x, y, index)
        else :
            can.drawCentredString(x, y, index)
        can.save()

        #move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PyPDF2.PdfFileReader(packet)


        # add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.getPage(i)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)

    # save the resulting PDF
    with open(filename,"wb") as handler:
        output.write(handler)


"""
Rotates all pages in a given PDF by "deg" degrees
in the specified direction given ("CW or CCW")
"""
def rotateAll(pdf_path,deg,direction):
    pdf_writer = PyPDF2.PdfFileWriter()
    pdf_reader = PyPDF2.PdfFileReader(pdf_path)
    filename = Path(pdf_path).resolve().stem + "-rotated.pdf"
    for i in range(pdf_reader.getNumPages()):
        try:
            # Rotate page
            if direction=="CW":
                tmp = pdf_reader.getPage(i).rotateClockwise(deg)
            elif direction=="CCW":
                tmp = pdf_reader.getPage(i).rotateCounterClockwise(deg)
            else:
                return -1
            pdf_writer.addPage(tmp)
        except:
            print("Please make sure deg is multiple of 90")
    with open(filename, 'wb') as handler:
        pdf_writer.write(handler)

"""
Rotates a given page CW or CCW and returns a
newly generated PDF file after modifying with the 
given config parameters
"""
def rotatePage(pdf_path,page,deg,direction):
    result_pdf = PyPDF2.PdfFileWriter()
    pdf = PyPDF2.PdfFileReader(pdf_path)
    filename = Path(pdf_path).resolve().stem + "-rotate.pdf"
    for i in range(pdf.getNumPages()):
        if(page==i+1):
            if direction=="CW":
                tmp=pdf.getPage(i).rotateClockwise(deg)
            else:
                tmp=pdf.getPage(i).rotateCounterClockwise(deg)
            result_pdf.addPage(tmp)
        else:
            result_pdf.addPage(pdf.getPage(i))
    with open(filename,"wb") as handler:
        result_pdf.write(handler)


"""
Rotates multiple pages in a PDF by varying degree and
direction. Input PDF path is supplied and a list of 
dictionaries is given for each query (page,deg,direction)
"""
def rotatePages(pdf_path,lst):
    result_pdf = PyPDF2.PdfFileWriter()
    pdf = PyPDF2.PdfFileReader(pdf_path)
    filename = Path(pdf_path).resolve().stem + "-rotation.pdf"
    pages=[]
    for i in lst:
        pages.append(i['page'])
    for i in range(pdf.getNumPages()) :
        if i in pages:
            page=i+1
            direction=lst[i-1]["direction"]
            deg=lst[i-1]["deg"]
            if (page == i + 1) :
                if direction == "CW" :
                    tmp = pdf.getPage(i).rotateClockwise(deg)
                else :
                    tmp = pdf.getPage(i).rotateCounterClockwise(deg)
                result_pdf.addPage(tmp)
        else :
            result_pdf.addPage(pdf.getPage(i))
    with open(filename, "wb") as handler :
        result_pdf.write(handler)


"""
Encrypts a PDF file using the supplied password 
"password" by the user. To decrypt, users will have
to enter the password you provided at the time
of encryption
"""
def encryptPDF(pdf_path,password):
    pdf_writer=PyPDF2.PdfFileWriter()
    pdf_reader=PyPDF2.PdfFileReader(pdf_path)
    filename = Path(pdf_path).resolve().stem + "-encrypted.pdf"
    for page in range(pdf_reader.getNumPages()):
        pdf_writer.addPage(pdf_reader.getPage(page))
    pdf_writer.encrypt(user_pwd=password,owner_pwd=None,use_128bit=True)
    with open(filename,"wb") as handler:
        pdf_writer.write(handler)


"""
Decrypts (or at least attempts to) PDF
and saves decrypted file according to the output
path (i.e. result.pdf). Standard rockyou.txt is used
as the wordlist of choice but one can certainly
get creative and customize it accordingly 
"""
def decrypt_pdf(input_path, output_path):
    # MUST HAVE ROCKYOU.TXT ALREADY DOWNLOADED
    passwords = [line.strip() for line in open("rockyou.txt",encoding="latin-1")]
    # iterate over passwords
    for pwd in tqdm(passwords, "Decrypting PDF") :
        try :
            # open PDF file
            with pikepdf.open(input_path, password=pwd) as pdf :
                # Password decrypted successfully, break out of the loop
                print("[+] Password found:", pwd)
                with open(input_path, 'rb') as input_file, \
                        open(output_path, 'wb') as output_file :
                    reader = PyPDF2.PdfFileReader(input_file)
                    reader.decrypt(pwd)
                    writer = PyPDF2.PdfFileWriter()
                    for i in range(reader.getNumPages()) :
                        writer.addPage(reader.getPage(i))
                    writer.write(output_file)
                break
        except pikepdf._qpdf.PasswordError as e :
            # wrong password, just continue in the loop
            continue
    print("No passwords were found :(")

"""
Signs given PDF file at the page specified.
For now, signature must be saved as image and in
the working directory, but for the web app
we will be able to dynamically get the coordinates
by drawing a bow to where we want to place it
"""
def sign(pdf_path,image_path,page,x1,y1,x2,y2):
    filename=Path(pdf_path).resolve().stem+"-signed.pdf"
    try:
        image_rect=fitz.Rect(x1,y1,x2,y2)
    except:
        print('bad coordinates')
    try:
        file_handle=fitz.open(pdf_path)
        pg=file_handle[page-1]
        # draw rectangle where image shall be placed
        pg.insertImage(image_rect,filename=image_path)
        file_handle.save(filename)
    except:
        print('invalid page range')


if __name__=="__main__":
    path="in"
    file="result"
    #merge(path,file)
    pdfPath="<your path here>"
    sign_path="<your path here>"
    #parseAll(pdfPath)
    #extract(pdfPath,[0,3,4])
    #delete(pdfPath,[0])
    #toJPG(pdfPath)
    #rearrange(pdfPath,[3,1,2])
    #addPageNumber(pdfPath,"BC")
    #rotateAll(pdfPath,360,"CCW")
    #lst=[{"page":0,"direction":"CW","deg":90},{"page":2,"direction":"CCW","deg":270}]
    #rotatePage(pdfPath,3,90,"CCW")
    #rotatePages(pdfPath,lst)
    #encryptPDF(pdfPath,"twofish")
    decrypt_pdf("path for file to encrypt",'<output>.pdf')

    #sign(pdfPath,sign_path,1,450,50,550,120)




