import PyPDF2
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def addPageNumber(pdf_path,position):
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

        # finally, write "output" to a real file
        """filename="page"+index+'.pdf'
        outputStream = open(filename, "wb")
        output.write(outputStream)
        outputStream.close()"""
    with open("result.pdf","wb") as handler:
        output.write(handler)

pdfPath="C:\\Users\\p\\Desktop\\pdfMerge\\in\\Homework_3.pdf"
