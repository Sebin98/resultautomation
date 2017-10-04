from appJar import gui
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
from bs4 import BeautifulSoup
import os
import webbrowser

global failed
failed=[]
global fail
fail=[]
sublist=['APPLIED ELECTRONICS & INSTRUMENTATION ENGINEERING','ELECTRONICS & COMMUNICATION ENGG','COMPUTER SCIENCE & ENGINEERING','ELECTRICAL AND ELECTRONICS ENGINEERING','INFORMATION TECHNOLOGY','MECHANICAL ENGINEERING','CIVIL ENGINEERING']

def updater(lo):
    pass
def setting(q):
    setter=gui("Settings")
    setter.setGeometry("700x400")
    setter.addLabel("t2","MASTER SETTINGS")
    setter.addFlashLabel("t3","Caution : Dont change unless required. Any unexpected change can result in invalid results!")
    setter.addLabelEntry("Collage Code")
    setter.setEntry("Collage Code","RET")
    setter.addLabel("t4","Branches   (branch names must be same as in the pdf file)")
    setter.addEntry("branch1")
    setter.addEntry("branch2")
    setter.addEntry("branch3")
    setter.addEntry("branch4")
    setter.addEntry("branch5")
    setter.addEntry("branch6")
    setter.addEntry("branch7")
    setter.setEntry("branch1",sublist[0])
    setter.setEntry("branch2",sublist[1])
    setter.setEntry("branch3",sublist[2])
    setter.setEntry("branch4",sublist[3])
    setter.setEntry("branch5",sublist[4])
    setter.setEntry("branch6",sublist[5])
    setter.setEntry("branch7",sublist[6])
    setter.addButton("Update",updater)
    setter.go()

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = file(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text

def loader(temp):
    try:
        global a
        a=loading.openBox(title="Select the marklist",dirName="D:")
        if a[-1:]!="f":
            loading.errorBox("WRONG INPUT","Please select file in the pdf format !")
        else:
            f=open("file.txt",'w')
            f.write(convert_pdf_to_txt(a))
            f.close()
            contents = open("file.txt","r")
            with open("htmlfile.html", "w") as e:
                e.write("<div class="+"temp"+">")   
                for lines in contents.readlines():
                        line=lines.strip()
                        if line in sublist:
                         e.write("</div>\n<div class=\""+line+"\">")
                        elif "RET" in line:
                            e.write("<span class=\"id\">"+line+"</span><br>\n")
                        else:  
                            e.write(line+"<br>\n")
                e.write("</div>\n")
            f.close()
            soup = BeautifulSoup(open("htmlfile.html"),"html.parser")
            for sub in sublist:
                detailsdata=soup.find_all('div',{'class':sub})
                p=detailsdata[0]
                f2=open(sub+".html",'w')
                f2.write(p.text)
                f2.close()
            global marks

            
            loading.stop()
            app.go()
    except:
         loading.errorBox("Loading Failed","Loder malfunction. make sure the file is a valid one")



def toolbar(btn):
    if btn == "EXIT": app.stop()
    elif btn == "Full Screen":
        if app.exitFullscreen():
            app.setGeometry("1000x1000")
        else:
            app.setGeometry("fullscreen")
    elif btn == "About":
        pass


def generator(btn):
    opted=app.getOptionBox("select the option : ")
    fail=[]    
    for i in sublist:
                count=0
                f=open(i+".html",'r')
                count=0
                global reg
                reg=[]
                mar=[]
                marks=[]
                text=[]
                for lines in f.readlines():
                            line=lines.strip()
                            if "RET" in line:
                                reg.append(line)
                            elif "(" in line:
                                mar.append(line)
                mar=mar[1:]
                mar = filter(None, mar)
                try :    
                    for i in range(0,len(mar)-1):
                        n=mar[i]
                        if n[-1]==',':
                            mar[i]=mar[i]+mar[i+1]
                            del mar[i+1]
                except :
                    pass
                
                for i in mar:
                    text=[x.strip() for x in i.split(',')]
                    #text=text+["NULL" for k in range(0,9-len(text))]   
                    marks.append(text[:])    
                for i in range(0,len(reg)):
                        count=0
                        for p in range(0,len(marks[i])):
                            if opted in marks[i][p]:
                                if count==0:
                                    fail.append(reg[i])
                                    fail.append(marks[i][p])
                                    count=count+1
                                else:
                                    fail.append(marks[i][p])
    
    j=open(opted+".html","w")
    j.write("<!DOCTYPE html><html><head><style>table {font-family: arial, sans-serif;border-collapse: collapse;width: 100%;}td, th {border: 1px solid #dddddd;text-align: left;padding: 8px;}tr:nth-child(even) {background-color: #dddddd;}</style></head><body>")
    j.write("<table align=\"center\"><tr><th align=\"center\" valign=\"middle\">Register Number</th><th colspan=\"6\"align=\"center\" valign=\"middle\">Subjects</th>")
    
    for i in fail:
        if "RET" in i:
            j.write("</tr>")
            j.write("<tr>")
            j.write("<td align=\"center\" valign=\"middle\">"+i+"</td>")
        else:
            j.write("<td align=\"center\" valign=\"middle\">"+i+"</td>")
    
    j.write("</table>")
    j.close()
    webbrowser.open(opted+".html")
    

app=gui("Generator")
app.setGeom ("600x600")
app.addToolbar(["About", "Full Screen","EXIT"], toolbar, findIcon=True)
app.setBg ("lightblue")
app.startTabbedFrame("VIEW")
app.startTab("GENERATOR")
app.addLabelOptionBox("select the option : ", ["(O)", "(A+)","(A)", "(B+)","(B)", "(C)","(P)", "(F)","(FE)","(Absent)"])
app.addButton("Generate&View",generator)

app.stopTab()

app.stopTabbedFrame()


loading=gui("KTU-Result Analyser")
loading.setBg ("grey")
loading.setGeom ( "300x200" ) 
loading.addLabel("title","KTU RESULT LIST GENERATOR")
loading.addButton("SETTINGS",setting)
loading.addLabel("title2","Please select the PDF file")
loading.addButton("LOAD",loader)

loading.go()
