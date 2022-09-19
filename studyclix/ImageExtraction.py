from pdfminer.layout import LAParams, LTTextBox, LTImage, LTFigure
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
import fitz # PyMuPDF
import io
import os
from PIL import Image
from collections import defaultdict
import glob, os
from docx2pdf import convert
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askdirectory
import json


def extract_text_images(file_path):
    """
    Return the list of text and images of all the document in a dictionary
    key: page number
    values: list of LTTextBox and LTFigure in the file
    """
    fp = open(file_path, 'rb')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pages = PDFPage.get_pages(fp)
    text_images = dict()
    i=0
    for page in pages:
        interpreter.process_page(page)
        layout = device.get_result()
        for lobj in layout:
            if isinstance(lobj, LTTextBox) or isinstance(lobj, LTFigure):
                try:
                    text_images[i].append((i,lobj))
                except:
                    text_images[i] = [(i,lobj)]
        i+=1
    fp.close()
    return text_images

def extract_images(filepath):
    """
    Return a dictionary with key image Name and PIL format
    """
    pdf_file = fitz.open(filepath)
    images = dict()
    for page in pdf_file:
        #image_list = page.getImageList()
        for image_index, img in enumerate(page.getImageList(), start=1):
            # get the XREF of the image
            xref = img[0]
            # extract the image bytes
            base_image = pdf_file.extractImage(xref)
            image_bytes = base_image["image"]
            # get the image extension
            #image_ext = base_image["ext"]
            # load it to PIL
            image = Image.open(io.BytesIO(image_bytes))
            images["Image"+str(xref)] = image
    pdf_file.close()
    return images

def assign_year_category(filepath):
    """
    Returns title and the mapping
    Assigns year and category to images
    """
    text_images = extract_text_images(filepath)
    out_all = sum(list(text_images.values()),[])
    out_all.sort(key=lambda x:(-x[0],x[1].y0), reverse=True)
    output = defaultdict(lambda: defaultdict(list))
    for i in range(len(out_all)):
        if isinstance(out_all[i][1], LTTextBox):
            title=out_all[i][1].get_text().replace('\n','').strip()
            if title != "":
                break
    index = 0
    i=0
    while i<len(out_all):
        element = out_all[i][1]
        if isinstance(element, LTTextBox):
            to_add_questions = []
            to_add_making_scheme = []
            verify_year = out_all[i][1].get_text().replace('\n','').strip()
            if verify_year.isnumeric():
                next_element_index = i+1
                while next_element_index<len(out_all):
                    next_element = out_all[next_element_index][1]
                    if isinstance(next_element,LTFigure):
                        to_add_questions.append(next_element.name)
#                     else:
#                         break
                    if isinstance(next_element,LTTextBox) and 'Marking' in next_element.get_text().replace('\n','').strip():
                        break
                    next_element_index += 1
                    i+=1
                while next_element_index<len(out_all):
                    next_element = out_all[next_element_index][1]
                    if isinstance(next_element,LTTextBox):
                        if next_element.get_text().replace('\n','').strip().isnumeric():
                            break
                    if isinstance(next_element,LTFigure):
                        to_add_making_scheme.append(next_element.name)
                    next_element_index += 1
                    i += 1
                output[verify_year]['Questions'].append(to_add_questions)
                output[verify_year]['Marking Scheme'].append(to_add_making_scheme)
        i+=1
    return title, output
letters = [""]+list("bcdefghijklmopqrstuvwxyz")
for i in range(1,len(letters)):
    letters[i] = '-'+letters[i]
def save_year(year, directory_questions, directory_marking_scheme, images, output, letters, directory_out):
    images_questions = [[images[element] for element in elements] for elements in output[year]['Questions']]
    images_marking_scheme = [[images[element] for element in elements] for elements in output[year]["Marking Scheme"]]
    if len(images_questions)>=1:
        for i in range(len(images_questions)):
            if images_questions[i]:
                pdf1_filename = directory_questions+'/'+ str(year)+letters[i]+'.pdf'
                images_questions[i][0].save(pdf1_filename, "PDF" ,resolution=100.0, save_all=True, append_images=images_questions[i][1:])

    if len(images_marking_scheme)>=1:
        for i in range(len(images_marking_scheme)):
            if images_marking_scheme[i]:
                pdf1_filename = directory_marking_scheme+'/'+ str(year)+letters[i]+'.pdf'
                images_marking_scheme[i][0].save(pdf1_filename, "PDF" ,resolution=100.0, save_all=True, append_images=list(images_marking_scheme[i][1:]))

def process_pdf_file(filepath, directory_out, starting_year, questions_save, answers_save):
    title, output = assign_year_category(filepath)
    title = title.replace('/','_')
    title = title.replace('.','')
    elements = ['<','>',':','"','/',"\\","|","?","*"," "]
    for element in elements:
        title = title.replace(element,'_')
    if "Server Error" in title:
        print(filepath +" is not good.")
        return
    # if not os.path.exists(directory_out+'/'+title+'/'):
    #     os.makedirs(directory_out+'/'+title)
    # if not os.path.exists(directory_out+'/'+title+'/'+"Questions/"):
    #     os.makedirs(directory_out+'/'+title+'/'+"Questions/")
    # if not os.path.exists(directory_out+'/'+title+'/'+'/'+"Marking Scheme/"):
    #     os.makedirs(directory_out+'/'+title+'/'+"Marking Scheme/")
    if not os.path.exists(questions_save + '/' + title):
        
        os.makedirs(questions_save + '/' + title)

    if not os.path.exists(answers_save + '/' + title):
        os.makedirs(answers_save + '/' + title)
    
    images = extract_images(filepath)
    
    for year in output:
        if int(year)>= int(starting_year):
            save_year(year, questions_save + '/' + title, answers_save + '/' + title, images, output, letters, directory_out)

    

def main():

    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    directory_path = askdirectory() # show an "Open" dialog box and return the path to the selected file
    
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    directory_out = askdirectory() # show an "Open" dialog box and return the path to the selected file
    files=[]
    for file in glob.glob(directory_path+"/*"+".docx"):
        files.append(file)
    info_list = directory_path.split("/")
    Exam = info_list[-3]
    Subject = info_list[-2]
    Level = info_list[-1]


    if not os.path.exists(directory_out+'/Questions'+'/'+ Exam +'/'+Subject+'/' + Level):
        os.makedirs(directory_out+'/Questions'+'/'+ Exam +'/' + Subject +'/' + Level)
    questions_save = directory_out+'/Questions'+'/'+ Exam +'/' + Subject +'/' + Level
    if not os.path.exists(directory_out+'/Answers'+'/'+ Exam +'/'+Subject +'/'+Level):
        os.makedirs(directory_out+'/Answers'+'/'+Exam + '/'+ Subject +'/' + Level)
    answers_save = directory_out+'/Answers'+'/'+ Exam +'/' + Subject +'/' + Level
    with open('settings.json', 'r') as fp:
        settings = json.load(fp)
    starting_year = settings['Starting_year']
    for file in files:
        try:
            os.remove(os.getcwd()+'\\converted_to_pdf.pdf')
        except:
            pass

        convert(file, os.getcwd()+'\\converted_to_pdf.pdf')
        process_pdf_file("converted_to_pdf.pdf", directory_out, starting_year, questions_save, answers_save)
    os.remove(os.getcwd()+'\\converted_to_pdf.pdf')

    

if __name__ == "__main__":
    main()
    
    
