from email.mime import base
from docx import Document
from docx2pdf import convert
from subprocess import Popen
from PyPDF2 import PdfFileReader as read
from pdf2image import convert_from_path
import numpy as np
import base64
import cv2
import os

def input_file_processing(input_file):
##    input: file docx cần xử lý
##    output: file pdf/doc/docx, lưu trong folder chứa input_file
##    cần chuyển docx sang pdf
    input_file_name = os.path.splitext(input_file)[0]
    out_folder = input_file_name.split("/")[:-1]
    out_folder = "/".join(out_folder)
    # convert trên window
    #convert(input_file)
    
    # convert trên linux
    LIBRE_OFFICE = r"/usr/bin/lowriter"
    p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_file])
    #print([LIBRE_OFFICE, '--convert-to', 'pdf', input_file])
    p.communicate()
    
    input_pdf = input_file_name + '.pdf'
    return input_pdf, input_file

def pdf_to_img(input_pdf):
    pdf = read(open(input_pdf,'rb'))
    #pdf = PdfFileReader(open(input_pdf,'rb'))
    number_page = pdf.getNumPages()
    folder = input_pdf.split('.pd')[0] + '_img'
    if not os.path.exists(folder):
       os.makedirs(folder)
    pages = convert_from_path(input_pdf, 200)
    counter = 0
    for page in pages:
        myfile = folder+'/'+str(counter)
        page.save(myfile, 'PNG')
        counter = counter+1
        if counter > number_page:
            break
    return folder,number_page
        
def imageToBase64(image): 
##    imencode mã hóa định dạng img thành data truyền trực tuyến và gán nó vào bộ nhớ đệm
##    nén định dạng dữ liệu hình ảnh để tạo điều kiện cho việc truyền mạng
##    b64encode mã hóa hình ảnh thành nhị phân hiển thị bằng ASCII
##    decode chuyển dạng mã hóa về string
    image = cv2.imread(image)
    retval, buffer = cv2.imencode('.jpg', image)
    jpg_as_text = base64.b64encode(buffer)
    image_data = jpg_as_text.decode("utf-8")
    image_data = str(image_data)
    #print(image_data)
    image_data = 'data:image/png;base64,' + image_data
    return image_data

def input_processing(input_file):
##    input: đường dẫn đến file docx cần xử lý
##    output: file đoc/docx, pdf, folder chứa images
##    docx - pdf - folder ảnh - chuyển ảnh thành data truyền mạng
    input_pdf, output_file = input_file_processing(input_file)
    input_image,number_page = pdf_to_img(input_pdf)
    img_org_base64 = []
    for file in range(number_page):
        image = input_image + '/' + str(file)
        img_org_base64.append(imageToBase64(image))
    return output_file,img_org_base64

def search_processing(input_file):
##    với ảnh có màu vàng + đỏ trong bước tìm kiếm
##    input: đường dẫn đến file docx cần xử lý
##    output: file đoc/docx, pdf, folder chứa images
##    docx - pdf - folder ảnh - chuyển ảnh thành data truyền mạng
    input_pdf, output_file = input_file_processing(input_file)
    input_image,number_page = pdf_to_img(input_pdf)
    img_org_base64 = []
    for file in range(number_page):
        image = input_image + '/' + str(file)
        img_org_base64.append(imageToBase64(image))
    colored_file = []
    for i in range (len(img_org_base64)):
        encoded_data = img_org_base64[i].split(',')[1]
        encoded_data += "="*((4 - len(encoded_data) % 4) % 4) 
        encoded_data = base64.urlsafe_b64decode(encoded_data)
        
        nparr = np.fromstring(encoded_data,np.uint8)
        image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)                

        # gán khoảng giá trị cho màu vàng
        lower_ye = np.array([22, 93, 0], dtype="uint8")
        upper_ye = np.array([45, 255, 255], dtype="uint8")

        # gán khoảng giá trị cho màu đỏ
        lower_red = np.array([160,100,20], dtype="uint8")
        upper_red = np.array([179,255,255], dtype="uint8")
                            
        # đánh dấu những ảnh nằm trong dải màu ở trên
        ye_mask = cv2.inRange(image, lower_ye, upper_ye)
        red_mask = cv2.inRange(image, lower_red, upper_red)

        pixels_1 = cv2.countNonZero(ye_mask)
        pixels_2 = cv2.countNonZero(red_mask)

        if pixels_1 > 0 and pixels_2 > 0:
            colored_file.append(img_org_base64[i])

    return output_file,colored_file
