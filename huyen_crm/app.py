from pickle import NONE
from flask import Flask, request, redirect, url_for , jsonify , send_file ,send_from_directory
from werkzeug.utils import secure_filename
from crm import *
import json
from fastapi import FastAPI

class DataModel:
    def __init__(self, result, message, item):
        self.result = result
        self.message = message
        self.item = item

class ErrorModel:
    def __init__(self, result, message,item):
        self.result = result
        self.message = message
        self.item = item

class ResponseModel:
    def __init__(self, data, error):
        self.data = data
        self.error = error

UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'docx','doc','pdf'}

app = Flask(__name__, static_url_path='/static')
app.config['upload_folder'] = UPLOAD_FOLDER 

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/upload_file', methods=['GET', 'POST'])
def upload_file():
    error = None
    data = None
    sess_id = request.args.get('sess_id')
    try:
        if request.method == 'POST':
            file = request.files['file']
            filename = secure_filename(file.filename)
            file_name, file_end = os.path.splitext(filename)
            path_id =  os.path.join(app.config['upload_folder'], sess_id)
            path_file = os.path.join(path_id, file_name)
            if not os.path.exists(path_file):
                os.makedirs(path_file)
            input_file = os.path.join(path_file, filename)
            file.save(input_file)
            img_org_base64 = start(input_file)
            item = {"sess_id": sess_id , "image" : img_org_base64, "path" : input_file}
            data = DataModel(True, " Xử lí file thành công ", item)
        if error is not None:
            error = vars(error)
        if data is not None:
            data = vars(data)
        response = ResponseModel(data, error)
    except:
        item = {"sess_id": sess_id}
        data = DataModel(False, " Chưa tải tệp tin! ", item)
        if error is not None:
            error = vars(error)
        if data is not None:
            data = vars(data)
        response = ResponseModel(data, error)
    return json.dumps(vars(response))

@app.route('/search', methods=['GET', 'POST'])
def search():
    error = None
    data = None
    sess_id = request.args.get('sess_id')
    input_file = request.args.get('input_file')
    if request.method == 'POST':
        key = request.form['text_change']
        try:
            img_base64, countKey = stage2(input_file,key)
        except:
            item = {"sess_id": sess_id, "input_file": input_file}
            data = DataModel(False, " Không tìm thấy tập tin! ", item)
            if key == "":
                data = DataModel(False, " Hãy nhập từ muốn thay thế! ", item)
            if error is not None:
                error = vars(error)
            if data is not None:
                data = vars(data)
            response = ResponseModel(data, error)
        else:
            if countKey == 0:
                item = {"sess_id": sess_id, "number_text": countKey}
                data = DataModel(False, " Không có chuỗi khớp ", item)
                if error is not None:
                    error = vars(error)
                if data is not None:
                    data = vars(data)
                response = ResponseModel(data, error)
            else:
                item = {"sess_id": sess_id, "number_text": countKey, "image":  img_base64}
                data = DataModel(True, " Ảnh trả về ", item)
                if error is not None:
                    error = vars(error)
                if data is not None:
                    data = vars(data)
                response = ResponseModel(data, error)
        return json.dumps(vars(response))

@app.route('/replace_file', methods=['GET', 'POST'])
def replace():
    error = None
    data = None
    sess_id = request.args.get('sess_id')
    input_file = request.args.get('input_file')
    contents = request.json
    count = 0
    dict_keys = []
    for content in contents:
        a_key = {}
        # count += 1
        key = content['name']
        value = content['replace_with']
        numberList = content['index']
        checkin = False
        for i in range(len(dict_keys)):
            if key in dict_keys[i]:
                dict_keys[i][key][value]=[num for num in numberList]
                checkin = True
        if not checkin:
            dict_keys.append({key:{value:numberList}})
    for dict_key in dict_keys:
        img_org_base64,input_file = stage3(input_file,dict_key,count)
    #key = contents[0]['name']
    #value = contents[0]['replace_with']
    #img_org_base64,output_file = stage3(input_file,key,value,numberList)
    item = {"sess_id" : sess_id , "img_org_base64": img_org_base64,"url": input_file}
    data = DataModel(True, "File thay đổi thành công ", item)
    if error is not None:
        error = vars(error)
    if data is not None:
        data = vars(data)
    response = ResponseModel(data, error)
    return json.dumps(vars(response))


@app.route('/view', methods=['GET', 'POST'])
def view():
    error = None
    data = None
    sess_id = request.args.get('sess_id')
    input_file = request.args.get('input_file')
    try:
        img_org_base64 = start(input_file)
        item = {"sess_id" : sess_id , "img_org_base64": img_org_base64,"url": input_file}
        data = DataModel(True, "Mở file thành công ", item)
    except:
        item = {"sess_id": sess_id, "input_file": input_file}
        data = DataModel(False, " Không tìm thấy tập tin! ", item)
    if error is not None:
        error = vars(error)
    if data is not None:
        data = vars(data)
    response = ResponseModel(data, error)
    return json.dumps(vars(response))

'''@app.route("/delete" ,methods=['GET', 'POST'])
def delete():
    error = None
    data = None
    path = './static/uploads'
    sess_id = request.args.get('sess_id')
    dirs = os.listdir(path)
    for folder in dirs:
        folder_id = folder.split('_')[-1]
        if sess_id == folder_id:
            path_del = os.path.join(path, folder)
            shutil.rmtree(path_del)
            item = {"sess_id": sess_id}
            data = DataModel(True, "Xóa file thành công ", item)
        else:
            item = {"sess_id": sess_id}
            error = ErrorModel(False, "File không tồn tại", item)
    if error is not None:
        error = vars(error)
    if data is not None:
        data = vars(data)
    response = ResponseModel(data, error)
    return json.dumps(vars(response))'''

@app.route("/static/<path:path>")
def static_dir(path):
    return send_from_directory("static", path)

if __name__ == "__main__":
    app.run (host='0.0.0.0', port=4000)