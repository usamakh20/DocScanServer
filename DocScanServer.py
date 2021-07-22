from flask import Flask, send_file, request
import ImageToDoc

app = Flask(__name__)


@app.route('/', methods=['POST'])
def recieveImageAndSendDoc():
    file = request.files['Document']
    ImageToDoc.process(file)
    return send_file("demo.docx", attachment_filename="demo.docx")


@app.route('/test')
def test():
    return "Hello World"


if __name__ == '__main__':
    app.run(host='0.0.0.0')
