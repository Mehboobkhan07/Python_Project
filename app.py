from flask import Flask, send_file
from src.generate_doc import create_doc

app = Flask(__name__)

@app.route("/")
def home():
    return "App is running successfully"

@app.route("/generate")
def generate():
    file_path = create_doc()
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run()
