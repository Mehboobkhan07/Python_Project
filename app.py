from flask import Flask, send_file
from src.generate_doc import doc

app = Flask(__name__)

@app.route("/")
def generate():
    return send_file(
        "output/form_A_corrected.docx",
        as_attachment=True
    )

if __name__ == "__main__":
    app.run()
