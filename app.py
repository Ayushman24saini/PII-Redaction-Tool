from flask import Flask, request, send_file, render_template_string
from pathlib import Path
import tempfile
import os

import pii_redactor


app = Flask(__name__)


HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>PII Redaction Tool</title>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 700px;
            margin: 60px auto;
            padding: 20px;
        }

        h1 {
            margin-bottom: 10px;
        }

        .box {
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 25px;
            margin-top: 25px;
        }

        input {
            margin: 15px 0;
        }

        button {
            padding: 10px 18px;
            cursor: pointer;
        }

        .error {
            color: #b00020;
            margin-top: 20px;
        }
    </style>
</head>

<body>

    <h1>PII Redaction Tool</h1>

    <p>
        Upload a Microsoft Word (.docx) document to detect and
        replace personally identifiable information.
    </p>

    <div class="box">

        <form method="POST" enctype="multipart/form-data">

            <input
                type="file"
                name="document"
                accept=".docx"
                required
            >

            <br>

            <button type="submit">
                Redact Document
            </button>

        </form>

        {% if error %}
            <p class="error">
                <strong>Error:</strong> {{ error }}
            </p>
        {% endif %}

    </div>

</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "GET":
        return render_template_string(HTML)

    uploaded_file = request.files.get("document")

    if not uploaded_file or not uploaded_file.filename:
        return render_template_string(
            HTML,
            error="Please select a DOCX file."
        )

    if not uploaded_file.filename.lower().endswith(".docx"):
        return render_template_string(
            HTML,
            error="Only .docx files are supported."
        )

    with tempfile.TemporaryDirectory() as temp_dir:

        input_path = Path(temp_dir) / "input.docx"
        output_path = Path(temp_dir) / "redacted.docx"

        uploaded_file.save(input_path)

        original_input = pii_redactor.INPUT_FILE
        original_output = pii_redactor.OUTPUT_FILE

        try:

            pii_redactor.INPUT_FILE = input_path
            pii_redactor.OUTPUT_FILE = output_path

            pii_redactor.process_document()

            if not output_path.exists():

                return render_template_string(
                    HTML,
                    error="Redaction completed but no output file was generated."
                )

            return send_file(
                output_path,
                as_attachment=True,
                download_name="REDACTED_DOCUMENT.docx",
                mimetype=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            )

        except Exception as exc:

            return render_template_string(
                HTML,
                error=f"Redaction failed: {exc}"
            )

        finally:

            pii_redactor.INPUT_FILE = original_input
            pii_redactor.OUTPUT_FILE = original_output


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5000))

    app.run(
        host="0.0.0.0",
        port=port
    )