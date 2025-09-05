from flask import Flask, request, send_file, render_template, send_from_directory
import pandas as pd
import os
import io
import zipfile

app = Flask(__name__, template_folder=os.path.dirname(os.path.abspath(__file__)))

# Serve CSS from the same directory
@app.route('/style.css')
def serve_css():
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), 'style.css')

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if not file:
            return "No file uploaded", 400

        df = pd.read_excel(file, engine="openpyxl")
        values_d = df.iloc[:, 3].tolist()
        values_e = df.iloc[:, 4].tolist()
        max_chunk_size = 199

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zipf:
            i = 0
            while i < len(values_d):
                current_e = values_e[i]
                chunk = []
                file_count = 1
                while i < len(values_d) and values_e[i] == current_e:
                    chunk.append(values_d[i])
                    i += 1
                    if len(chunk) == max_chunk_size:
                        zipf.writestr(f"{current_e}_{file_count}.txt", "\n".join(map(str, chunk)))
                        file_count += 1
                        chunk = []
                if chunk:
                    zipf.writestr(f"{current_e}_{file_count}.txt", "\n".join(map(str, chunk)))

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            download_name="output_txt.zip",
            as_attachment=True
        )

    return render_template("index.html")

if __name__ == "__main__":
    import webbrowser
    webbrowser.open("http://localhost:5000", new=0)
    app.run(host="0.0.0.0", port=5000)