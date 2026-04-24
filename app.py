# -*- coding: utf-8 -*-
from flask import Flask, request, send_file, render_template, jsonify
import io, os, traceback
from converter import convert

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def do_convert():
    try:
        companies = []
        idx = 0
        while True:
            key = f"company_{idx}"
            if key not in request.form:
                break

            fmt        = request.form.get(f"fmt_{idx}", "excel")
            company    = request.form.get(f"company_{idx}", "").strip()
            sheet_name = request.form.get(f"sheet_{idx}", company[:10]).strip()
            currency   = request.form.get(f"currency_{idx}", "USD").strip().upper()
            eoy_rate   = float(request.form.get(f"eoy_{idx}", 0))
            avg_rate   = float(request.form.get(f"avg_{idx}", 0))
            prior_re   = float(request.form.get(f"prior_{idx}", 0) or 0)

            if fmt == "excel":
                f = request.files.get(f"excel_{idx}")
                if not f: break
                bs_stream = io.BytesIO(f.read())
                pl_stream = None
            else:
                pdf_f = request.files.get(f"pdf_{idx}")
                if not pdf_f: break
                bs_stream = io.BytesIO(pdf_f.read())
                pl_stream = None  # converter가 단일 PDF에서 BS/PL 자동 분류

            companies.append({
                "fmt": fmt, "company": company, "sheet_name": sheet_name,
                "currency": currency, "eoy_rate": eoy_rate,
                "avg_rate": avg_rate, "prior_re": prior_re,
                "bs_stream": bs_stream, "pl_stream": pl_stream,
            })
            idx += 1

        if not companies:
            return jsonify({"error": "파일이 없습니다."}), 400

        xlsx_bytes = convert(companies)
        return send_file(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="해외현지법인재무제표_변환.xlsx",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
