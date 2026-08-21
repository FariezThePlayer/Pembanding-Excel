import os
import shutil
import tempfile
import time
import uuid

from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.exceptions import HTTPException

from comparator import (
    load_workbook_safe,
    sheet_names,
    clean_ledger_by_formatting,
)

app = Flask(__name__)
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

# Semua file yang diunggah pengguna HANYA disimpan sementara di folder temp
# sistem operasi (bukan di dalam folder proyek) dan dihapus otomatis:
#   - setiap kali ada request baru, job yang sudah lebih tua dari
#     JOB_TTL_SECONDS akan disapu dan dihapus (lihat cleanup_expired_jobs), dan
#   - setiap kali job dipakai, "jam" job tersebut di-reset supaya sesi yang
#     masih aktif dipakai tidak keburu terhapus.
# Tidak ada file unggahan yang disimpan permanen.
JOBS_DIR = os.path.join(tempfile.gettempdir(), "excel_cleaner_jobs")
os.makedirs(JOBS_DIR, exist_ok=True)
JOB_TTL_SECONDS = 30 * 60  # 30 menit tidak dipakai -> dihapus otomatis


def cleanup_expired_jobs():
    """Hapus folder job (hasil unggahan sementara) yang sudah lebih tua dari JOB_TTL_SECONDS."""
    now = time.time()
    try:
        entries = os.listdir(JOBS_DIR)
    except OSError:
        return
    for entry in entries:
        full_path = os.path.join(JOBS_DIR, entry)
        try:
            age = now - os.path.getmtime(full_path)
        except OSError:
            continue
        if age > JOB_TTL_SECONDS:
            shutil.rmtree(full_path, ignore_errors=True)


def touch_job(job_id):
    """Perbarui waktu 'terakhir dipakai' sebuah job supaya tidak keburu dihapus otomatis."""
    job_dir = os.path.join(JOBS_DIR, job_id)
    try:
        os.utime(job_dir, None)
    except OSError:
        pass


def safe_preview_value(v):
    if v is None:
        return ""
    return v.isoformat() if hasattr(v, "isoformat") else str(v)


@app.after_request
def add_no_store_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.errorhandler(Exception)
def handle_any_exception(e):
    if isinstance(e, HTTPException):
        return jsonify({"error": e.description}), e.code
    return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/inspect", methods=["POST"])
def inspect():
    """Upload a workbook and return its sheet names so the user can pick one."""
    cleanup_expired_jobs()

    file = request.files.get("file")
    if not file or not getattr(file, "filename", ""):
        return jsonify({"error": "File Excel wajib diunggah."}), 400

    job_id = uuid.uuid4().hex
    job_dir = os.path.join(JOBS_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    source_path = os.path.join(job_dir, "source.xlsx")
    file.save(source_path)

    try:
        wb = load_workbook_safe(source_path)
        sheets = sheet_names(wb)
        touch_job(job_id)
        return jsonify({
            "job_id": job_id,
            "file_name": file.filename,
            "sheets": sheets,
        })
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": f"Gagal membaca file: {e}"}), 400


@app.route("/api/clean", methods=["POST"])
def clean():
    """Run the font-based ledger cleaner and return a preview + download link."""
    cleanup_expired_jobs()

    data = request.get_json(silent=True) or {}
    job_id = data.get("job_id")
    sheet = data.get("sheet")
    font_name = (data.get("font_name") or "Arial").strip() or "Arial"
    try:
        font_size = float(data.get("font_size", 9))
    except (TypeError, ValueError):
        font_size = 9.0

    if not job_id or not sheet:
        return jsonify({"error": "job_id dan sheet wajib diisi."}), 400

    job_dir = os.path.join(JOBS_DIR, job_id)
    source_path = os.path.join(job_dir, "source.xlsx")
    if not os.path.exists(source_path):
        return jsonify({"error": "Sesi sudah kedaluwarsa, silakan unggah ulang file."}), 400

    output_path = os.path.join(job_dir, "cleaned.xlsx")

    try:
        wb = load_workbook_safe(source_path)
        new_wb, kept_rows = clean_ledger_by_formatting(
            wb, sheet, output_path, font_name=font_name, font_size=font_size
        )

        new_ws = new_wb.active
        preview_rows = []
        headers = [safe_preview_value(c.value) for c in new_ws[1]]
        max_preview = min(new_ws.max_row, 21)  # header + 20 data rows
        for row in new_ws.iter_rows(min_row=2, max_row=max_preview):
            preview_rows.append([safe_preview_value(c.value) for c in row])

        touch_job(job_id)

        return jsonify({
            "job_id": job_id,
            "sheet": sheet,
            "font_name": font_name,
            "font_size": font_size,
            "total_rows": kept_rows,
            "headers": headers,
            "preview_rows": preview_rows,
            "preview_truncated": kept_rows > 20,
        })
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal membersihkan file: {e}"}), 500


@app.route("/api/download/<job_id>", methods=["GET"])
def download(job_id):
    job_dir = os.path.join(JOBS_DIR, job_id)
    output_path = os.path.join(job_dir, "cleaned.xlsx")
    if not os.path.exists(output_path):
        return jsonify({"error": "File hasil tidak ditemukan atau sudah kedaluwarsa."}), 404

    touch_job(job_id)
    return send_file(
        output_path,
        as_attachment=True,
        download_name="cleaned_ledger.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
