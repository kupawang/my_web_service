#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
상표 OA 분석 웹앱
브라우저에서 OA ZIP 파일을 업로드하면 엑셀로 변환해 다운로드할 수 있게 해주는
로컬 실행용 Flask 서버.

실행:
    python app.py
그 다음 브라우저에서 http://localhost:5000 접속
"""

import os
import uuid
import shutil
import threading
import traceback

from flask import Flask, request, jsonify, send_file, render_template

import core

app = Flask(__name__)

WORK_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "work")
os.makedirs(WORK_DIR, exist_ok=True)

# 작업 상태를 메모리에 보관 (로컬 1인 사용 기준의 단순한 구조)
JOBS = {}


def _process_job(job_id: str, zip_path: str, output_path: str, api_key: str, count: int):
    def progress_cb(done, total, msg):
        JOBS[job_id]["done"] = done
        JOBS[job_id]["total"] = total
        JOBS[job_id]["message"] = msg

    try:
        core.run_job(zip_path, output_path, api_key, max_count=count, progress_cb=progress_cb)
        JOBS[job_id]["status"] = "done"
        JOBS[job_id]["message"] = "완료"
    except Exception as e:
        traceback.print_exc()
        JOBS[job_id]["status"] = "error"
        JOBS[job_id]["message"] = str(e)
    finally:
        if os.path.exists(zip_path):
            os.remove(zip_path)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    api_key = request.form.get("api_key", "").strip()
    count_raw = request.form.get("count", "0").strip()
    uploaded = request.files.get("zip_file")

    if not api_key:
        return jsonify({"error": "Anthropic API 키를 입력해주세요."}), 400
    if not uploaded or uploaded.filename == "":
        return jsonify({"error": "ZIP 파일을 선택해주세요."}), 400
    if not uploaded.filename.lower().endswith(".zip"):
        return jsonify({"error": "ZIP 파일만 업로드할 수 있습니다."}), 400

    try:
        count = int(count_raw) if count_raw else 0
    except ValueError:
        count = 0

    job_id = uuid.uuid4().hex
    job_dir = os.path.join(WORK_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    zip_path = os.path.join(job_dir, "input.zip")
    uploaded.save(zip_path)
    output_path = os.path.join(job_dir, "oa_analysis.xlsx")

    JOBS[job_id] = {
        "status": "processing",
        "done": 0,
        "total": 0,
        "message": "시작 중...",
        "output_path": output_path,
    }

    thread = threading.Thread(
        target=_process_job,
        args=(job_id, zip_path, output_path, api_key, count),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "존재하지 않는 작업입니다."}), 404
    return jsonify({
        "status": job["status"],
        "done": job["done"],
        "total": job["total"],
        "message": job["message"],
    })


@app.route("/download/<job_id>")
def download(job_id):
    job = JOBS.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "아직 처리가 완료되지 않았습니다."}), 400
    return send_file(
        job["output_path"],
        as_attachment=True,
        download_name="oa_analysis.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
