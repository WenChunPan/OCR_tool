from flask import Flask, request, jsonify
import requests
from flask_cors import CORS
import base64
from openpyxl import load_workbook
from io import BytesIO

# 初始化Flask
app = Flask(__name__)

# 啟用CORS，讓前端可以跨域請求
CORS(app)

# ocr api的url
ocr_api_url = (
    # "http://192.168.0.160:30020/ai/service/v2/recognize/table/multipage?excel=1"
    "http://leda-textin.seadeep.ai/ai/service/v2/recognize/table/multipage?excel=1"
)


# call ocr api
@app.route("/ocr", methods=["POST"])
def ocr_proxy():
    # 如果沒有上傳檔案
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    # 如果有上傳檔案
    file = request.files["file"]
    payload = file.read()

    # 發送請求到ocr api
    headers = {"Content-Type": "application/octet-stream"}
    try:
        # 嘗試發送請求
        ocr_response = requests.post(ocr_api_url, headers=headers, data=payload)
    except Exception as e:
        # 如果發生錯誤，回傳錯誤訊息
        return jsonify({"error": "Failed to contact OCR API", "detail": str(e)}), 500

    if ocr_response.status_code != 200:
        # 如果回傳狀態碼不是200，回傳錯誤訊息
        return jsonify(
            {
                "error": "OCR API returned an error",
                "status": ocr_response.status_code,
                "detail": ocr_response.text,
            }
        ), 500

    # 解析回傳資料
    try:
        result_data = ocr_response.json()
        excel_base64 = result_data["result"]["excel"]
        json_result = result_data["result"]
    except Exception as e:
        return jsonify(
            {"error": "Failed to parse OCR API response", "detail": str(e)}
        ), 500

    # 檢查 Excel 是否為空（A1有無值）
    try:
        excel_bytes = base64.b64decode(excel_base64)
        workbook = load_workbook(filname=BytesIO(excel_bytes), data_only=True)
        a1_value = workbook.active["A1"].value
        is_excel_empty = not bool(a1_value)
    except Exception as e:
        return jsonify({"error": "Failed to process Excel file", "detail": str(e)}), 500

    # 回傳給前端
    return jsonify(
        {
            "filename": file.filename,
            "excel_base64": excel_base64,
            "ocr_json": json_result,
            "has_excel_data": not is_excel_empty,  # 回傳 true 或 false 給前端
        }
    )


# 啟動Flask
if __name__ == "__main__":
    app.run(debug=True, port=5000)
