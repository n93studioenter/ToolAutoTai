import time
import os
import sys
import base64
import ddddocr

# ═══════════════════════════════════════════════════════════
# XÁC ĐỊNH BASE = thư mục chứa exe (KHÔNG phải _internal)
# ═══════════════════════════════════════════════════════════
if getattr(sys, 'frozen', False):
    # Chạy dạng .exe → lấy thư mục chứa exe
    BASE = os.path.dirname(sys.executable)
else:
    # Chạy dạng .py
    BASE = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE)

print("🔄 Đang khởi tạo mô hình ddddocr...")
ocr = ddddocr.DdddOcr(show_ad=False, beta=True)

REQUEST_FILE = os.path.join(BASE, "request.txt")
RESULT_FILE = os.path.join(BASE, "result.txt")
READY_FILE = os.path.join(BASE, "ready.flag")

print(f"🚀 Python OCR Service đang chạy tại: {BASE}")
print(f"   - Request file: {REQUEST_FILE}")
print(f"   - Result file:  {RESULT_FILE}")

# Ghi ready flag để C# biết đã sẵn sàng
with open(READY_FILE, "w") as f:
    f.write("ok")

last_mtime = 0

while True:
    try:
        if os.path.exists(REQUEST_FILE):
            mtime = os.path.getmtime(REQUEST_FILE)

            if mtime > last_mtime:
                last_mtime = mtime

                with open(REQUEST_FILE, "r", encoding="utf-8") as f:
                    content = f.read().strip()

                if content:
                    image_bytes = None

                    # Nếu content là đường dẫn file ảnh
                    if os.path.exists(content):
                        with open(content, "rb") as img_f:
                            image_bytes = img_f.read()
                    else:
                        # Nếu content là base64
                        try:
                            base64_data = content
                            if "base64," in base64_data:
                                base64_data = base64_data.split("base64,")[1]
                            image_bytes = base64.b64decode(base64_data)
                        except Exception as decode_err:
                            print(f"❌ Lỗi giải mã dữ liệu ảnh: {decode_err}")

                    if image_bytes:
                        raw_result = ocr.classification(image_bytes)
                        result = "".join([c for c in raw_result if c.isdigit()])[:6]

                        with open(RESULT_FILE, "w", encoding="utf-8") as f:
                            f.write(result)

                        print(f"✅ Xong ---> Kết quả: {result}")

                    try:
                        os.remove(REQUEST_FILE)
                    except:
                        pass

        time.sleep(0.02)
    except Exception as e:
        print(f"❌ Lỗi Service: {e}")
        time.sleep(0.5)