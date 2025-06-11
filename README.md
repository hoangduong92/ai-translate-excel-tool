# AI Excel Translation Tool

## 0. Cài đặt thư viện

- Cài đặt các thư viện cần thiết bằng lệnh:

  ```
  pip install -r requirements.txt
  ```

## 1. Cấu hình Google API Key

- Đăng ký và lấy Google API Key từ [Get API key | Google AI Studio](https://aistudio.google.com/apikey)

## 2. Cấu hình file .env

- Tạo file `.env` từ file `.env.sample` bằng cách xóa đuôi `.sample`
- Dán API key vào file `.env`:
  ```
  GEMINI_API_KEY=YOUR_API_KEY_HERE
  ```

## 3. File System Prompt cho các hướng dịch

- Dịch Việt → Nhật: sử dụng file `system_prompt_vi_to_ja.txt`
- Dịch Nhật → Việt: sử dụng file `system_prompt_ja_to_vi.txt`

## 4. Cấu hình hướng dịch

- Mở file `trans-tool.py`
- Sửa biến global `SYSTEM_PROMPT_FILE` thành tên file system prompt tương ứng:
  ```python
  # Để dịch từ tiếng Việt sang tiếng Nhật
  SYSTEM_PROMPT_FILE = "system_prompt_vi_to_ja.txt"

  # Hoặc để dịch từ tiếng Nhật sang tiếng Việt
  SYSTEM_PROMPT_FILE = "system_prompt_ja_to_vi.txt"
  ```

## 5. Chuẩn bị file cần dịch

- Đặt file Excel cần dịch vào thư mục `input`

## 6. Dịch file

- Chạy lệnh sau trong terminal:

  ```
  python trans-tool.py
  ```

  hoặc

  ```
  python trans-tool.py --file_path input/your-file.xlsx --sheet_name Sheet1
  ```

## 7. File đã dịch

File đã dịch sẽ được lưu trong thư mục `output`

## 8. Chú ý về API_DELAY

* Điều chỉnh theo limit quota của API. Hiện tại trong source để là 2 giây vì phù hợp với free tier 30RPM của gemini-2.0-flash-lite.
