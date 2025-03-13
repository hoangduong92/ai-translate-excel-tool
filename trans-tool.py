import os
from openai import OpenAI
from openpyxl import load_workbook
import time
from tqdm import tqdm
import argparse
import threading
import logging
from datetime import datetime
import signal
import sys
import json
import re
from langdetect import detect
from concurrent.futures import ThreadPoolExecutor, as_completed
import traceback  # Thêm import này nếu bạn sử dụng traceback
from dotenv import load_dotenv
load_dotenv()

print(os.getenv("GEMINI_API_KEY"))

client = OpenAI(
    # base_url="https://api.groq.com/openai/v1",
    #base_url="https://api.openai.com/v1/",
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
    api_key=os.getenv("GEMINI_API_KEY"),
   # organization=os.getenv("GROQ_AI_ORG")
)

class ExcelTranslator:
    def __init__(self, workers=3, cache_file=None, log_file=None):
        """
        Khởi tạo translator với OpenAI API
        """
        # Thiết lập logging trước
        self.setup_logging(log_file)

        # Lấy API key từ biến môi trường
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            self.logger.error(
                "Không tìm thấy GEMINI_API_KEY trong biến môi trường")
            raise ValueError(
                "Không tìm thấy GEMINI_API_KEY trong biến môi trường")

        # Khởi tạo các thuộc tính khác
        client.api_key = api_key
        self.cached_translations = {}
        self.workers = workers
        self.should_exit = False
        self.cache_file = cache_file
        self.max_retries = 3
        self.base_delay = 1.0

        # Khóa để đảm bảo thread-safety
        self.cache_lock = threading.Lock()

        # Load cache và thiết lập signal handler
        self.load_cache()
        signal.signal(signal.SIGINT, self.handle_sigint)

    def setup_logging(self, log_file=None):
        """Thiết lập logging với định dạng chi tiết hơn"""
        if log_file is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = f"translation_log_{timestamp}.log"

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - [%(threadName)s] - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)

    def load_cache(self):
        """Nạp cache từ file JSON"""
        if self.cache_file and os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self.cached_translations = json.load(f)
                self.logger.info(
                    f"Đã nạp {len(self.cached_translations)} bản dịch từ cache")
            except Exception as e:
                self.logger.error(f"Lỗi khi nạp cache: {str(e)}")

    def save_cache(self):
        """Lưu cache vào file JSON"""
        if self.cache_file:
            with self.cache_lock:
                try:
                    with open(self.cache_file, 'w', encoding='utf-8') as f:
                        json.dump(self.cached_translations, f,
                                  ensure_ascii=False, indent=2)
                    self.logger.info(
                        f"Đã lưu {len(self.cached_translations)} bản dịch vào cache")
                except Exception as e:
                    self.logger.error(f"Lỗi khi lưu cache: {str(e)}")

    def handle_sigint(self, signum, frame):
        """Xử lý tín hiệu Ctrl+C"""
        self.logger.info("\nĐang dừng chương trình...")
        self.should_exit = True
        self.save_cache()
        sys.exit(0)

    def clean_text(self, text):
        """
        Chuẩn hóa text trước khi dịch
        """
        if not isinstance(text, str):
            return str(text)

        text = ''.join(char for char in text if ord(
            char) >= 32 or char in '\n\t')
        text = ' '.join(text.split())

        return text.strip()

    def should_translate_cell(self, cell, cell_info=""):
        """
        Kiểm tra xem một ô có cần dịch hay không
        """
        if cell.value is None:
            return False, "Ô trống"

        value = self.clean_text(str(cell.value))

        if not value:
            return False, "Ô chỉ chứa khoảng trắng hoặc ký tự đặc biệt"

        if len(value) <= 1:
            return False, "Văn bản quá ngắn"

        if re.match(r'^[\d\s,.-]+$', value):
            return False, "Ô chỉ chứa số và ký tự định dạng số"

        if value.startswith('='):
            return False, "Ô chứa công thức Excel"

        # # Phát hiện ngôn ngữ
        # try:
        #     lang = detect(value)
        #     if lang == 'ja':
        #         return False, "Văn bản đã là tiếng Nhật"
        # except:
        #     pass  # Nếu không phát hiện được ngôn ngữ, tiếp tục xử lý

        return True, "Cần dịch"

    def translate_to_japanese(self, text, cell_info=""):
        """
        Dịch văn bản sang tiếng Nhật sử dụng OpenAI API
        """
        try:
            with self.cache_lock:
                if text in self.cached_translations:
                    self.logger.debug(
                        f"{cell_info} - Sử dụng bản dịch từ cache")
                    return self.cached_translations[text]

            text = self.clean_text(text)
            if not text:
                return text

            for attempt in range(self.max_retries):
                try:
                    start_time = time.time()

                    response = client.chat.completions.create(
                        model="gemini-2.0-flash-lite",
                        messages=[
                            {"role": "system", "content": """You are a professional translator. Follow these rules strictly:
1. Output ONLY the Japanese translation, nothing else
2. DO NOT include the original text in your response
3. DO NOT add any explanations or notes
4. Keep IDs and special characters unchanged
5. Convert regular numbers to Japanese numbers (1->１, 2->２, etc.)
6. Use standard Japanese IT terminology for technical terms
7. Preserve the original formatting (spaces, line breaks)
8. For mixed language text, translate all non-Japanese parts to Japanese
9. Use proper Japanese particle usage (の, を, に, etc.)

Examples:

# Simple text
Input: "Save File"
Output: "ファイルを保存"

# Pure Japanese (keep unchanged)
Input: "CSV出力の設定"
Output: "CSV出力の設定"

# Mixed Vietnamese-Japanese with technical terms
Input: "1. Các item thuộc 検索 - Logo hiển thị đúng như design"
Output: "１．検索に属する項目 - ロゴはデザイン通りに表示されます"

# Mixed Vietnamese-Japanese with system terms
Input: "2. Kiểm tra 設定画面 và các chức năng liên quan"
Output: "２．設定画面と関連機能を確認します"

# Mixed English-Vietnamese-Japanese with line breaks
Input: "Check hiển thị default
1. Trên menu click バースデーカード"
Output: "デフォルト表示確認
１．メニューのバースデーカードをクリック"

# Test case format
Input: "TC01 - Kiểm tra màn hình 設定 - Check default value"
Output: "TC01 - 設定画面の確認 - デフォルト値を確認"

# Button and action terms
Input: "3. Click button 選択 để chọn file"
Output: "３．ファイルを選択するために選択ボタンをクリックします"
"""},
                            {"role": "user", "content": f"Translate the following text to Japanese:\n\n{text}"}
                        ]
                    )

                    translation_time = time.time() - start_time
                    translated_text = response.choices[0].message.content

                    if translated_text:
                        self.logger.info(
                            f"{cell_info}\n"
                            f"Văn bản gốc ({len(text)} ký tự): {text}\n"
                            f"Bản dịch ({len(translated_text)} ký tự): {translated_text}\n"
                            f"Thời gian dịch: {translation_time:.2f}s"
                        )

                        # with self.cache_lock:
                            # self.cached_translations[text] = translated_text
                        time.sleep(self.base_delay)
                        return translated_text

                    self.logger.warning(
                        f"{cell_info} - Lần thử {attempt + 1}: Không nhận được kết quả dịch"
                    )
                    time.sleep(self.base_delay * (attempt + 1))

                except Exception as e:
                    self.logger.warning(
                        f"{cell_info} - Lần thử {attempt + 1} thất bại: {str(e)}"
                    )
                    if attempt < self.max_retries - 1:
                        time.sleep(self.base_delay * (attempt + 1))
                        continue
                    raise

            self.logger.error(
                f"{cell_info} - Không thể dịch sau {self.max_retries} lần thử")
            return text

        except Exception as e:
            self.logger.error(f"{cell_info} - Lỗi dịch: {str(e)}")
            return text

    def process_excel_file(self, input_path):
        """
        Xử lý một file Excel: đọc nội dung, xác định các ô cần dịch, thực hiện dịch và lưu kết quả.
        """
        try:
            filename, ext = os.path.splitext(input_path)
            output_path = f"{filename}_translated{ext}"

            self.logger.info(
                f"Bắt đầu xử lý file: {os.path.basename(input_path)}")
            wb = load_workbook(input_path, data_only=False)

            file_start_time = time.time()
            translation_stats = {
                'total_cells': 0,
                'processed_cells': 0,
                'translated_cells': 0,
                'cached_translations': 0,
                'failed_translations': 0
            }

            for sheet_name in wb.sheetnames:
                if self.should_exit:
                    raise KeyboardInterrupt

                ws = wb[sheet_name]
                sheet_start_time = time.time()

                self.logger.info(f"Bắt đầu xử lý sheet: {sheet_name}")

                # Duyệt qua các ô trong sheet
                cells_to_translate = []

                for row in ws.iter_rows():
                    for cell in row:
                        translation_stats['total_cells'] += 1

                        cell_info = (
                            f"Sheet: {sheet_name}, "
                            f"Ô: {cell.coordinate}"
                        )

                        should_translate, reason = self.should_translate_cell(
                            cell, cell_info)

                        if should_translate:
                            cells_to_translate.append((cell, cell_info))
                            translation_stats['processed_cells'] += 1
                            self.logger.debug(f"{cell_info} - {reason}")
                        else:
                            self.logger.debug(
                                f"{cell_info} - Bỏ qua: {reason}")

                self.logger.info(
                    f"Thống kê sheet {sheet_name}:\n"
                    f"- Tổng số ô: {translation_stats['total_cells']}\n"
                    f"- Ô cần dịch: {translation_stats['processed_cells']}"
                )

                # Dịch các ô
                with tqdm(total=len(cells_to_translate), desc=f"Dịch sheet {sheet_name}", unit='cell') as pbar:
                    with ThreadPoolExecutor(max_workers=self.workers) as executor:
                        future_to_cell = {
                            executor.submit(
                                self.translate_to_japanese,
                                cell.value,
                                cell_info
                            ): (cell, cell_info)
                            for cell, cell_info in cells_to_translate
                        }

                        for future in as_completed(future_to_cell):
                            if self.should_exit:
                                raise KeyboardInterrupt

                            cell, cell_info = future_to_cell[future]
                            try:
                                translated_text = future.result()
                                if translated_text != cell.value:
                                    cell.value = translated_text
                                    translation_stats['translated_cells'] += 1
                                    with self.cache_lock:
                                        if cell.value in self.cached_translations:
                                            translation_stats['cached_translations'] += 1
                                pbar.update(1)
                            except Exception as e:
                                self.logger.error(
                                    f"Lỗi khi dịch {cell_info}: {str(e)}"
                                )
                                translation_stats['failed_translations'] += 1

                sheet_time = time.time() - sheet_start_time
                self.logger.info(
                    f"Hoàn thành sheet {sheet_name}:\n"
                    f"- Thời gian xử lý: {sheet_time:.2f}s"
                )

            wb.save(output_path)
            self.save_cache()

            total_time = time.time() - file_start_time
            self.logger.info(
                f"Hoàn thành file {os.path.basename(input_path)}:\n"
                f"- Tổng số ô: {translation_stats['total_cells']}\n"
                f"- Ô đã xử lý: {translation_stats['processed_cells']}\n"
                f"- Ô đã dịch: {translation_stats['translated_cells']}\n"
                f"- Dịch từ cache: {translation_stats['cached_translations']}\n"
                f"- Dịch thất bại: {translation_stats['failed_translations']}\n"
                f"- Tổng thời gian: {total_time:.2f}s\n"
                f"- Tốc độ trung bình: {translation_stats['translated_cells']/total_time:.2f} ô/giây"
            )

        except KeyboardInterrupt:
            self.logger.info("Đang dừng xử lý file...")
            try:
                wb.save(output_path)
                self.save_cache()
                self.logger.info(f"Đã lưu các thay đổi vào: {output_path}")
            except:
                self.logger.error("Không thể lưu file.")
            raise
        except Exception as e:
            self.logger.error(f"Lỗi khi xử lý file: {str(e)}")
            raise


def main():
    parser = argparse.ArgumentParser(
        description='Dịch nội dung file Excel sang tiếng Nhật')
    parser.add_argument('-d', '--directory',
                        help='Đường dẫn đến thư mục chứa file Excel cần dịch',
                        default=os.getcwd())
    parser.add_argument('-f', '--file',
                        help='Đường dẫn đến file Excel cần dịch',
                        default=None)
    parser.add_argument('-w', '--workers',
                        help='Số lượng worker threads (mặc định: 3)',
                        type=int,
                        default=3)
    parser.add_argument('-c', '--cache',
                        help='File cache để lưu các bản dịch',
                        default='translation_cache.json')
    parser.add_argument('-l', '--log',
                        help='File log output',
                        default=None)

    args = parser.parse_args()

    try:
        print(f"Khởi tạo translator...")
        translator = ExcelTranslator(
            workers=args.workers,
            cache_file=args.cache,
            log_file=args.log
        )

        # Process a single file if specified
        if args.file:
            if not os.path.isfile(args.file):
                print(f"Lỗi: File '{args.file}' không tồn tại!")
                return
                
            if not args.file.endswith('.xlsx') or args.file.endswith('_translated.xlsx'):
                print(f"Lỗi: '{args.file}' không phải là file Excel hợp lệ!")
                return
                
            print(f"\nĐang xử lý file: {os.path.basename(args.file)}")
            translator.process_excel_file(args.file)
            return

        print(f"Kiểm tra thư mục: {args.directory}")
        if not os.path.exists(args.directory):
            print(f"Lỗi: Thư mục '{args.directory}' không tồn tại!")
            return
            
        if not os.path.isdir(args.directory):
            print(f"Lỗi: '{args.directory}' không phải là thư mục!")
            return

        print("Tìm các file Excel...")
        excel_files = [
            f for f in os.listdir(args.directory)
            if f.endswith('.xlsx') and not f.endswith('_translated.xlsx')
        ]

        if not excel_files:
            print(
                f"Không tìm thấy file Excel nào trong thư mục '{args.directory}'")
            return

        print(f"Tìm thấy {len(excel_files)} file Excel:")
        for filename in excel_files:
            print(f"- {filename}")

        print("\nBắt đầu xử lý các file:")
        for filename in excel_files:
            file_path = os.path.join(args.directory, filename)
            try:
                print(f"\nĐang xử lý file: {filename}")
                translator.process_excel_file(file_path)
            except KeyboardInterrupt:
                print("\nĐã dừng chương trình theo yêu cầu.")
                break
            except Exception as e:
                print(f"Lỗi khi xử lý file {filename}: {str(e)}")

    except KeyboardInterrupt:
        print("\nĐã dừng chương trình.")
        sys.exit(0)
    except Exception as e:
        print(f"Lỗi: {str(e)}")
        traceback.print_exc()  # In ra stack trace đầy đủ
        sys.exit(1)


if __name__ == "__main__":
    main()
