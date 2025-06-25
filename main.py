import os
import logging
from read import readExecl
from log_config import setup_logging

# ...existing code...
def main():
    base_dir = r"D:\BT_FPT_HMD\DAT"
# ...existing code...
    
    input_file = os.path.join(base_dir, 'input.xlsx')
    download_dir = base_dir  
    output_file = os.path.join(base_dir, 'output.xlsx')
    logging.info("--- BẮT ĐẦU QUY TRÌNH TỰ ĐỘNG ---")
    logging.info("Đang tải và xử lý hóa đơn...")
    try:
        readExecl(input_file, download_dir, output_file)
        logging.info("Tải và xử lý hóa đơn hoàn tất.")
    except Exception:
        logging.error("Đã xảy ra lỗi nghiêm trọng trong quá trình tải và xử lý.", exc_info=True)
        return

    logging.info("--- QUY TRÌNH HOÀN TẤT THÀNH CÔNG ---")
    logging.info(f"File kết quả đã được lưu tại: {output_file}")


if __name__ == "__main__":
    setup_logging()
    main()