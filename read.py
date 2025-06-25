import xml.etree.ElementTree as ET
from selenium.webdriver.common.by import By
from selenium import webdriver
from pandas import read_excel
import pandas as pd
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import logging

logger = logging.getLogger(__name__)

#Lấy dữ liệu từ file XML
def parse_invoice_xml(xml_file_path): 
    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        dlhdon = next((elem for elem in root.iter() if elem.tag.endswith('DLHDon')), None)
        if dlhdon is None:
            logger.warning(f"Không tìm thấy tag <DLHDon> trong file: {xml_file_path}")
            return None

        def find_text(element, path):
            parts = path.split('/')
            current = element
            for part in parts:
                if current is None:
                    return None
                current = next((child for child in current if child.tag.endswith(part)), None)
            return current.text.strip() if current is not None and current.text else None

        def find_in_ttkhac(parent, ttruong):
            ttkhac = next((c for c in parent if c.tag.endswith('TTKhac')), None)
            if ttkhac:
                for ttin in ttkhac:
                    if ttin.tag.endswith('TTin'):
                        name = value = None
                        for sub in ttin:
                            if sub.tag.endswith('TTruong') and (sub.text or '').strip() == ttruong:
                                name = sub
                            if sub.tag.endswith('DLieu'):
                                value = sub
                        if name and value:
                            return value.text.strip() if value.text else None
            return None

        stknhang = find_text(dlhdon, 'NDHDon/NBan/STKNHang')
        if not stknhang:
            nban = next((c for c in dlhdon.iter() if c.tag.endswith('NBan')), None)
            if nban:
                stknhang = find_in_ttkhac(nban, 'SellerBankAccount')

        file_name = os.path.basename(xml_file_path)

        data = {
            'Số hóa đơn': find_text(dlhdon, 'TTChung/SHDon') or '',
            'Đơn vị bán hàng': find_text(dlhdon, 'NDHDon/NBan/Ten') or '',
            'Mã số thuế bán': find_text(dlhdon, 'NDHDon/NBan/MST') or '',
            'Địa chỉ bán': find_text(dlhdon, 'NDHDon/NBan/DChi') or '',
            'Số tài khoản bán': stknhang or '',
            'Họ tên người mua hàng': find_text(dlhdon, 'NDHDon/NMua/Ten') or '',
            'Địa chỉ mua': find_text(dlhdon, 'NDHDon/NMua/DChi') or '',
            'Mã số thuế mua': find_text(dlhdon, 'NDHDon/NMua/MST') or '',
            'Tên file XML': file_name
        }

        logger.info(f"Parsed XML file '{file_name}' thành công.")
        return data

    except ET.ParseError:
        logger.error(f"Lỗi parsing XML: {xml_file_path}", exc_info=True)
        return None
    except Exception:
        logger.error(f"Lỗi không xác định khi xử lý file {xml_file_path}", exc_info=True)
        return None

def handle_fpt(driver, wait, url, ma_so_thue, ma_tra_cuu, logger):
    try:
        driver.get(url)
        input_mst = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='MST bên bán']")))
        input_mst.send_keys(ma_so_thue)
        input_mtc = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Mã tra cứu hóa đơn']")
        input_mtc.send_keys(ma_tra_cuu)
        search_button = driver.find_element(By.CSS_SELECTOR, "body > div.webix_view.webix_scrollview.scoll-page > div > div > div.webix_view.webix-container.subview-container.webix_layout_line > div > div.webix_view.webix_form.search-form.bgb.webix_layout_form > div > div.webix_view.search-form-btn.webix_layout_line > div.webix_view.webix_control.webix_el_button.webix_secondary > div > button")
        search_button.click()

        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'MST không hợp lệ')]")))
            logger.warning(f"FPT: MST không hợp lệ: {ma_so_thue}, bỏ qua.")
            return False
        except TimeoutException:
            pass

        invoice_button_selector = (By.CSS_SELECTOR, "body > div.webix_view.webix_scrollview.scoll-page > div > div > div.webix_view.webix-container.subview-container.webix_layout_line > div > div.webix_view.webix_form.search-form.bgb.webix_layout_form > div > div.webix_view.search-form-btn.webix_layout_line > div:nth-child(2) > div > button")
        invoice_button = wait.until(EC.element_to_be_clickable(invoice_button_selector))
        invoice_button.click()
        logger.info(f"Đã yêu cầu tải hóa đơn FPT cho MST: {ma_so_thue}")
        return True
    except TimeoutException:
        logger.warning(f"Không tìm thấy nút tải hóa đơn FPT cho MST: {ma_so_thue}")
    except Exception as e:
        logger.error(f"Lỗi FPT ({ma_so_thue}): {e}", exc_info=False)
    return False


def handle_meinvoice(driver, wait, url, ma_tra_cuu, logger):
    try:
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Nhập mã tra cứu hóa đơn']"))).send_keys(ma_tra_cuu)
        driver.find_element(By.ID, "btnSearchInvoice").click()

        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(),'Không tìm thấy hóa đơn')]")))
            logger.warning(f"MeInvoice: Mã tra cứu không hợp lệ hoặc không tìm thấy: {ma_tra_cuu}")
            return False
        except TimeoutException:
            logger.info(f"MeInvoice: Tìm thấy hóa đơn cho mã tra cứu: {ma_tra_cuu}. Bắt đầu tải.")

        download_span = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.download-invoice")))
        download_span.click()
        logger.info("Đã nhấp vào nút tải xuống MeInvoice")

        xml_download_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.txt-download-xml")))
        xml_download_link.click()
        logger.info(f"Đã yêu cầu tải file XML từ MeInvoice cho Mã Tra Cứu: {ma_tra_cuu}")
        return True
    except TimeoutException:
        logger.warning(f"MeInvoice: Không tìm thấy thành phần cần thiết cho {ma_tra_cuu}")
    except Exception as e:
        logger.error(f"Lỗi MeInvoice ({ma_tra_cuu}): {e}", exc_info=False)
    return False


def handle_hoadon(driver, wait, url, ma_tra_cuu, logger):
    try:
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.ID, "txtInvoiceCode"))).send_keys(ma_tra_cuu)
        driver.find_element(By.CLASS_NAME, "btnSearch").click()
        time.sleep(2)
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameViewInvoice')))
        logger.info("Đã chuyển vào iframe")

        if "khong dung" in driver.page_source.lower() or "khong tim thay" in driver.page_source.lower():
            logger.warning(f"eHoaDon: Mã tra cứu không hợp lệ: {ma_tra_cuu}")
            driver.find_element(By.XPATH, "//button[contains(text(),'Đóng')]").click()
            return False

        btn_download = wait.until(EC.element_to_be_clickable((By.ID, "btnDownload")))
        btn_download.click()
        logger.info("Đã click nút tải")

        link_xml = wait.until(EC.element_to_be_clickable((By.ID, "LinkDownXML")))
        link_xml.click()
        logger.info("Đã click tải XML")
        return True
    except TimeoutException:
        logger.warning(f"HoaDon: Timeout khi xử lý {ma_tra_cuu}")
    except Exception as e:
        logger.error(f"Lỗi eHoaDon ({ma_tra_cuu}): {e}", exc_info=False)
    return False

def readExecl(input_file: str, download_dir: str, output_file: str):
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True
    })
    driver = webdriver.Chrome(options=chrome_options)
    
    extracted_data = []
    
    try:
        df_input = read_excel(input_file, dtype={'Mã số thuế': str, 'Mã tra cứu': str, 'URL': str})
        logger.info(f"Đã đọc {len(df_input)} dòng từ file Excel '{os.path.basename(input_file)}'.")

        processed_files = set()

        for index, row in df_input.iterrows():
            ma_so_thue = row.get('Mã số thuế', '')
            ma_tra_cuu = row.get('Mã tra cứu', '')
            url = row.get('URL', '')

            if not url or not pd.notna(ma_tra_cuu) or not ma_tra_cuu:
                logger.warning(f"Dòng {index+2} thiếu dữ liệu URL hoặc Mã tra cứu, bỏ qua.")
                continue

            wait = WebDriverWait(driver, 10)
            
            if "fpt" in url.lower():
                if not handle_fpt(driver, wait, url, ma_so_thue, ma_tra_cuu, logger):
                    continue

            elif "meinvoice" in url.lower():
                if not handle_meinvoice(driver, wait, url, ma_tra_cuu, logger):
                    continue

            elif "hoadon" in url.lower():
                if not handle_hoadon(driver, wait, url, ma_tra_cuu, logger):
                    continue

            time.sleep(4)  

            try:
                xml_files = [os.path.join(download_dir, f) for f in os.listdir(download_dir) if f.endswith('.xml')]
                new_files = [f for f in xml_files if f not in processed_files]
                
                if new_files:
                    latest_file = max(new_files, key=os.path.getmtime)
                    logger.info(f"Đang xử lý file: {os.path.basename(latest_file)}")
                    data = parse_invoice_xml(latest_file)
                    if data:
                        logger.info(f"Dữ liệu trích xuất: {data}")
                        data['Mã tra cứu'] = ma_tra_cuu
                        extracted_data.append(data)
                        processed_files.add(latest_file)
                else:
                    logger.warning(f"Không tìm thấy file XML mới nào được tải về cho mã tra cứu: {ma_tra_cuu}")
            except Exception as e:
                logger.error(f"Lỗi khi xử lý file XML cho MTC {ma_tra_cuu}: {e}")

    finally:
        logger.info("Hoàn tất tất cả tác vụ tải về. Đóng trình duyệt.")
        time.sleep(2)
        driver.quit()

    if not extracted_data:
        logger.warning("Không trích xuất được dữ liệu từ bất kỳ file XML nào. Ghi file output không có dữ liệu mới.")
        df_input.to_excel(output_file, index=False, engine='openpyxl')
        return

    df_xml = pd.DataFrame(extracted_data)
    df_input['Mã tra cứu'] = df_input['Mã tra cứu'].astype(str).str.strip()
    
    df_xml.dropna(subset=['Mã tra cứu'], inplace=True)
    
    logger.info("Kết hợp dữ liệu từ Excel và XML...")
    df_merged = pd.merge(df_input, df_xml, on='Mã tra cứu', how='left')

    for col in df_input.columns:
        if f"{col}_x" in df_merged.columns and f"{col}_y" in df_merged.columns:
            df_merged[col] = df_merged[f"{col}_y"].fillna(df_merged[f"{col}_x"])
            df_merged.drop(columns=[f"{col}_x", f"{col}_y"], inplace=True)

    df_merged.to_excel(output_file, index=False, engine='openpyxl')
    logger.info(f"Thành công! Dữ liệu đã được kết hợp và lưu vào file: {output_file}")


