import streamlit as st
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pytz
from streamlit_autorefresh import st_autorefresh

# Tự động làm mới trang mỗi 30 giây (30000 milliseconds) để cập nhật dữ liệu liên tục
st_autorefresh(interval=60000, key="data_refresh")

# Cài đặt trang web
st.set_page_config(page_title="Theo dõi tình trạng mổ", page_icon="🏥", layout="centered")

# Tên người cần theo dõi
TEN_BENH_NHAN = "LÊ THỊ THU HẰNG"
URL = "https://gmhs.bvungbuou.vn/tv/index2"

# Hàm lấy dữ liệu trực tiếp từ bệnh viện
def get_patient_info():
    try:
        response = requests.get(URL, timeout=10)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Quét tất cả các dòng trong bảng
        rows = soup.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            if len(cols) >= 5 and TEN_BENH_NHAN in cols[1].text:
                return {
                    "stt": cols[0].text.strip(),
                    "ten": cols[1].text.strip(),
                    # Chuyển các thẻ <br> thành dấu gạch ngang cho dễ đọc
                    "trang_thai": cols[2].get_text(separator=' - ', strip=True), 
                    "tinh_trang": cols[3].text.strip(),
                    "chuyen_khoa": cols[4].get_text(separator=' - ', strip=True)
                }
        return None
    except Exception as e:
        return "ERROR"

# Lấy thời gian hiện tại ở Việt Nam
tz = pytz.timezone('Asia/Ho_Chi_Minh')
current_time = datetime.now(tz).strftime("%H:%M:%S - %d/%m/%Y")

st.markdown("<h2 style='text-align: center; color: #333;'>Bảng theo dõi ca mổ</h2>", unsafe_allow_html=True)

# Lấy dữ liệu
data = get_patient_info()

# Style CSS cho chữ gradient và card
css = """
<style>
    .card {
        background-color: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.1);
        border-top: 5px solid #d81b60;
        margin-top: 20px;
    }
    .label {
        color: #777;
        font-size: 14px;
        margin-bottom: 2px;
        text-transform: uppercase;
        font-weight: 600;
    }
    .gradient-text {
        background: linear-gradient(to right, #d81b60, #8e24aa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: bold;
        font-size: 22px;
        margin-bottom: 15px;
    }
    .time-update {
        text-align: center;
        color: #888;
        font-style: italic;
        margin-top: 20px;
        font-size: 14px;
    }
</style>
"""
st.markdown(css, unsafe_allow_html=True)

# Hiển thị nội dung
if data == "ERROR":
    st.error("Lỗi kết nối đến hệ thống bệnh viện. Đang tự động thử lại...")
elif data:
    # Hiển thị thông tin trong một Card với HTML
    card_html = f"""
    <div class="card">
        <div class="label">Tên bệnh nhân</div>
        <div class="gradient-text">{data['ten']}</div>
        
        <div class="label">Trạng thái</div>
        <div class="gradient-text">{data['trang_thai'] if data['trang_thai'] else 'Chưa có'}</div>
        
        <div class="label">Tình trạng bệnh</div>
        <div class="gradient-text">{data['tinh_trang'] if data['tinh_trang'] else 'Đang cập nhật'}</div>
        
        <div class="label">Chuyển khoa</div>
        <div class="gradient-text">{data['chuyen_khoa'] if data['chuyen_khoa'] else 'Chưa chuyển'}</div>
    </div>
    """
    # DÙNG st.html Ở ĐÂY ĐỂ TRÁNH BỊ LỖI THỤT LỀ CỦA MARKDOWN
    st.html(card_html)
else:
    st.info(f"Đang theo dõi thông tin của bệnh nhân **{TEN_BENH_NHAN}**... Chưa có cập nhật hoặc chưa đến lượt.")

# Hiển thị thời gian làm mới
st.markdown(f"<div class='time-update'>Cập nhật lần cuối: {current_time} (Tự động tải lại sau mỗi 1 phút)</div>", unsafe_allow_html=True)
