import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import date

# --- 設定頁面 ---
st.set_page_config(page_title="凱德訂單生成器", layout="wide")

st.title("📝 快速 Excel 訂單生成器")
st.markdown("輸入客戶資訊與商品明細，自動套用格式並輸出 Excel。")

# --- 側邊欄：設定與上傳 ---
st.sidebar.header("1. 系統設定")
uploaded_template = st.sidebar.file_uploader("上傳 Excel 範本 (template.xlsx)", type=["xlsx"])

# 如果沒有上傳，嘗試讀取本地預設檔案
template_source = uploaded_template if uploaded_template else "template.xlsx"

# --- 主畫面：客戶資訊 ---
st.header("2. 客戶資訊")

col1, col2 = st.columns(2)

with col1:
    customer_name = st.text_input("客戶名稱", "凱德科技股份有限公司")
    department = st.text_input("隸屬部門", "管理部")
    contact_person = st.text_input("聯絡人", "游豐聰 Arnode Yu")
    phone = st.text_input("公司電話", "02-77161899 ext 208")

with col2:
    mobile = st.text_input("行動電話", "0931-107-252")
    email = st.text_input("E-mail", "arnode@cadex.com.tw")
    address = st.text_input("公司地址", "11494台北市內湖區新湖二路168號2樓")
    quotation_date = st.date_input("報價日期", date.today())

# --- 主畫面：商品明細 ---
st.header("3. 商品明細")
st.info("💡 直接在表格中輸入，點擊下方「+」新增列，完成後勾選刪除多餘空行。")

# 初始化預設表格資料
if "df_items" not in st.session_state:
    st.session_state.df_items = pd.DataFrame(
        [
            {"廠牌": "DELL", "型號": "Pro Max Tower T2", "規格": "U7-265 / 64GB / 1TB SSD", "數量": 1, "單價": 83880},
            {"廠牌": "Service", "型號": "NBD", "規格": "FC Support Warranty", "數量": 1, "單價": 0},
        ]
    )

# 顯示可編輯的表格
edited_df = st.data_editor(
    st.session_state.df_items,
    num_rows="dynamic",  # 允許使用者新增刪除列
    column_config={
        "數量": st.column_config.NumberColumn(min_value=1, format="%d"),
        "單價": st.column_config.NumberColumn(format="$%d"),
    },
    use_container_width=True
)

# --- 核心邏輯：生成 Excel ---
def generate_excel(template_src, data, items_df):
    try:
        # 載入 Excel
        wb = openpyxl.load_workbook(template_src)
        ws = wb.active
        
        # --- 填寫客戶資料 (座標需依照您的實際 Excel 調整) ---
        # 這裡的座標是根據您之前的 CSV 推測的，請打開您的 template.xlsx 確認並修改
        ws['B12'] = data['customer_name'] # 客戶名稱
        ws['B13'] = data['department']    # 隸屬部門
        ws['B14'] = data['contact_person']# 聯絡人
        ws['B15'] = data['phone']         # 公司電話
        ws['B17'] = data['mobile']        # 行動電話
        ws['B19'] = data['address']       # 公司地址
        ws['B20'] = data['email']         # Email
        
        # 報價日期 (假設在右下角或右上角，請自行調整)
        # ws['F45'] = data['quotation_date'] 

        # --- 填寫商品明細 ---
        start_row = 21  # 商品起始列
        
        for index, row in items_df.iterrows():
            current_row = start_row + index
            
            # 確保不會填寫太少資料
            if not row["廠牌"] and not row["型號"]:
                continue

            ws[f'A{current_row}'] = row['廠牌']
            ws[f'B{current_row}'] = row['型號']
            ws[f'C{current_row}'] = row['規格']
            ws[f'D{current_row}'] = row['數量']
            ws[f'E{current_row}'] = row['單價']
            
            # 計算小計 (如果 Excel 範本裡該格已有公式，這行可以註解掉)
            ws[f'F{current_row}'] = row['數量'] * row['單價']

        # 儲存到記憶體中 (不存硬碟)
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"發生錯誤: {e}")
        return None

# --- 按鈕區 ---
st.divider()
if st.button("🚀 生成報價單 Excel", type="primary"):
    # 準備資料字典
    customer_data = {
        "customer_name": customer_name,
        "department": department,
        "contact_person": contact_person,
        "phone": phone,
        "mobile": mobile,
        "email": email,
        "address": address,
        "quotation_date": quotation_date
    }
    
    # 執行生成
    excel_file = generate_excel(template_source, customer_data, edited_df)
    
    if excel_file:
        file_name = f"報價單_{customer_name}_{date.today()}.xlsx"
        st.success("檔案生成成功！請點擊下方按鈕下載。")
        st.download_button(
            label="📥 下載 Excel 檔案",
            data=excel_file,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )