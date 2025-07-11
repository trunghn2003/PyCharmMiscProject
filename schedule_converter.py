import pandas as pd
import os
from datetime import datetime, timedelta

def convert_day_to_vietnamese(day_num):
    """Chuyển đổi số thứ sang tên thứ tiếng Việt"""
    days = {
        2: "Thứ 2",
        3: "Thứ 3", 
        4: "Thứ 4",
        5: "Thứ 5",
        6: "Thứ 6",
        7: "Thứ 7",
        8: "Chủ nhật"
    }
    return days.get(day_num, f"Thứ {day_num}")

def calculate_date_range(week_schedule):
    """Tính toán ngày bắt đầu và kết thúc dựa trên lịch tuần"""
    # Ngày bắt đầu học kỳ (25/08/2025)
    start_date = datetime(2025, 8, 25)
    
    # Tìm tuần đầu tiên có dấu 'x'
    first_week = None
    last_week = None
    
    week_columns = ['08/25', '09/25', '10/25', '11/25', '12/25']
    
    for i, col in enumerate(week_columns):
        if week_schedule.get(col) == 'x':
            if first_week is None:
                first_week = i
            last_week = i
    
    if first_week is not None and last_week is not None:
        # Tính ngày bắt đầu và kết thúc
        start = start_date + timedelta(weeks=first_week)
        end = start_date + timedelta(weeks=last_week + 1)
        return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
    
    return "2025-08-25", "2025-12-31"

def get_period_description(week_data):
    """Tạo mô tả thời gian học dựa trên các tuần có 'x'"""
    period_descriptions = []
    
    week_periods = {
        '08/25': "cuối tháng 8/2025",
        '09/25': "đầu tháng 9/2025", 
        '10/25': "đầu tháng 10/2025",
        '11/25': "đầu tháng 11/2025",
        '12/25': "đầu tháng 12/2025"
    }
    
    for week_col, description in week_periods.items():
        if str(week_data.get(week_col, '')).lower() == 'x':
            period_descriptions.append(description)
    
    if period_descriptions:
        return "Từ " + ", ".join(period_descriptions)
    return "Không xác định"

def process_schedule_data(df):
    """Xử lý dữ liệu thời khóa biểu - BAO GỒM TẤT CẢ CÁC CỘT"""
    
    # Lọc các dòng có tên môn học (ít nghiêm ngặt hơn)
    df_filtered = df.dropna(subset=['Tên môn học/ học phần'])
    
    # Lọc bỏ các dòng header
    df_filtered = df_filtered[df_filtered['Tên môn học/ học phần'] != 'Tên môn học/ học phần']
    
    # Lọc các dòng có ít nhất một tuần có lịch học
    week_cols = ['08/25', '09/25', '10/25', '11/25', '12/25']
    has_schedule = df_filtered[week_cols].apply(
        lambda row: row.astype(str).str.lower().str.contains('x', na=False).any(), axis=1
    )
    df_clean = df_filtered[has_schedule]
    
    print(f"📊 Số dòng sau khi lọc: {len(df_clean)} (từ {len(df)} dòng ban đầu)")
    
    schedule_list = []
    
    for index, row in df_clean.iterrows():
        try:
            # Lấy TẤT CẢ thông tin từ file gốc
            schedule_item = {}
            
            # Các cột cơ bản
            schedule_item['TT'] = row.get('TT', '')
            schedule_item['Mã môn học'] = str(row.get('Mã môn học', '')).strip() if pd.notna(row.get('Mã môn học')) else ""
            schedule_item['Tên môn học/ học phần'] = str(row.get('Tên môn học/ học phần', '')).strip()
            schedule_item['Khóa'] = str(row.get('Khóa', '')).strip() if pd.notna(row.get('Khóa')) else ""
            schedule_item['Ngành'] = str(row.get('Ngành', '')).strip() if pd.notna(row.get('Ngành')) else ""
            schedule_item['Lớp'] = str(row.get('Lớp', '')).strip() if pd.notna(row.get('Lớp')) else ""
            schedule_item['Nhóm'] = str(row.get('Nhóm', '')).strip() if pd.notna(row.get('Nhóm')) else ""
            schedule_item['Tổ hợp'] = str(row.get('Tổ hợp', '')).strip() if pd.notna(row.get('Tổ hợp')) else ""
            schedule_item['Tổ TH'] = str(row.get('Tổ TH', '')).strip() if pd.notna(row.get('Tổ TH')) else ""
            
            # Thông tin thời gian
            day_num = row.get('Thứ')
            schedule_item['Thứ'] = convert_day_to_vietnamese(day_num) if pd.notna(day_num) else ""
            schedule_item['Tiết BĐ'] = row.get('Tiết BĐ', '')
            schedule_item['Số tiết'] = row.get('Số tiết', '')
            
            # Tạo thời gian học chi tiết
            if pd.notna(schedule_item['Tiết BĐ']) and pd.notna(schedule_item['Số tiết']):
                try:
                    start_period = int(schedule_item['Tiết BĐ'])
                    num_periods = int(schedule_item['Số tiết'])
                    end_period = start_period + num_periods - 1
                    schedule_item['Thời gian chi tiết'] = f"Tiết {start_period}-{end_period}"
                except:
                    schedule_item['Thời gian chi tiết'] = f"Tiết {schedule_item['Tiết BĐ']}"
            else:
                schedule_item['Thời gian chi tiết'] = ""
            
            # Thông tin giảng viên và đơn vị
            schedule_item['Giảng viên giảng dạy'] = str(row.get('Giảng viên giảng dạy', '')).strip() if pd.notna(row.get('Giảng viên giảng dạy')) else ""
            schedule_item['Khoa'] = str(row.get('Khoa', '')).strip() if pd.notna(row.get('Khoa')) else ""
            schedule_item['Bộ môn'] = str(row.get('Bộ môn', '')).strip() if pd.notna(row.get('Bộ môn')) else ""
            schedule_item['Ghi chú'] = str(row.get('Ghi chú', '')).strip() if pd.notna(row.get('Ghi chú')) else ""
            
            # Thông tin địa điểm
            schedule_item['Phòng'] = str(row.get('Phòng', '')).strip() if pd.notna(row.get('Phòng')) else ""
            schedule_item['Nhà'] = str(row.get('Nhà', '')).strip() if pd.notna(row.get('Nhà')) else ""
            schedule_item['Số TC'] = row.get('Số TC', '')
            
            # Thông tin phân bổ chương trình
            schedule_item['LT'] = row.get('LT', '') if 'LT' in df.columns else ""
            schedule_item['TL/ BT'] = row.get('TL/ BT', '') if 'TL/ BT' in df.columns else ""
            schedule_item['BTL'] = row.get('BTL', '') if 'BTL' in df.columns else ""
            schedule_item['TH/ TN'] = row.get('TH/ TN', '') if 'TH/ TN' in df.columns else ""
            schedule_item['Tự học'] = row.get('Tự học', '') if 'Tự học' in df.columns else ""
            
            # Tính toán thời gian học theo tuần - THAY THẾ NGÀY THÁNG
            week_data = {}
            for week_col in week_cols:
                week_data[week_col] = row.get(week_col)
            
            # Thêm mô tả thời gian học
            schedule_item['Thời gian học'] = get_period_description(week_data)
            
            # Thêm chi tiết từng tuần (giữ lại để tham khảo)
            schedule_item['Tuần 25/8'] = 'Có' if str(row.get('08/25', '')).lower() == 'x' else 'Không'
            schedule_item['Tuần 1/9'] = 'Có' if str(row.get('09/25', '')).lower() == 'x' else 'Không'
            schedule_item['Tuần 6/10'] = 'Có' if str(row.get('10/25', '')).lower() == 'x' else 'Không'
            schedule_item['Tuần 3/11'] = 'Có' if str(row.get('11/25', '')).lower() == 'x' else 'Không'
            schedule_item['Tuần 1/12'] = 'Có' if str(row.get('12/25', '')).lower() == 'x' else 'Không'
            
            # Tạo địa điểm đầy đủ
            if schedule_item['Phòng'] and schedule_item['Nhà']:
                schedule_item['Địa điểm đầy đủ'] = f"Phòng {schedule_item['Phòng']}, Nhà {schedule_item['Nhà']}"
            elif schedule_item['Phòng']:
                schedule_item['Địa điểm đầy đủ'] = f"Phòng {schedule_item['Phòng']}"
            elif schedule_item['Nhà']:
                schedule_item['Địa điểm đầy đủ'] = f"Nhà {schedule_item['Nhà']}"
            else:
                schedule_item['Địa điểm đầy đủ'] = ""
                
            schedule_list.append(schedule_item)
            
        except Exception as e:
            print(f"⚠️ Lỗi xử lý dòng {index}: {e}")
            continue
    
    return schedule_list

def display_schedule_table(schedule_list):
    """Hiển thị bảng thời khóa biểu với TẤT CẢ CÁC CỘT"""
    if not schedule_list:
        print("Không có dữ liệu để hiển thị")
        return None
    
    # Tạo DataFrame từ danh sách
    df_result = pd.DataFrame(schedule_list)
    
    # Hiển thị bảng đẹp
    print("\n" + "="*180)
    print("📚 THỜI KHÓA BIỂU HỌC KỲ I NĂM HỌC 2025-2026 - ĐẦY ĐỦ TẤT CẢ CÁC CỘT")
    print("="*180)
    
    # Hiển thị 10 dòng đầu với tất cả các cột
    print("📄 10 DÒNG ĐẦU TIÊN VỚI TẤT CẢ THÔNG TIN:")
    print("-" * 180)
    
    # Thiết lập hiển thị pandas
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 25)
    
    print(df_result.head(10).to_string(index=False))
    
    if len(df_result) > 10:
        print(f"\n... và {len(df_result) - 10} dòng nữa")
    
    print(f"\n📊 Tổng số lớp học phần: {len(schedule_list)}")
    
    # Thống kê theo ngành
    if 'Ngành' in df_result.columns and not df_result['Ngành'].isnull().all():
        major_stats = df_result['Ngành'].value_counts().head(10)
        print("\n📈 Top 10 ngành có nhiều lớp nhất:")
        for major, count in major_stats.items():
            print(f"  - {major}: {count} lớp")
    
    # Thống kê theo khóa
    if 'Khóa' in df_result.columns and not df_result['Khóa'].isnull().all():
        year_stats = df_result['Khóa'].value_counts()
        print("\n📅 Thống kê theo khóa:")
        for year, count in year_stats.items():
            print(f"  - Khóa {year}: {count} lớp")
    
    # Thống kê thời gian học
    if 'Thời gian học' in df_result.columns:
        time_stats = df_result['Thời gian học'].value_counts().head(10)
        print("\n🕐 Thống kê thời gian học:")
        for time_period, count in time_stats.items():
            print(f"  - {time_period}: {count} lớp")
    
    print(f"\n📋 Tổng số cột dữ liệu: {len(df_result.columns)}")
    print(f"📝 Các cột: {', '.join(df_result.columns)}")
    
    return df_result

def save_to_excel(df_result, output_file="tkb_full_columns.xlsx"):
    """Lưu kết quả với TẤT CẢ CÁC CỘT vào file Excel mới"""
    try:
        # Tạo file Excel với nhiều sheet
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet chính với tất cả dữ liệu
            df_result.to_excel(writer, sheet_name='Thời khóa biểu đầy đủ', index=False)
            
            # Sheet tóm tắt với các cột chính
            summary_cols = ['TT', 'Mã môn học', 'Tên môn học/ học phần', 'Lớp', 'Nhóm', 'Thứ', 
                          'Tiết BĐ', 'Số tiết', 'Giảng viên giảng dạy', 'Phòng', 'Nhà', 
                          'Thời gian học', 'Địa điểm đầy đủ']
            df_summary = df_result[summary_cols]
            df_summary.to_excel(writer, sheet_name='Tóm tắt', index=False)
            
            # Sheet thống kê
            stats_data = []
            
            # Thống kê theo ngành
            if 'Ngành' in df_result.columns:
                major_stats = df_result['Ngành'].value_counts()
                for major, count in major_stats.items():
                    stats_data.append({'Loại': 'Ngành', 'Tên': major, 'Số lượng': count})
            
            # Thống kê theo khóa
            if 'Khóa' in df_result.columns:
                year_stats = df_result['Khóa'].value_counts()
                for year, count in year_stats.items():
                    stats_data.append({'Loại': 'Khóa', 'Tên': year, 'Số lượng': count})
            
            if stats_data:
                df_stats = pd.DataFrame(stats_data)
                df_stats.to_excel(writer, sheet_name='Thống kê', index=False)
        
        print(f"\n💾 Đã lưu dữ liệu với {len(df_result.columns)} cột vào file: {output_file}")
        print(f"📄 File chứa {len(df_result)} dòng dữ liệu trên 3 sheet:")
        print("   - 'Thời khóa biểu đầy đủ': Tất cả các cột")
        print("   - 'Tóm tắt': Các cột chính")
        print("   - 'Thống kê': Thống kê theo ngành và khóa")
        
    except Exception as e:
        print(f"Lỗi khi lưu file: {e}")

def main():
    """Hàm chính để xử lý file Excel"""
    # Đường dẫn tới file Excel
    excel_file = "0528_TKB_HK1_NAM_HOC_2025_20261-1.xlsx"
    
    # Kiểm tra xem file có tồn tại không
    if os.path.exists(excel_file):
        try:
            # Đọc file Excel
            print(f"Đang đọc file: {excel_file}")
            print("=" * 50)
            
            # Đọc tất cả các sheet trong file Excel
            excel_data = pd.ExcelFile(excel_file)
            print(f"Các sheet trong file: {excel_data.sheet_names}")
            print("=" * 50)
            
            # Xử lý sheet TKB CHINH
            if 'TKB CHINH' in excel_data.sheet_names:
                print("\n🔄 Đang xử lý sheet TKB CHINH...")
                
                # Đọc sheet TKB CHINH với header đúng
                df_main = pd.read_excel(excel_file, sheet_name='TKB CHINH', header=8)
                
                print(f"Số dòng dữ liệu: {len(df_main)}")
                print(f"Các cột: {list(df_main.columns)}")
                
                # Xử lý và chuyển đổi dữ liệu
                schedule_list = process_schedule_data(df_main)
                
                # Hiển thị kết quả
                df_result = display_schedule_table(schedule_list)
                
                # Lưu vào file Excel mới
                if df_result is not None and not df_result.empty:
                    save_to_excel(df_result)
                    
                    # Xuất mẫu dữ liệu theo format YÊU CẦU MỚI
                    print("\n📄 BẢNG THỜI KHÓA BIỂU VỚI TẤT CẢ THÔNG TIN CHÍNH:")
                    print("-" * 200)
                    print(f"{'TT':<5} | {'Mã MH':<10} | {'Tên môn học':<35} | {'Lớp':<6} | {'Nhóm':<6} | {'Thứ':<8} | {'Tiết':<8} | {'GV':<20} | {'Phòng':<8} | {'Thời gian học':<30}")
                    print("-" * 200)
                    
                    for i, row in df_result.head(15).iterrows():
                        subject = str(row['Tên môn học/ học phần'])[:35] + "..." if len(str(row['Tên môn học/ học phần'])) > 35 else str(row['Tên môn học/ học phần'])
                        teacher = str(row['Giảng viên giảng dạy'])[:20] + "..." if len(str(row['Giảng viên giảng dạy'])) > 20 else str(row['Giảng viên giảng dạy'])
                        time_desc = str(row['Thời gian học'])[:30] + "..." if len(str(row['Thời gian học'])) > 30 else str(row['Thời gian học'])
                        
                        print(f"{str(row['TT']):<5} | {str(row['Mã môn học']):<10} | {subject:<35} | {str(row['Lớp']):<6} | {str(row['Nhóm']):<6} | {str(row['Thứ']):<8} | {str(row['Thời gian chi tiết']):<8} | {teacher:<20} | {str(row['Phòng']):<8} | {time_desc:<30}")
                    
                    if len(df_result) > 15:
                        print(f"... và {len(df_result) - 15} dòng nữa")
                        
                    # Hiển thị ví dụ thời gian học mới
                    print("\n🕐 VÍ DỤ MÔ TẢ THỜI GIAN HỌC:")
                    unique_times = df_result['Thời gian học'].value_counts().head(5)
                    for time_desc, count in unique_times.items():
                        print(f"  - {time_desc}: {count} lớp")
            
            else:
                print("Không tìm thấy sheet 'TKB CHINH'")
                
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {e}")
            print("Có thể file bị hỏng hoặc định dạng không đúng.")
            
    else:
        print(f"Không tìm thấy file: {excel_file}")
        print("Các file trong thư mục hiện tại:")
        for file in os.listdir("."):
            if file.endswith((".xlsx", ".xls")):
                print(f"  - {file}")

if __name__ == "__main__":
    main()
