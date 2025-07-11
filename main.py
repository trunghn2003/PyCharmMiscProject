import pandas as pd
import os
from datetime import datetime, timedelta

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
            print("\n� Đang xử lý sheet TKB CHINH...")
            
            # Đọc sheet TKB CHINH với các tùy chọn để xử lý header
            df_main = pd.read_excel(excel_file, sheet_name='TKB CHINH', header=7)
            
            print(f"Số dòng dữ liệu: {len(df_main)}")
            
            # Xử lý và chuyển đổi dữ liệu
            schedule_list = process_schedule_data(df_main)
            
            # Hiển thị kết quả
            df_result = display_schedule_table(schedule_list)
            
            # Lưu vào file Excel mới
            if df_result is not None and not df_result.empty:
                save_to_excel(df_result)
                
                # Xuất một số mẫu CSV cho dễ xem
                print("\n📄 Một số mẫu dữ liệu:")
                print(df_result.head(10)[['Lớp', 'Mã lớp', 'Bắt đầu', 'Kết thúc', 'Thứ', 'Giảng viên', 'Môn học']].to_string(index=False))
        
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

def process_schedule_data(df):
    """Xử lý dữ liệu thời khóa biểu và chuyển đổi sang định dạng mong muốn"""
    
    # Lọc các dòng có dữ liệu hợp lệ (bỏ qua header và dòng trống)
    df_clean = df.dropna(subset=['Tên môn học/ học phần', 'Lớp', 'Giảng viên giảng dạy'])
    
    # Lọc các dòng không phải header
    df_clean = df_clean[df_clean['Tên môn học/ học phần'] != 'Tên môn học/ học phần']
    
    schedule_list = []
    
    for index, row in df_clean.iterrows():
        try:
            # Lấy thông tin cơ bản
            class_name = str(row.get('Lớp', '')).strip()
            class_code = str(row.get('Nhóm', '')).strip() if pd.notna(row.get('Nhóm')) else ""
            subject_name = str(row.get('Tên môn học/ học phần', '')).strip()
            teacher_name = str(row.get('Giảng viên giảng dạy', '')).strip()
            course_year = str(row.get('Khóa', '')).strip()
            major = str(row.get('Ngành', '')).strip()
            day_num = row.get('Thứ')
            start_period = row.get('Tiết BĐ')
            num_periods = row.get('Số tiết')
            room = str(row.get('Phòng', '')).strip() if pd.notna(row.get('Phòng')) else ""
            building = str(row.get('Nhà', '')).strip() if pd.notna(row.get('Nhà')) else ""
            credits = row.get('Số TC') if pd.notna(row.get('Số TC')) else ""
            notes = str(row.get('Ghi chú', '')).strip() if pd.notna(row.get('Ghi chú')) else ""
            
            # Chuyển đổi thứ
            day_name = convert_day_to_vietnamese(day_num) if pd.notna(day_num) else ""
            
            # Tính toán ngày bắt đầu và kết thúc
            week_data = {
                '08/25': row.get('08/25'),
                '09/25': row.get('09/25'), 
                '10/25': row.get('10/25'),
                '11/25': row.get('11/25'),
                '12/25': row.get('12/25')
            }
            start_date, end_date = calculate_date_range(week_data)
            
            # Tạo thời gian học
            time_slot = ""
            if pd.notna(start_period) and pd.notna(num_periods):
                end_period = int(start_period) + int(num_periods) - 1
                time_slot = f"Tiết {int(start_period)}-{end_period}"
            
            # Tạo địa điểm
            location = ""
            if room and building:
                location = f"Phòng {room}, Nhà {building}"
            elif room:
                location = f"Phòng {room}"
            elif building:
                location = f"Nhà {building}"
            
            schedule_item = {
                'Lớp': class_name,
                'Mã lớp': class_code,
                'Bắt đầu': start_date,
                'Kết thúc': end_date,
                'Thứ': day_name,
                'Giảng viên': teacher_name,
                'Môn học': subject_name,
                'Khóa': course_year,
                'Ngành': major,
                'Thời gian': time_slot,
                'Địa điểm': location,
                'Số tín chỉ': credits,
                'Ghi chú': notes
            }
            
            schedule_list.append(schedule_item)
            
        except Exception as e:
            print(f"Lỗi xử lý dòng {index}: {e}")
            continue
    
    return schedule_list

def display_schedule_table(schedule_list):
    """Hiển thị bảng thời khóa biểu theo định dạng yêu cầu"""
    if not schedule_list:
        print("Không có dữ liệu để hiển thị")
        return
    
    # Tạo DataFrame từ danh sách
    df_result = pd.DataFrame(schedule_list)
    
    # Hiển thị bảng đẹp
    print("\n" + "="*150)
    print("📚 THỜI KHÓA BIỂU HỌC KỲ I NĂM HỌC 2025-2026")
    print("="*150)
    
    # Hiển thị với định dạng đẹp
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 30)
    
    print(df_result.to_string(index=False))
    
    print(f"\n📊 Tổng số lớp học phần: {len(schedule_list)}")
    
    # Thống kê theo ngành
    if 'Ngành' in df_result.columns:
        major_stats = df_result['Ngành'].value_counts()
        print("\n📈 Thống kê theo ngành:")
        for major, count in major_stats.items():
            print(f"  - {major}: {count} lớp")
    
    # Thống kê theo khóa
    if 'Khóa' in df_result.columns:
        year_stats = df_result['Khóa'].value_counts()
        print("\n📅 Thống kê theo khóa:")
        for year, count in year_stats.items():
            print(f"  - Khóa {year}: {count} lớp")
    
    return df_result

def save_to_excel(df_result, output_file="tkb_formatted.xlsx"):
    """Lưu kết quả vào file Excel mới"""
    try:
        df_result.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n💾 Đã lưu dữ liệu vào file: {output_file}")
    except Exception as e:
        print(f"Lỗi khi lưu file: {e}")