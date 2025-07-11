import pandas as pd
import os
from datetime import datetime, timedelta

def convert_day_to_vietnamese(day_num):
    """Chuyển đổi số thứ sang tên thứ tiếng Việt"""
    if pd.isna(day_num):
        return ""
    days = {
        2: "Thứ 2",
        3: "Thứ 3", 
        4: "Thứ 4",
        5: "Thứ 5",
        6: "Thứ 6",
        7: "Thứ 7",
        8: "Chủ nhật"
    }
    return days.get(int(day_num), f"Thứ {int(day_num)}")

def calculate_date_range_improved(week_schedule):
    """Tính toán ngày bắt đầu và kết thúc dựa trên lịch tuần - PHIÊN BẢN CẢI TIẾN"""
    
    # Các tuần trong học kỳ với ngày bắt đầu thực tế
    week_start_dates = {
        '08/25': datetime(2025, 8, 25),   # Tuần 1: 25/8/2025
        '09/25': datetime(2025, 9, 1),   # Tuần 2: 1/9/2025  
        '10/25': datetime(2025, 10, 6),  # Tuần 3: 6/10/2025
        '11/25': datetime(2025, 11, 3),  # Tuần 4: 3/11/2025
        '12/25': datetime(2025, 12, 1),  # Tuần 5: 1/12/2025
    }
    
    # Tìm tuần đầu tiên và cuối cùng có dấu 'x'
    first_week_date = None
    last_week_date = None
    
    for week_col, start_date in week_start_dates.items():
        if str(week_schedule.get(week_col, '')).lower() == 'x':
            if first_week_date is None:
                first_week_date = start_date
            last_week_date = start_date
    
    if first_week_date and last_week_date:
        # Ngày kết thúc là cuối tuần cuối cùng (thêm 6 ngày)
        end_date = last_week_date + timedelta(days=6)
        return first_week_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")
    
    # Mặc định nếu không tìm thấy lịch
    return "2025-08-25", "2025-12-31"

def process_schedule_data_improved(df):
    """Xử lý dữ liệu thời khóa biểu - PHIÊN BẢN CẢI TIẾN"""
    
    # Lọc các dòng có tên môn học (không yêu cầu đầy đủ tất cả thông tin)
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
            # Lấy thông tin cơ bản - cho phép thiếu một số thông tin
            class_name = str(row.get('Lớp', '')).strip() if pd.notna(row.get('Lớp')) else ""
            class_code = str(row.get('Nhóm', '')).strip() if pd.notna(row.get('Nhóm')) else ""
            subject_name = str(row.get('Tên môn học/ học phần', '')).strip()
            teacher_name = str(row.get('Giảng viên giảng dạy', '')).strip() if pd.notna(row.get('Giảng viên giảng dạy')) else ""
            course_year = str(row.get('Khóa', '')).strip() if pd.notna(row.get('Khóa')) else ""
            major = str(row.get('Ngành', '')).strip() if pd.notna(row.get('Ngành')) else ""
            day_num = row.get('Thứ')
            start_period = row.get('Tiết BĐ')
            num_periods = row.get('Số tiết')
            room = str(row.get('Phòng', '')).strip() if pd.notna(row.get('Phòng')) else ""
            building = str(row.get('Nhà', '')).strip() if pd.notna(row.get('Nhà')) else ""
            credits = row.get('Số TC') if pd.notna(row.get('Số TC')) else ""
            notes = str(row.get('Ghi chú', '')).strip() if pd.notna(row.get('Ghi chú')) else ""
            subject_code = str(row.get('Mã môn học', '')).strip() if pd.notna(row.get('Mã môn học')) else ""
            
            # Chuyển đổi thứ
            day_name = convert_day_to_vietnamese(day_num)
            
            # Tính toán ngày bắt đầu và kết thúc CHÍNH XÁC
            week_data = {}
            for week_col in week_cols:
                week_data[week_col] = row.get(week_col)
            
            start_date, end_date = calculate_date_range_improved(week_data)
            
            # Tạo thời gian học
            time_slot = ""
            if pd.notna(start_period) and pd.notna(num_periods):
                try:
                    end_period = int(start_period) + int(num_periods) - 1
                    time_slot = f"Tiết {int(start_period)}-{end_period}"
                except:
                    time_slot = f"Tiết {start_period}"
            
            # Tạo địa điểm
            location = ""
            if room and building:
                location = f"Phòng {room}, Nhà {building}"
            elif room:
                location = f"Phòng {room}"
            elif building:
                location = f"Nhà {building}"
            
            # Bỏ qua các dòng không có thông tin cơ bản tối thiểu
            if not subject_name:
                continue
                
            schedule_item = {
                'Lớp': class_name,
                'Mã lớp': class_code,
                'Bắt đầu': start_date,
                'Kết thúc': end_date,
                'Thứ': day_name,
                'Giảng viên': teacher_name,
                'Môn học': subject_name,
                'Mã môn học': subject_code,
                'Khóa': course_year,
                'Ngành': major,
                'Thời gian': time_slot,
                'Địa điểm': location,
                'Số tín chỉ': credits,
                'Ghi chú': notes
            }
            
            schedule_list.append(schedule_item)
            
        except Exception as e:
            print(f"⚠️ Lỗi xử lý dòng {index}: {e}")
            continue
    
    return schedule_list

def display_schedule_table(schedule_list):
    """Hiển thị bảng thời khóa biểu theo định dạng yêu cầu"""
    if not schedule_list:
        print("Không có dữ liệu để hiển thị")
        return None
    
    # Tạo DataFrame từ danh sách
    df_result = pd.DataFrame(schedule_list)
    
    # Hiển thị bảng đẹp
    print("\n" + "="*150)
    print("📚 THỜI KHÓA BIỂU HỌC KỲ I NĂM HỌC 2025-2026 (ĐÃ SỬA)")
    print("="*150)
    
    # Hiển thị với định dạng đẹp - chỉ 20 dòng đầu để tránh quá dài
    print("📄 20 DÒNG ĐẦU:")
    print(df_result.head(20).to_string(index=False))
    
    print(f"\n📊 Tổng số lớp học phần: {len(schedule_list)}")
    
    # Thống kê theo ngành
    if 'Ngành' in df_result.columns and not df_result['Ngành'].empty:
        major_stats = df_result['Ngành'].value_counts().head(10)
        print("\n📈 Top 10 ngành có nhiều lớp nhất:")
        for major, count in major_stats.items():
            print(f"  - {major}: {count} lớp")
    
    # Thống kê theo khóa
    if 'Khóa' in df_result.columns and not df_result['Khóa'].empty:
        year_stats = df_result['Khóa'].value_counts()
        print("\n📅 Thống kê theo khóa:")
        for year, count in year_stats.items():
            print(f"  - Khóa {year}: {count} lớp")
    
    return df_result

def save_to_excel(df_result, output_file="tkb_formatted_fixed.xlsx"):
    """Lưu kết quả vào file Excel mới"""
    try:
        df_result.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n💾 Đã lưu dữ liệu vào file: {output_file}")
    except Exception as e:
        print(f"Lỗi khi lưu file: {e}")

def main():
    """Hàm chính để xử lý file Excel - PHIÊN BẢN ĐÃ SỬA"""
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
                
                print(f"Số dòng dữ liệu ban đầu: {len(df_main)}")
                
                # Xử lý và chuyển đổi dữ liệu với logic mới
                schedule_list = process_schedule_data_improved(df_main)
                
                # Hiển thị kết quả
                df_result = display_schedule_table(schedule_list)
                
                # Lưu vào file Excel mới
                if df_result is not None and not df_result.empty:
                    save_to_excel(df_result)
                    
                    # Xuất mẫu dữ liệu theo format yêu cầu
                    print("\n📄 BẢNG THỜI KHÓA BIỂU THEO ĐỊNH DẠNG YÊU CẦU (ĐÃ SỬA):")
                    print("-" * 120)
                    print(f"{'Lớp':<8} | {'Mã lớp':<8} | {'Bắt đầu':<12} | {'Kết thúc':<12} | {'Thứ':<10} | {'Giảng viên':<20} | {'Môn học'}")
                    print("-" * 120)
                    
                    for i, row in df_result.head(15).iterrows():
                        subject = row['Môn học'][:50] + "..." if len(str(row['Môn học'])) > 50 else row['Môn học']
                        print(f"{str(row['Lớp']):<8} | {str(row['Mã lớp']):<8} | {row['Bắt đầu']:<12} | {row['Kết thúc']:<12} | {str(row['Thứ']):<10} | {str(row['Giảng viên'])[:20]:<20} | {subject}")
                    
                    if len(df_result) > 15:
                        print(f"... và {len(df_result) - 15} dòng nữa")
                        
                    # Kiểm tra một vài ví dụ về thời gian
                    print("\n🕐 Kiểm tra thời gian của một vài lớp:")
                    sample_classes = df_result.head(5)
                    for i, row in sample_classes.iterrows():
                        print(f"  {row['Lớp']} - {row['Môn học'][:30]}: {row['Bắt đầu']} đến {row['Kết thúc']}")
            
            else:
                print("Không tìm thấy sheet 'TKB CHINH'")
                
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {e}")
            import traceback
            traceback.print_exc()
            
    else:
        print(f"Không tìm thấy file: {excel_file}")
        print("Các file trong thư mục hiện tại:")
        for file in os.listdir("."):
            if file.endswith((".xlsx", ".xls")):
                print(f"  - {file}")

if __name__ == "__main__":
    main()
