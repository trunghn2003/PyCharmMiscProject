import pandas as pd
import os
from datetime import datetime, timedelta

def convert_day_to_vietnamese(day_num):
    """Chuyển đổi số thứ sang tên thứ tiếng Việt"""
    if pd.isna(day_num):
        return ""
    days = {
        2: "2",
        3: "3", 
        4: "4",
        5: "5",
        6: "6",
        7: "7",
        8: "CN"
    }
    return days.get(int(day_num), f"Thứ {int(day_num)}")

def calculate_date_range_improved(week_schedule):
    """Tính toán ngày bắt đầu và kết thúc dựa trên lịch tuần - CHÍNH XÁC THEO THỰC TẾ"""
    
    # Ánh xạ CHI TIẾT từng cột đến ngày cụ thể dựa trên cấu trúc Excel thực tế
    detailed_week_mapping = {
        # Tháng 8/2025
        '08/25': (2025, 8, 11, 17, 1),      # Tuần 1: 11-17/8
        'Unnamed: 18': (2025, 8, 18, 24, 2), # Tuần 2: 18-24/8  
        'Unnamed: 19': (2025, 8, 25, 31, 3), # Tuần 3: 25-31/8
        
        # Tháng 9/2025
        '09/25': (2025, 9, 1, 7, 4),        # Tuần 4: 1-7/9
        'Unnamed: 21': (2025, 9, 8, 14, 5), # Tuần 5: 8-14/9
        'Unnamed: 22': (2025, 9, 15, 21, 6), # Tuần 6: 15-21/9
        'Unnamed: 23': (2025, 9, 22, 28, 7), # Tuần 7: 22-28/9
        'Unnamed: 24': (2025, 9, 29, 30, 8), # Tuần 8: 29-30/9 (chuyển sang tháng 10)
        
        # Tháng 10/2025
        '10/25': (2025, 10, 6, 12, 9),      # Tuần 9: 6-12/10
        'Unnamed: 26': (2025, 10, 13, 19, 10), # Tuần 10: 13-19/10
        'Unnamed: 27': (2025, 10, 20, 26, 11), # Tuần 11: 20-26/10
        'Unnamed: 28': (2025, 10, 27, 31, 12), # Tuần 12: 27-31/10 (chuyển sang tháng 11)
        
        # Tháng 11/2025
        '11/25': (2025, 11, 3, 9, 13),       # Tuần 13: 3-9/11
        'Unnamed: 30': (2025, 11, 10, 16, 14), # Tuần 14: 10-16/11
        'Unnamed: 31': (2025, 11, 17, 23, 15), # Tuần 15: 17-23/11
        'Unnamed: 32': (2025, 11, 24, 30, 16), # Tuần 16: 24-30/11
        
        # Tháng 12/2025
        '12/25': (2025, 12, 1, 7, 17),       # Tuần 17: 1-7/12
    }
    
    # Tìm tất cả các khoảng thời gian có 'x' và số tuần tương ứng
    active_periods = []
    active_weeks = []  # Danh sách các tuần học
    
    for col_name, (year, month, start_day, end_day, week_num) in detailed_week_mapping.items():
        if col_name in week_schedule:
            value = week_schedule[col_name]
            if pd.notna(value) and 'x' in str(value).lower():
                try:
                    start_date = datetime(year, month, start_day)
                    end_date = datetime(year, month, end_day)
                    active_periods.append((start_date, end_date))
                    active_weeks.append(week_num)  # Thêm số tuần
                except ValueError:
                    continue
    
    if not active_periods:
        return "2025-08-11", "2025-09-14", []  # Trả về ngày mặc định và danh sách tuần trống
    
    # Sắp xếp và lấy ngày đầu tiên và cuối cùng
    active_periods.sort(key=lambda x: x[0])
    overall_start = active_periods[0][0]
    overall_end = active_periods[-1][1]
    
    # Sắp xếp danh sách tuần
    active_weeks.sort()
    
    return overall_start.strftime("%Y-%m-%d"), overall_end.strftime("%Y-%m-%d"), active_weeks

def process_schedule_data_improved(df):
    """Xử lý dữ liệu thời khóa biểu - LẤY TẤT CẢ DỮ LIỆU VÀ THÊM TUẦN HỌC"""
    
    # Không lọc gì cả - lấy TẤT CẢ dữ liệu
    print(f"📊 Xử lý TẤT CẢ {len(df)} dòng dữ liệu gốc")
    
    # Danh sách tất cả các cột tuần (bao gồm cả Unnamed)
    all_week_cols = ['08/25', 'Unnamed: 18', 'Unnamed: 19',
                     '09/25', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23', 'Unnamed: 24',
                     '10/25', 'Unnamed: 26', 'Unnamed: 27', 'Unnamed: 28',
                     '11/25', 'Unnamed: 30', 'Unnamed: 31', 'Unnamed: 32',
                     '12/25']
    
    schedule_list = []
    
    for index, row in df.iterrows():
        try:
            # Lấy TẤT CẢ thông tin, kể cả khi thiếu
            class_name = str(row.get('Lớp', '')).strip() if pd.notna(row.get('Lớp')) else ""
            class_code = str(row.get('Nhóm', '')).strip() if pd.notna(row.get('Nhóm')) else ""
            subject_name = str(row.get('Tên môn học/ học phần', '')).strip() if pd.notna(row.get('Tên môn học/ học phần')) else ""
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
            row_number = str(row.get('TT', '')).strip() if pd.notna(row.get('TT')) else ""
            
            # Chuyển đổi thứ
            day_name = convert_day_to_vietnamese(day_num)
            
            # Tính toán ngày bắt đầu và kết thúc với TẤT CẢ các cột
            week_data = {}
            for week_col in all_week_cols:
                if week_col in df.columns:
                    week_data[week_col] = row.get(week_col)
            
            start_date, end_date, week_numbers = calculate_date_range_improved(week_data)
            
            # Tạo thời gian học
            time_slot = ""
            if pd.notna(start_period) and pd.notna(num_periods):
                try:
                    end_period = int(start_period) + int(num_periods) - 1
                    time_slot = f"{int(start_period)}-{end_period}"
                except:
                    time_slot = f"{start_period}"
            
            # Tạo địa điểm
            location = ""
            if room and building:
                location = f"Phòng {room}, Nhà {building}"
            elif room:
                location = f"Phòng {room}"
            elif building:
                location = f"Nhà {building}"
            
            # Tạo mô tả lịch học từ tất cả cột tuần
            schedule_pattern = ""
            active_weeks = []
            for col in all_week_cols:
                if col in week_data and pd.notna(week_data[col]) and 'x' in str(week_data[col]).lower():
                    active_weeks.append(col)
            if active_weeks:
                schedule_pattern = ", ".join(active_weeks)
            
            # Tạo chuỗi tuần học
            week_list = ""
            if week_numbers:
                week_list = ", ".join([f"{w}" for w in week_numbers])
                
            schedule_item = {
                'STT': row_number,
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
                'Ghi chú': notes,
                'Tuần học': week_list,
            }
            
            # Thêm tất cả các cột tuần gốc để tham khảo
            for col in all_week_cols:
                if col in df.columns:
                    schedule_item[f'Gốc_{col}'] = row.get(col)
            
            schedule_list.append(schedule_item)
            
        except Exception as e:
            print(f"⚠️ Lỗi xử lý dòng {index}: {e}")
            # Vẫn thêm dòng lỗi để không mất dữ liệu
            schedule_item = {
                'STT': index,
                'Lớp': 'Lỗi',
                'Mã lớp': 'Lỗi',
                'Bắt đầu': 'Lỗi',
                'Kết thúc': 'Lỗi',
                'Thứ': 'Lỗi',
                'Giảng viên': 'Lỗi',
                'Môn học': f'Lỗi dòng {index}: {str(e)}',
                'Mã môn học': 'Lỗi',
                'Khóa': 'Lỗi',
                'Ngành': 'Lỗi',
                'Thời gian': 'Lỗi',
                'Địa điểm': 'Lỗi',
                'Số tín chỉ': 'Lỗi',
                'Ghi chú': 'Lỗi',
                'Tuần học': 'Lỗi',
            }
            schedule_list.append(schedule_item)
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

def save_to_excel(df_result, output_file="tkb_full_data.xlsx"):
    """Lưu TẤT CẢ dữ liệu vào file Excel mới"""
    try:
        df_result.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n💾 Đã lưu TẤT CẢ {len(df_result)} dòng dữ liệu vào file: {output_file}")
    except Exception as e:
        print(f"Lỗi khi lưu file: {e}")

def get_period_description_correct(week_data):
    """Tạo mô tả thời gian học dựa trên các tuần có 'x' - ĐÚNG THEO NGÀY THỰC TẾ"""
    period_descriptions = []
    
    week_periods = {
        '08/25': "tháng 8/2025 (ngày 11-31)",
        '09/25': "tháng 9/2025 (ngày 1-28)", 
        '10/25': "tháng 9-10/2025 (ngày 29/9-26/10)",
        '11/25': "tháng 10-11/2025 (ngày 27/10-23/11)",
        '12/25': "tháng 11-12/2025 (ngày 24/11-7/12)"
    }
    
    for week_col, description in week_periods.items():
        if str(week_data.get(week_col, '')).lower() == 'x':
            period_descriptions.append(description)
    
    if period_descriptions:
        return "Từ " + ", ".join(period_descriptions)
    return "Không xác định"

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
                    
                    # Xuất mẫu dữ liệu theo format yêu cầu BÌNH THƯỜNG
                    print("\n📄 BẢNG THỜI KHÓA BIỂU THEO ĐỊNH DẠNG YÊU CẦU:")
                    print("-" * 150)
                    print(f"{'Lớp':<8} | {'Mã lớp':<8} | {'Bắt đầu':<12} | {'Kết thúc':<12} | {'Thứ':<10} | {'Tuần học':<25} | {'Giảng viên':<20} | {'Môn học'}")
                    print("-" * 150)
                    
                    for i, row in df_result.head(15).iterrows():
                        subject = row['Môn học'][:40] + "..." if len(str(row['Môn học'])) > 40 else row['Môn học']
                        weeks = row['Tuần học'][:23] + "..." if len(str(row['Tuần học'])) > 23 else row['Tuần học']
                        print(f"{str(row['Lớp']):<8} | {str(row['Mã lớp']):<8} | {row['Bắt đầu']:<12} | {row['Kết thúc']:<12} | {str(row['Thứ']):<10} | {weeks:<25} | {str(row['Giảng viên'])[:20]:<20} | {subject}")
                    
                    if len(df_result) > 15:
                        print(f"... và {len(df_result) - 15} dòng nữa")
                        
                    # Kiểm tra một vài ví dụ về thời gian VÀ TUẦN HỌC
                    print("\n🕐 Kiểm tra thời gian và tuần học của một vài lớp:")
                    sample_classes = df_result.head(5)
                    for i, row in sample_classes.iterrows():
                        print(f"  {row['Lớp']} - {row['Môn học'][:30]}: {row['Bắt đầu']} đến {row['Kết thúc']}")
                        print(f"    ➤ {row['Tuần học']}")
                        print()
            
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
