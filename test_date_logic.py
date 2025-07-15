import pandas as pd
from datetime import datetime, timedelta

def calculate_date_range_correct(week_schedule):
    """Tính toán ngày bắt đầu và kết thúc dựa trên lịch tuần - THEO NGÀY THỰC TẾ"""
    
    # Các tuần trong học kỳ với ngày bắt đầu và kết thúc THỰC TẾ theo file Excel
    week_periods = {
        '08/25': {'start': datetime(2025, 8, 11), 'end': datetime(2025, 8, 17)},    # Tuần 11-17/8
        '09/25': {'start': datetime(2025, 8, 18), 'end': datetime(2025, 8, 24)},    # Tuần 18-24/8  
        '10/25': {'start': datetime(2025, 8, 25), 'end': datetime(2025, 8, 31)},    # Tuần 25-31/8
        '11/25': {'start': datetime(2025, 9, 1), 'end': datetime(2025, 9, 7)},      # Tuần 1-7/9
        '12/25': {'start': datetime(2025, 9, 8), 'end': datetime(2025, 9, 14)},     # Tuần 8-14/9
    }
    
    # Tìm tuần đầu tiên và cuối cùng có dấu 'x'
    first_week_start = None
    last_week_end = None
    
    for week_col, period in week_periods.items():
        if str(week_schedule.get(week_col, '')).lower() == 'x':
            if first_week_start is None:
                first_week_start = period['start']
            last_week_end = period['end']
    
    if first_week_start and last_week_end:
        return first_week_start.strftime("%Y-%m-%d"), last_week_end.strftime("%Y-%m-%d")
    
    return "2025-08-11", "2025-09-14"

def get_period_description_correct(week_data):
    """Tạo mô tả thời gian học dựa trên các tuần có 'x' - ĐÚNG THEO NGÀY THỰC TẾ"""
    period_descriptions = []
    
    week_periods = {
        '08/25': "tuần 11-17/8/2025",
        '09/25': "tuần 18-24/8/2025", 
        '10/25': "tuần 25-31/8/2025",
        '11/25': "tuần 1-7/9/2025",
        '12/25': "tuần 8-14/9/2025"
    }
    
    for week_col, description in week_periods.items():
        if str(week_data.get(week_col, '')).lower() == 'x':
            period_descriptions.append(description)
    
    if period_descriptions:
        return "Từ " + ", ".join(period_descriptions)
    return "Không xác định"

# Test các trường hợp
print("🧪 KIỂM TRA LOGIC TÍNH NGÀY MỚI:")
print("=" * 50)

test_cases = [
    # Test case 1: Chỉ có tuần đầu
    {'08/25': 'x', '09/25': '', '10/25': '', '11/25': '', '12/25': ''},
    # Test case 2: Có 2 tuần liên tiếp
    {'08/25': 'x', '09/25': 'x', '10/25': '', '11/25': '', '12/25': ''},
    # Test case 3: Có 3 tuần
    {'08/25': 'x', '09/25': 'x', '10/25': 'x', '11/25': '', '12/25': ''},
    # Test case 4: Có tuần giữa và cuối
    {'08/25': '', '09/25': 'x', '10/25': '', '11/25': 'x', '12/25': ''},
    # Test case 5: Tất cả các tuần
    {'08/25': 'x', '09/25': 'x', '10/25': 'x', '11/25': 'x', '12/25': 'x'},
]

for i, test_case in enumerate(test_cases, 1):
    start_date, end_date = calculate_date_range_correct(test_case)
    description = get_period_description_correct(test_case)
    
    print(f"Test {i}:")
    print(f"  Input: {test_case}")
    print(f"  Ngày: {start_date} đến {end_date}")
    print(f"  Mô tả: {description}")
    print("-" * 30)
