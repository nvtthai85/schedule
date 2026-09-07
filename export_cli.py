#!/usr/bin/env python3
"""
CLI Tool: Xuất Kế Hoạch Kiểm Soát FECT
Cách dùng:
    python3 export_cli.py --month 8 --output Ke_hoach_T8.xlsx
"""
import argparse
import os
import sys

from excel_processor import generate_control_plan_workbook, DEFAULT_MATRIX_PATH, DEFAULT_TEMPLATE_PATH

def main():
    parser = argparse.ArgumentParser(description="Tạo file Kế hoạch kiểm soát FECT từ file Ma trận")
    parser.add_argument("-m", "--month", type=int, required=True, help="Tháng cần lập lịch (1 - 12)")
    parser.add_argument("-y", "--year", type=int, default=2026, help="Năm (mặc định: 2026)")
    parser.add_argument("-i", "--matrix", type=str, default=DEFAULT_MATRIX_PATH, help="Đường dẫn file ma trận")
    parser.add_argument("-t", "--template", type=str, default=DEFAULT_TEMPLATE_PATH, help="Đường dẫn file template")
    parser.add_argument("-o", "--output", type=str, default=None, help="Tên file xuất ra")
    parser.add_argument("--fsc-all", action="store_true", help="Tạo đủ 3 sheet FSC CT, FSC ST, FSC HG")
    parser.add_argument("--no-calendar", action="store_true", help="Không tạo sheet Lịch tổng quan")
    parser.add_argument("--skip-empty", action="store_true", help="Bỏ qua các sheet không có nhiệm vụ")
    parser.add_argument("--no-schedule", action="store_true", help="Không tự động xếp lịch (để trống Cột H, I, J)")

    args = parser.parse_args()

    if args.month < 1 or args.month > 12:
        print("❌ Lỗi: Tháng phải từ 1 đến 12")
        sys.exit(1)

    month_str = f"T{args.month:02d}"
    output_filename = args.output or f"Ke_hoach_kiem_soat_{month_str}_{args.year}.xlsx"

    print(f"📊 Đang xử lý trích xuất Ma trận cho Tháng {args.month}/{args.year}...")
    print(f"   - File ma trận: {os.path.basename(args.matrix)}")
    print(f"   - File template: {os.path.basename(args.template)}")
    print(f"   - Tự động xếp lịch: {'Tắt' if args.no_schedule else 'Bật (6-23 hàng tháng, T2-T6, <=4 việc/buổi)'}")

    stream = generate_control_plan_workbook(
        matrix_source=args.matrix,
        template_source=args.template,
        month_num=args.month,
        year=args.year,
        fsc_mode="all_fsc" if args.fsc_all else "fsc_ct",
        include_calendar=not args.no_calendar,
        include_empty_sheets=not args.skip_empty,
        auto_schedule=not args.no_schedule,
    )

    with open(output_filename, "wb") as f:
        f.write(stream.getvalue())

    print(f"✅ Đã tạo thành công file: {output_filename}")
    print(f"   Kích thước: {os.path.getsize(output_filename) / 1024:.1f} KB")

if __name__ == "__main__":
    main()
