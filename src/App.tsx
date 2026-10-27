import React, { useState, useCallback, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, Download, Check, Search } from 'lucide-react';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Label, Cell, Legend } from 'recharts';
import { cn } from './lib/utils';

// --- Interfaces for Dashboard Data ---
interface ScoreRangeCount {
  name: string;
  count: number;
}

interface DashboardData {
  subjects: string[];
  classes: string[];
  overall: Record<string, ScoreRangeCount[]>;
  byClass: Record<string, Record<string, ScoreRangeCount[]>>;
  byClassGrouped: Record<string, any[]>;
  overallAvg: Record<string, number>;
  students: any[];
  blocks: Record<string, string[]>;
}
// -------------------------------------

const getPythonCode = (kyThi: string, namHoc: string, khoiLop: string) => `import pandas as pd
import numpy as np
from google.colab import files
import io
import os
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

def process_scores():
    print("Vui lòng tải lên các file Excel điểm môn học (VD: Toan.xlsx, Van.xlsx,...)")
    uploaded = files.upload()
    
    if not uploaded:
        print("Chưa có file nào được tải lên.")
        return

    student_data = {}
    subjects = []

    for filename, content in uploaded.items():
        if not filename.endswith('.xlsx') and not filename.endswith('.xls'):
            continue
            
        subject_name = os.path.splitext(filename)[0]
        subjects.append(subject_name)
        
        # Tự động tìm dòng header (mẫu THCS Mỹ Xá header ở dòng 5, index 4)
        header_idx = 4
        col_hoten = 1    # Cột B
        col_ngaysinh = 2 # Cột C
        col_sbd = 6      # Cột G
        col_diem = 8     # Cột I
        col_lop = 9      # Cột J (Ghi chú)

        try:
            temp_df = pd.read_excel(io.BytesIO(content), header=None, nrows=10)
            for r_idx, r in temp_df.iterrows():
                row_str = " ".join([str(val).lower() for val in r.values if pd.notna(val)])
                if 'họ và tên' in row_str or 'số bd' in row_str or 'sbd' in row_str:
                    header_idx = r_idx
                    for c_idx, c_val in enumerate(r.values):
                        c_str = str(c_val).lower().strip()
                        if 'họ và tên' in c_str or c_str == 'họ tên': col_hoten = c_idx
                        elif 'ngày sinh' in c_str: col_ngaysinh = c_idx
                        elif 'số bd' in c_str or c_str == 'sbd': col_sbd = c_idx
                        elif 'điểm' in c_str and 'địa điểm' not in c_str: col_diem = c_idx
                        elif 'ghi chú' in c_str or c_str == 'lớp': col_lop = c_idx
                    break
            df = pd.read_excel(io.BytesIO(content), header=header_idx)
        except Exception as e:
            print(f"Lỗi đọc file {filename}: {e}")
            continue
            
        for index, row in df.iterrows():
            try:
                sbd = str(row.iloc[col_sbd]).strip() # Cột G - Số BD
                if sbd.endswith('.0') and sbd[:-2].isdigit():
                    sbd = sbd[:-2]
                ho_ten = str(row.iloc[col_hoten]).strip() # Cột B - Họ tên
                
                if pd.isna(sbd) or sbd == 'nan' or sbd == '' or '...' in sbd:
                    continue
                if pd.isna(ho_ten) or ho_ten == 'nan' or ho_ten == '':
                    continue
                    
                if sbd not in student_data:
                    # Xử lý ngày sinh
                    ngay_sinh = row.iloc[col_ngaysinh]
                    if isinstance(ngay_sinh, pd.Timestamp):
                        ngay_sinh = ngay_sinh.strftime('%d/%m/%Y')
                    elif pd.isna(ngay_sinh) or str(ngay_sinh) == 'nan':
                        ngay_sinh = ''
                    else:
                        ngay_sinh = str(ngay_sinh).strip()
                        
                    lop_val = row.iloc[col_lop]
                    lop = '' if pd.isna(lop_val) or str(lop_val) == 'nan' else str(lop_val).strip()

                    student_data[sbd] = {
                        'Họ và tên': ho_ten,
                        'Ngày sinh': ngay_sinh,
                        'Số BD': sbd,
                        'Lớp': lop,
                        'scores': {}
                    }
                
                score = row.iloc[col_diem] # Cột I - Điểm
                if pd.notna(score) and str(score).strip() not in ['', 'nan', '-', 'v', 'V']:
                    val_str = str(score).strip().replace(',', '.')
                    try:
                        student_data[sbd]['scores'][subject_name] = float(val_str)
                    except ValueError:
                        pass
            except Exception as e:
                continue

    # Lọc bỏ học sinh không có điểm nào
    student_data = {sbd: info for sbd, info in student_data.items() if len(info['scores']) > 0}

    if not student_data:
        print("Không tìm thấy dữ liệu hợp lệ trong các file.")
        return

    # Sắp xếp môn học
    subject_order = ['toán', 'toan', 'văn', 'van', 'anh', 'lí', 'li', 'lý', 'ly', 'hoá', 'hoa', 'hóa', 'sinh', 'sử', 'su', 'địa', 'dia']
    def get_subject_index(sub):
        s = sub.lower()
        if s in subject_order:
            return subject_order.index(s)
        return 999
    subjects.sort(key=lambda x: (get_subject_index(x), x))

    # Chuyển đổi thành DataFrame
    data_list = []
    for sbd, info in student_data.items():
        row_data = {
            'Họ và tên': info['Họ và tên'],
            'Ngày sinh': info['Ngày sinh'],
            'SBD': info['Số BD'],
            'Lớp': info['Lớp']
        }
        row_data.update(info['scores'])
        data_list.append(row_data)

    df_final = pd.DataFrame(data_list)
    
    # Tính Tổng điểm và TBM
    df_final['Tổng điểm'] = df_final[subjects].sum(axis=1, skipna=True)
    # TBM chỉ tính trên các môn có điểm
    df_final['TBM'] = df_final[subjects].mean(axis=1, skipna=True).round(2)
    
    # Xếp thứ (XT) theo Tổng điểm
    df_final['XT'] = df_final['Tổng điểm'].rank(method='min', ascending=False).astype('Int64')

    # Sắp xếp theo SBD
    df_final = df_final.sort_values(by=['SBD'])
    df_final.insert(0, 'STT', range(1, len(df_final) + 1))

    # Sắp xếp các môn học theo thứ tự yêu cầu
    subject_order = ['toán', 'văn', 'anh', 'khtn', 'lí', 'lý', 'hoá', 'hóa', 'sinh', 'sử', 'địa', 'gdcd', 'tin', 'công nghệ']
    def get_subject_weight(sub):
        sub_lower = sub.lower()
        for idx, s in enumerate(subject_order):
            if sub_lower == s or sub_lower.startswith(s):
                return idx
        return 999
    subjects.sort(key=get_subject_weight)

    # Sắp xếp lại cột (Bảng điểm chung không tính điểm các khối)
    cols = ['STT', 'Họ và tên', 'Ngày sinh', 'SBD', 'Lớp'] + subjects + ['Tổng điểm', 'XT', 'TBM']
    df_final = df_final[[c for c in cols if c in df_final.columns]]

    # Thêm dòng ĐIỂM TRUNG BÌNH
    avg_row = {col: '' for col in df_final.columns}
    avg_row['Họ và tên'] = 'ĐIỂM TRUNG BÌNH'
    cols_to_average = subjects + ['Tổng điểm', 'TBM']
    
    df_for_avg = df_final[df_final['Lớp'] != 'TSTD']
    df_for_avg = df_for_avg[df_for_avg['Họ và tên'] != 'Đặng Mai Minh Hoàng']
        
    for col in cols_to_average:
        if col in df_for_avg.columns and df_for_avg[col].notna().any():
            avg_row[col] = round(df_for_avg[col].mean(), 2)
    
    df_final = pd.concat([df_final, pd.DataFrame([avg_row])], ignore_index=True)

    # Tạo workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng Điểm Tổng Hợp"

    # Header dòng 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    ws['A1'] = f"BẢNG ĐIỂM {kyThi.upper()} {khoiLop.upper()} - NĂM HỌC {namHoc}"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    # Header dòng 3 (Tên cột)
    header_row = 3
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    for col_num, col_name in enumerate(df_final.columns, 1):
        cell = ws.cell(row=header_row, column=col_num, value=col_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.fill = header_fill
        cell.border = thin_border

    # Ghi dữ liệu
    for r_idx, row in enumerate(df_final.values, header_row + 1):
        for c_idx, value in enumerate(row, 1):
            # Xử lý NaN
            if pd.isna(value):
                val = ""
            else:
                val = value
                
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            
            # Căn giữa cho các cột không phải Họ tên
            if df_final.columns[c_idx-1] not in ['Họ và tên']:
                cell.alignment = Alignment(horizontal='center')

    # Auto-fit columns
    for col in ws.columns:
        column_letter = col[0].column_letter
        header_val = ws[f'{column_letter}3'].value
        
        if header_val == 'STT': width = 5
        elif header_val == 'Họ và tên': width = 25
        elif header_val == 'Ngày sinh': width = 12
        elif header_val == 'SBD': width = 12
        elif header_val == 'Lớp': width = 8
        elif header_val in subjects: width = 6
        elif header_val == 'Tổng điểm': width = 10
        elif header_val == 'TBM': width = 6
        elif str(header_val).startswith('Khối'): width = 8
        elif str(header_val) == 'XT': width = 5
        else: width = 8
        
        ws.column_dimensions[column_letter].width = width

    # --- Sheet Top ---
    top_limit = 15 if "${khoiLop}" in ["Khối 6", "Khối 7", "Khối 8", "Khối 10", "Khối 11"] else 5
    ws_top5 = wb.create_sheet(title=f"Top {top_limit}")
    
    ws_top5.merge_cells('A1:L1')
    ws_top5['A1'] = f"DANH SÁCH HỌC SINH - TOP {{top_limit}} - {kyThi.upper()} {khoiLop.upper()} - NĂM HỌC {namHoc}"
    ws_top5['A1'].font = Font(bold=True, size=12)
    ws_top5['A1'].border = thin_border

    current_row = 3

    def render_top5_table(title, score_col, subject_cols, limit):
        nonlocal current_row
        
        if score_col not in df_final.columns:
            return
            
        df_valid = df_final[df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH'].dropna(subset=[score_col])
        if df_valid.empty:
            return
            
        top5 = df_valid.sort_values(by=score_col, ascending=False).head(limit)
        
        cell = ws_top5.cell(row=current_row, column=1, value=title)
        cell.font = Font(bold=True)
        current_row += 1
        
        headers = ['STT', 'Họ và tên', 'Số BD', 'Lớp'] + subject_cols + ['Điểm']
        for c_idx, header in enumerate(headers, 1):
            cell = ws_top5.cell(row=current_row, column=c_idx, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        current_row += 1
        
        for i, (_, row_data) in enumerate(top5.iterrows(), 1):
            data = [i, row_data['Họ và tên'], row_data['SBD'], row_data['Lớp']]
            for sub in subject_cols:
                data.append(row_data.get(sub, ''))
            data.append(row_data[score_col])
            
            for c_idx, val in enumerate(data, 1):
                cell = ws_top5.cell(row=current_row, column=c_idx, value=val if pd.notna(val) else '')
                cell.border = thin_border
                if c_idx != 2:
                    cell.alignment = Alignment(horizontal='center')
            current_row += 1
            
        current_row += 1

    render_top5_table(f'TOP {top_limit} HỌC SINH CAO ĐIỂM NHẤT - TỔNG ĐIỂM', 'Tổng điểm', subjects, top_limit)

    for i, col in enumerate(ws_top5.columns):
        column_letter = col[0].column_letter
        if i == 0: width = 5
        elif i == 1: width = 25
        elif i == 2: width = 12
        elif i == 3: width = 8
        elif i < 4 + len(subjects): width = 6
        elif i == 4 + len(subjects): width = 8
        else: width = 8
        ws_top5.column_dimensions[column_letter].width = width

    # --- Sheets for each class ---
    classes = df_final[df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH']['Lớp'].dropna().unique()
    classes = sorted([str(c) for c in classes])

    def get_score_range(score):
        if pd.isna(score): return None
        if score < 1.0: return '<1,0'
        if score < 2.0: return '1,0-<2,0'
        if score < 3.0: return '2,0-<3,0'
        if score < 4.0: return '3,0-<4,0'
        if score < 5.0: return '4,0-<5,0'
        if score < 6.0: return '5,0-<6,0'
        if score < 7.0: return '6,0-<7,0'
        if score < 8.0: return '7,0-<8,0'
        if score < 9.0: return '8,0-<9,0'
        if score < 10.0: return '9,0-<10,0'
        return '= 10,0'

    ranges = ['<1,0', '1,0-<2,0', '2,0-<3,0', '3,0-<4,0', '4,0-<5,0', '5,0-<6,0', '6,0-<7,0', '7,0-<8,0', '8,0-<9,0', '9,0-<10,0', '= 10,0']

    # --- Sheet TBDIEM ---
    ws_tbdiem = wb.create_sheet(title="TBDIEM")
    
    # Calculate class averages
    class_avgs = []
    cols_to_avg = ['TBM'] + subjects + ['Tổng điểm']
    
    for cls in classes:
        df_cls = df_final[(df_final['Lớp'] == cls) & (df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH')]
        df_cls = df_cls[df_cls['Họ và tên'] != 'Đặng Mai Minh Hoàng']
        avg_data = {'Lớp': cls}
        for col in cols_to_avg:
            if col in df_cls.columns and df_cls[col].notna().any():
                avg_data[col] = round(df_cls[col].mean(), 2)
            else:
                avg_data[col] = None
        class_avgs.append(avg_data)
        
    df_tbdiem = pd.DataFrame(class_avgs)
    
    # Calculate ranks
    for col in cols_to_avg:
        if col in df_tbdiem.columns and df_tbdiem[col].notna().any():
            df_tbdiem[f'XT_{col}'] = df_tbdiem[col].rank(method='min', ascending=False).astype('Int64')
        else:
            df_tbdiem[f'XT_{col}'] = None
            
    # Header dòng 1
    ws_tbdiem.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(cols_to_avg)*2)
    ws_tbdiem['A1'] = f"BẢNG TỔNG HỢP ĐIỂM TRUNG BÌNH CÁC LỚP - {kyThi.upper()} {khoiLop.upper()} - NĂM HỌC {namHoc}"
    ws_tbdiem['A1'].font = Font(bold=True, size=14)
    ws_tbdiem['A1'].alignment = Alignment(vertical='center')

    # Header dòng 3
    headers_tbdiem = ['Lớp']
    for col in cols_to_avg:
        headers_tbdiem.extend([col, 'XT'])
        
    for c_idx, header in enumerate(headers_tbdiem, 1):
        cell = ws_tbdiem.cell(row=3, column=c_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        
    # Data rows
    current_row = 4
    for _, row_data in df_tbdiem.iterrows():
        ws_tbdiem.cell(row=current_row, column=1, value=row_data['Lớp']).border = thin_border
        c_idx = 2
        for col in cols_to_avg:
            val = row_data.get(col)
            xt_val = row_data.get(f'XT_{col}')
            
            cell_val = ws_tbdiem.cell(row=current_row, column=c_idx, value=val if pd.notna(val) else '')
            cell_val.border = thin_border
            cell_val.alignment = Alignment(horizontal='center')
            
            cell_xt = ws_tbdiem.cell(row=current_row, column=c_idx+1, value=xt_val if pd.notna(xt_val) else '')
            cell_xt.border = thin_border
            cell_xt.alignment = Alignment(horizontal='center')
            
            c_idx += 2
        current_row += 1
        
    # TB Khối row
    ws_tbdiem.cell(row=current_row, column=1, value='TB Khối').font = Font(bold=True)
    ws_tbdiem.cell(row=current_row, column=1).border = thin_border
    
    df_students = df_final[df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH']
    df_students = df_students[df_students['Họ và tên'] != 'Đặng Mai Minh Hoàng']
    c_idx = 2
    for col in cols_to_avg:
        if col in df_students.columns and df_students[col].notna().any():
            overall_avg = round(df_students[col].mean(), 2)
        else:
            overall_avg = ''
            
        cell_val = ws_tbdiem.cell(row=current_row, column=c_idx, value=overall_avg)
        cell_val.font = Font(bold=True)
        cell_val.border = thin_border
        cell_val.alignment = Alignment(horizontal='center')
        
        cell_xt = ws_tbdiem.cell(row=current_row, column=c_idx+1, value='')
        cell_xt.border = thin_border
        
        c_idx += 2

    # Auto-fit columns
    for col in ws_tbdiem.columns:
        column_letter = col[0].column_letter
        header_val = ws_tbdiem[f'{column_letter}3'].value
        if header_val == 'Lớp': width = 8
        elif header_val == 'XT': width = 5
        else: width = 8
        ws_tbdiem.column_dimensions[column_letter].width = width

    # --- Sheet TKSo ---
    ws_tkso = wb.create_sheet(title="TKSo")
    
    tk_ranges = [
        {'label': 'Điểm 0', 'check': lambda s: s == 0},
        {'label': '>= 0 ÷ < 1', 'check': lambda s: 0 < s < 1},
        {'label': '>= 1 ÷ < 2', 'check': lambda s: 1 <= s < 2},
        {'label': '>= 2 ÷ < 3', 'check': lambda s: 2 <= s < 3},
        {'label': '>= 3 ÷ < 4', 'check': lambda s: 3 <= s < 4},
        {'label': '>= 4 ÷ < 5', 'check': lambda s: 4 <= s < 5},
        {'label': '>= 5 ÷ < 6', 'check': lambda s: 5 <= s < 6},
        {'label': '>= 6 ÷ < 7', 'check': lambda s: 6 <= s < 7},
        {'label': '>= 7 ÷ < 8', 'check': lambda s: 7 <= s < 8},
        {'label': '>= 8 ÷ < 9', 'check': lambda s: 8 <= s < 9},
        {'label': '>= 9 ÷ <= 10', 'check': lambda s: 9 <= s <= 10},
        {'label': 'Điểm 10', 'check': lambda s: s == 10}
    ]
    
    # Header dòng 1
    ws_tkso.cell(row=1, column=3, value="PHỔ ĐIỂM").font = Font(bold=True)
    
    # Header dòng 2 & 3
    ws_tkso.cell(row=3, column=1, value="Môn").font = Font(bold=True)
    ws_tkso.cell(row=3, column=1).border = thin_border
    ws_tkso.cell(row=3, column=1).alignment = Alignment(horizontal='center', vertical='center')
    
    ws_tkso.cell(row=3, column=2, value="Tổng số HS KS").font = Font(bold=True)
    ws_tkso.cell(row=3, column=2).border = thin_border
    ws_tkso.cell(row=3, column=2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    c_idx = 3
    for r in tk_ranges:
        ws_tkso.merge_cells(start_row=2, start_column=c_idx, end_row=2, end_column=c_idx+1)
        cell_label = ws_tkso.cell(row=2, column=c_idx, value=r['label'])
        cell_label.border = thin_border
        cell_label.alignment = Alignment(horizontal='center', vertical='center')
        ws_tkso.cell(row=2, column=c_idx+1).border = thin_border
        
        cell_sl = ws_tkso.cell(row=3, column=c_idx, value="SL")
        cell_sl.border = thin_border
        cell_sl.alignment = Alignment(horizontal='center', vertical='center')
        
        cell_tl = ws_tkso.cell(row=3, column=c_idx+1, value="TL %")
        cell_tl.border = thin_border
        cell_tl.alignment = Alignment(horizontal='center', vertical='center')
        
        c_idx += 2
        
    # Data rows
    current_row = 4
    df_students = df_final[df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH']
    
    for sub in subjects:
        if sub not in df_students.columns:
            continue
            
        scores = df_students[sub].dropna()
        total_hs = len(scores)
        
        ws_tkso.cell(row=current_row, column=1, value=sub.capitalize()).border = thin_border
        ws_tkso.cell(row=current_row, column=2, value=total_hs).border = thin_border
        
        c_idx = 3
        for r in tk_ranges:
            count = sum(1 for s in scores if r['check'](s))
            tl = round((count / total_hs * 100), 2) if total_hs > 0 else 0
            
            cell_sl = ws_tkso.cell(row=current_row, column=c_idx, value=count)
            cell_sl.border = thin_border
            cell_sl.alignment = Alignment(horizontal='center')
            
            cell_tl = ws_tkso.cell(row=current_row, column=c_idx+1, value=tl)
            cell_tl.border = thin_border
            cell_tl.alignment = Alignment(horizontal='center')
            
            c_idx += 2
            
        current_row += 1
        
    # Auto-fit columns
    for col in ws_tkso.columns:
        column_letter = col[0].column_letter
        if column_letter == 'A': width = 12
        elif column_letter == 'B': width = 15
        else: width = 8
        ws_tkso.column_dimensions[column_letter].width = width

    # --- Sheet DSTruot ---
    ws_dstruot = wb.create_sheet(title="DSTruot")
    
    # Filter students with any score < 5
    dstruot_cols = ['STT', 'Họ và tên', 'Ngày sinh', 'SBD', 'Lớp'] + subjects + ['Tổng điểm', 'TBM']
    
    # Create a mask for students with at least one subject < 5
    mask = pd.Series(False, index=df_final.index)
    for sub in subjects:
        if sub in df_final.columns:
            mask = mask | (df_final[sub] < 5)
            
    df_dstruot = df_final[mask & (df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH')].copy()
    
    # Re-calculate STT for this sheet
    df_dstruot['STT'] = range(1, len(df_dstruot) + 1)
    
    # Header dòng 1
    ws_dstruot.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(dstruot_cols))
    ws_dstruot['A1'] = f"DANH SÁCH HỌC SINH CÓ ĐIỂM DƯỚI 5 - {kyThi.upper()} {khoiLop.upper()} - NĂM HỌC {namHoc}"
    ws_dstruot['A1'].font = Font(bold=True, size=12)
    
    # Header dòng 3
    for c_idx, col_name in enumerate(dstruot_cols, 1):
        cell = ws_dstruot.cell(row=3, column=c_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    # Data rows
    for r_idx, row in enumerate(df_dstruot[dstruot_cols].values, 4):
        for c_idx, value in enumerate(row, 1):
            val = "" if pd.isna(value) else value
            cell = ws_dstruot.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if dstruot_cols[c_idx-1] not in ['Họ và tên']:
                cell.alignment = Alignment(horizontal='center')
                
    # Auto-fit columns
    for col in ws_dstruot.columns:
        column_letter = col[0].column_letter
        header_val = ws_dstruot[f'{column_letter}3'].value
        if header_val == 'STT': width = 5
        elif header_val == 'Họ và tên': width = 25
        elif header_val == 'Ngày sinh': width = 12
        elif header_val == 'SBD': width = 12
        elif header_val == 'Lớp': width = 8
        elif header_val in subjects: width = 6
        elif header_val == 'Tổng điểm': width = 10
        elif header_val == 'TBM': width = 6
        else: width = 8
        ws_dstruot.column_dimensions[column_letter].width = width

    # --- Sheet TKTruot ---
    ws_tktruot = wb.create_sheet(title="TKTruot")
    
    tktruot_cols = ['Lớp'] + subjects + ['Tổng số HS trượt']
    
    # Header dòng 1
    ws_tktruot.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(tktruot_cols))
    ws_tktruot['A1'] = f"BẢNG THỐNG KÊ SỐ LƯỢNG HỌC SINH CÓ ĐIỂM DƯỚI 5 THEO LỚP - {kyThi.upper()} {khoiLop.upper()} - NĂM HỌC {namHoc}"
    ws_tktruot['A1'].font = Font(bold=True, size=12)
    
    # Header dòng 3
    for c_idx, col_name in enumerate(tktruot_cols, 1):
        cell = ws_tktruot.cell(row=3, column=c_idx, value=col_name.capitalize() if col_name in subjects else col_name)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    current_row = 4
    df_valid = df_final[(df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH') & (df_final['Lớp'] != 'TSTD')]
    
    total_row_data = {col: 0 for col in subjects}
    total_hs_truot_all = 0
    
    for cls in classes:
        if cls == 'TSTD': continue
        df_cls = df_valid[df_valid['Lớp'] == cls]
        
        ws_tktruot.cell(row=current_row, column=1, value=cls).border = thin_border
        
        # Count students with < 5 in any subject for this class
        mask_cls = pd.Series(False, index=df_cls.index)
        for sub in subjects:
            if sub in df_cls.columns:
                mask_cls = mask_cls | (df_cls[sub] < 5)
        
        hs_truot_cls = mask_cls.sum()
        total_hs_truot_all += hs_truot_cls
        
        for c_idx, sub in enumerate(subjects, 2):
            if sub in df_cls.columns:
                count_sub = (df_cls[sub] < 5).sum()
            else:
                count_sub = 0
            
            total_row_data[sub] += count_sub
            
            cell = ws_tktruot.cell(row=current_row, column=c_idx, value=count_sub)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
            
        cell_total = ws_tktruot.cell(row=current_row, column=len(tktruot_cols), value=hs_truot_cls)
        cell_total.border = thin_border
        cell_total.alignment = Alignment(horizontal='center')
        
        current_row += 1
        
    # TỔNG CỘNG row
    cell_tc = ws_tktruot.cell(row=current_row, column=1, value='TỔNG CỘNG')
    cell_tc.font = Font(bold=True)
    cell_tc.border = thin_border
    
    for c_idx, sub in enumerate(subjects, 2):
        cell = ws_tktruot.cell(row=current_row, column=c_idx, value=total_row_data[sub])
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')
        
    cell_total_all = ws_tktruot.cell(row=current_row, column=len(tktruot_cols), value=total_hs_truot_all)
    cell_total_all.font = Font(bold=True)
    cell_total_all.border = thin_border
    cell_total_all.alignment = Alignment(horizontal='center')
    
    # Auto-fit columns
    for col in ws_tktruot.columns:
        column_letter = col[0].column_letter
        header_val = ws_tktruot[f'{column_letter}3'].value
        if header_val == 'Lớp': width = 8
        elif header_val == 'Tổng số HS trượt': width = 15
        else: width = 8
        ws_tktruot.column_dimensions[column_letter].width = width

    zip_filename = "Danh_sach_cac_lop.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for cls in classes:
            # Create sheet in main workbook
            ws_cls = wb.create_sheet(title=cls)
            
            # Create individual workbook for the class
            wb_ind = Workbook()
            ws_ind = wb_ind.active
            ws_ind.title = cls
            
            df_cls = df_final[(df_final['Lớp'] == cls) & (df_final['Họ và tên'] != 'ĐIỂM TRUNG BÌNH')].copy()
            
            # Determine active subjects for this class
            active_subjects = [sub for sub in subjects if sub in df_cls.columns and df_cls[sub].notna().any()]
            
            class_cols = ['STT', 'Họ và tên', 'Ngày sinh', 'SBD', 'Lớp'] + active_subjects + ['Tổng điểm', 'XT', 'TBM']
            class_cols = [c for c in class_cols if c in df_cls.columns or c == 'STT']
            
            # Xếp thứ lại theo từng lớp
            if 'Tổng điểm' in df_cls.columns:
                df_cls['XT'] = df_cls['Tổng điểm'].rank(method='min', ascending=False).astype('Int64')
            
            def build_class_sheet(sheet):
                # Header dòng 1
                sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(class_cols))
                sheet['A1'] = f"BẢNG ĐIỂM {kyThi.upper()} LỚP {{cls}} - NĂM HỌC {namHoc}"
                sheet['A1'].font = Font(bold=True, size=14)
                sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')

                # Header dòng 3
                for col_num, col_name in enumerate(class_cols, 1):
                    cell = sheet.cell(row=3, column=col_num, value=col_name)
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.fill = header_fill
                    cell.border = thin_border

                # Ghi dữ liệu học sinh
                current_row = 4
                df_cls_filtered = df_cls[class_cols]
                for r_idx, row in enumerate(df_cls_filtered.values, current_row):
                    for c_idx, value in enumerate(row, 1):
                        val = "" if pd.isna(value) else value
                        cell = sheet.cell(row=r_idx, column=c_idx, value=val)
                        cell.border = thin_border
                        if class_cols[c_idx-1] not in ['Họ và tên']:
                            cell.alignment = Alignment(horizontal='center')
                    current_row += 1

                # Dòng ĐIỂM TRUNG BÌNH
                avg_row = {col: '' for col in class_cols}
                avg_row['Họ và tên'] = 'ĐIỂM TRUNG BÌNH'
                cols_to_average = active_subjects + ['Tổng điểm', 'TBM']
                
                df_cls_avg = df_cls.copy()
                df_cls_avg = df_cls_avg[df_cls_avg['Họ và tên'] != 'Đặng Mai Minh Hoàng']

                for col in cols_to_average:
                    if col in df_cls_avg.columns and df_cls_avg[col].notna().any():
                        avg_row[col] = round(df_cls_avg[col].mean(), 2)
                
                for c_idx, col_name in enumerate(class_cols, 1):
                    cell = sheet.cell(row=current_row, column=c_idx, value=avg_row[col_name])
                    cell.font = Font(bold=True)
                    cell.border = thin_border
                    if col_name not in ['Họ và tên']:
                        cell.alignment = Alignment(horizontal='center')
                current_row += 2

                # Bảng phổ điểm
                sheet.cell(row=current_row, column=1, value='Điểm').font = Font(bold=True)
                sheet.cell(row=current_row, column=1).border = thin_border
                for c_idx, sub in enumerate(active_subjects, 2):
                    cell = sheet.cell(row=current_row, column=c_idx, value=sub.capitalize())
                    cell.font = Font(bold=True)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center')
                current_row += 1

                for r_name in ranges:
                    sheet.cell(row=current_row, column=1, value=r_name).border = thin_border
                    for c_idx, sub in enumerate(active_subjects, 2):
                        if sub in df_cls.columns:
                            count = sum(1 for score in df_cls[sub] if get_score_range(score) == r_name)
                            cell = sheet.cell(row=current_row, column=c_idx, value=count)
                        else:
                            cell = sheet.cell(row=current_row, column=c_idx, value=0)
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center')
                    current_row += 1

                # Auto-fit columns
                for col in sheet.columns:
                    column_letter = col[0].column_letter
                    header_val = sheet[f'{column_letter}3'].value
                    
                    if header_val == 'STT': width = 5
                    elif header_val == 'Họ và tên': width = 25
                    elif header_val == 'Ngày sinh': width = 12
                    elif header_val == 'SBD': width = 12
                    elif header_val == 'Lớp': width = 8
                    elif header_val in active_subjects: width = 6
                    elif header_val == 'Tổng điểm': width = 10
                    elif header_val == 'TBM': width = 6
                    elif str(header_val).startswith('Khối'): width = 8
                    elif str(header_val) == 'XT': width = 5
                    else: width = 8
                    
                    sheet.column_dimensions[column_letter].width = width

            build_class_sheet(ws_cls)
            build_class_sheet(ws_ind)
            
            # Save individual workbook and add to zip
            ind_filename = f"Bang_diem_lop_{cls}.xlsx"
            wb_ind.save(ind_filename)
            zipf.write(ind_filename)
            os.remove(ind_filename)

    wb.save(output_filename)
    print(f"Đã xử lý xong! File kết quả: {output_filename} và {zip_filename}")
    files.download(output_filename)
    files.download(zip_filename)

process_scores()`;

export default function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [kyThi, setKyThi] = useState('Giữa HKI');
  const [namHoc, setNamHoc] = useState('2026-2027');
  const [khoiLop, setKhoiLop] = useState('Khối 6');
  const [dashboardData, setDashboardData] = useState<DashboardData | null>(null);
  const [dashSubject, setDashSubject] = useState<string>('');
  const [dashClass, setDashClass] = useState<string>('');
  const [searchTerm, setSearchTerm] = useState<string>('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  const kyThiOptions: string[] = [
    'Giữa HKI',
    'Cuối HKI',
    'Giữa HKII',
    'Cuối HKII',
    'Thử TS 10 Lần 1',
    'Thử TS 10 Lần 2',
    'Thử TS 10 Lần 3'
  ];

  const namHocOptions: string[] = [
    '2025-2026',
    '2026-2027',
    '2027-2028'
  ];

  const khoiLopOptions: string[] = [
    'Khối 6',
    'Khối 7',
    'Khối 8',
    'Khối 9'
  ];

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const droppedFiles = Array.from(e.dataTransfer.files).filter(
      (f: File) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    setFiles(prev => [...prev, ...droppedFiles]);
  }, []);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const selectedFiles = Array.from(e.target.files).filter(
        (f: File) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
      );
      setFiles(prev => [...prev, ...selectedFiles]);
    }
  }, []);

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const processExcelFiles = async () => {
    if (files.length === 0) return;
    setIsProcessing(true);

    try {
      const studentData: Record<string, any> = {};
      const subjects: string[] = [];

      // 1. Read all files
      for (const file of files) {
        let subjectName = file.name.replace(/\.[^/.]+$/, "").trim();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const worksheet = workbook.worksheets[0];

        // Tìm tên môn học từ dòng 4 hoặc các dòng đầu (VD: "BẢNG ĐIỂM THI MÔN: VĂN 6")
        for (let r = 1; r <= 5; r++) {
          const row = worksheet.getRow(r);
          let foundSubject = false;
          row.eachCell((cell) => {
            const text = cell.value?.toString().trim() || '';
            const match = text.match(/bảng\s+điểm\s+thi\s+môn\s*[:\-]\s*(.+)/i) || text.match(/môn\s*[:\-]\s*(.+)/i);
            if (match && match[1]) {
              const detected = match[1].trim();
              if (detected && (subjectName.length > 25 || subjectName.toLowerCase().startsWith('bảng điểm') || subjectName.toLowerCase().startsWith('bang_diem') || subjectName.toLowerCase().includes('export') || !subjectName)) {
                subjectName = detected;
              }
              foundSubject = true;
            }
          });
          if (foundSubject) break;
        }

        if (!subjects.includes(subjectName)) {
          subjects.push(subjectName);
        }

        // Tự động nhận dạng vị trí cột và dòng tiêu đề
        // Cấu trúc chuẩn THCS Mỹ Xá:
        // Dòng 5: Cột A (1): STT | Cột B (2): Họ và tên | Cột C (3): Ngày sinh | Cột D (4): Giới | Cột E (5): Phòng thi | Cột F (6): Địa điểm thi | Cột G (7): Số BD | Cột H (8): Mã đề | Cột I (9): Điểm | Cột J (10): Ghi chú (Lớp)
        let headerRowNumber = 5;
        let colHoTen = 2; // Cột B
        let colNgaySinh = 3; // Cột C
        let colSbd = 7; // Cột G
        let colDiem = 9; // Cột I
        let colLop = 10; // Cột J (Ghi chú)

        for (let r = 1; r <= 10; r++) {
          const row = worksheet.getRow(r);
          let foundHeader = false;
          row.eachCell((cell, colNumber) => {
            const raw = cell.value?.toString().trim().toLowerCase() || '';
            if (raw === 'họ và tên' || raw === 'họ tên' || raw.includes('họ và tên')) {
              colHoTen = colNumber;
              foundHeader = true;
            } else if (raw === 'số bd' || raw === 'sbd' || raw === 'số báo danh' || raw.includes('số bd')) {
              colSbd = colNumber;
              foundHeader = true;
            } else if (raw === 'ngày sinh' || raw.includes('ngày sinh')) {
              colNgaySinh = colNumber;
            } else if (raw === 'điểm' || (raw.includes('điểm') && !raw.includes('địa điểm'))) {
              colDiem = colNumber;
            } else if (raw === 'ghi chú' || raw === 'lớp' || raw.includes('ghi chú')) {
              colLop = colNumber;
            }
          });
          if (foundHeader) {
            headerRowNumber = r;
            break;
          }
        }

        const startDataRow = headerRowNumber + 1; // Bắt đầu từ dòng 6 (Học sinh đầu tiên STT 1)

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber < startDataRow) return;

          const sbdCell = row.getCell(colSbd);
          let sbd = sbdCell.value !== null && sbdCell.value !== undefined ? sbdCell.value.toString().trim() : '';
          if (sbd.endsWith('.0') && !isNaN(Number(sbd))) {
            sbd = sbd.replace(/\.0$/, '');
          }

          const hoTenCell = row.getCell(colHoTen);
          const hoTen = hoTenCell.value?.toString().trim();
          
          // Bỏ qua dòng rác, dòng trống hoặc dòng ký tên cuối bảng
          if (!sbd || sbd.includes('...') || !hoTen) return;
          if (/^(stt|họ và tên|tổng số|tổng cộng|giáo viên|người lập|hiệu trưởng)/i.test(hoTen)) return;

          if (!studentData[sbd]) {
            let ngaySinhRaw = row.getCell(colNgaySinh).value;
            let ngaySinh = '';
            if (ngaySinhRaw instanceof Date) {
              const d = ngaySinhRaw.getDate().toString().padStart(2, '0');
              const m = (ngaySinhRaw.getMonth() + 1).toString().padStart(2, '0');
              const y = ngaySinhRaw.getFullYear();
              ngaySinh = `${d}/${m}/${y}`;
            } else if (typeof ngaySinhRaw === 'object' && ngaySinhRaw !== null) {
              const txt = (ngaySinhRaw as any).text || (ngaySinhRaw as any).result || '';
              ngaySinh = txt.toString().trim();
            } else if (ngaySinhRaw !== null && ngaySinhRaw !== undefined) {
              ngaySinh = ngaySinhRaw.toString().trim();
            }

            let lopRaw = row.getCell(colLop).value;
            let lop = '';
            if (typeof lopRaw === 'object' && lopRaw !== null) {
              lop = ((lopRaw as any).text || (lopRaw as any).result || '').toString().trim();
            } else if (lopRaw !== null && lopRaw !== undefined) {
              lop = lopRaw.toString().trim();
            }

            studentData[sbd] = {
              hoTen: hoTen,
              ngaySinh: ngaySinh,
              sbd: sbd,
              lop: lop,
              scores: {}
            };
          }

          const scoreRaw = row.getCell(colDiem).value;
          let parsedScore: number | null = null;
          if (typeof scoreRaw === 'number') {
            if (!isNaN(scoreRaw)) parsedScore = scoreRaw;
          } else if (typeof scoreRaw === 'string') {
            const str = scoreRaw.trim().replace(',', '.');
            if (str !== '' && str !== '-' && str.toLowerCase() !== 'v') {
              const num = parseFloat(str);
              if (!isNaN(num)) parsedScore = num;
            }
          } else if (typeof scoreRaw === 'object' && scoreRaw !== null) {
            const res = (scoreRaw as any).result ?? (scoreRaw as any).text;
            if (typeof res === 'number') {
              if (!isNaN(res)) parsedScore = res;
            } else if (typeof res === 'string') {
              const str = res.trim().replace(',', '.');
              if (str !== '' && str !== '-' && str.toLowerCase() !== 'v') {
                const num = parseFloat(str);
                if (!isNaN(num)) parsedScore = num;
              }
            }
          }

          if (parsedScore !== null) {
            studentData[sbd].scores[subjectName] = parsedScore;
          }
        });
      }

      // 2. Process Data
      let dataList = Object.values(studentData);

      if (dataList.length === 0) {
        alert("Không tìm thấy danh sách học sinh hợp lệ trong file tải lên. Vui lòng kiểm tra lại file Excel.");
        setIsProcessing(false);
        return;
      }

      // Nếu có học sinh đã nhập điểm, lọc những học sinh có ít nhất 1 điểm; nếu file mẫu chưa nhập điểm, giữ toàn bộ danh sách
      const studentsWithScores = dataList.filter((student: any) => Object.keys(student.scores).length > 0);
      if (studentsWithScores.length > 0) {
        dataList = studentsWithScores;
      }

      // Sắp xếp môn học theo chuẩn THCS
      const subjectOrder = [
        'toán', 'toan',
        'văn', 'van', 'ngữ văn', 'ngu van',
        'anh', 'tiếng anh', 'tieng anh', 'ngoại ngữ',
        'khtn', 'khoa học tự nhiên',
        'lí', 'li', 'lý', 'ly', 'vật lí', 'vật lý',
        'hoá', 'hoa', 'hóa', 'hoá học', 'hóa học',
        'sinh', 'sinh học',
        'ls&đl', 'lịch sử và địa lí', 'sử', 'su', 'lịch sử',
        'địa', 'dia', 'địa lí', 'địa lý',
        'gdcd', 'tin', 'tin học', 'công nghệ'
      ];
      const getSubjectOrderIndex = (sub: string) => {
        const s = sub.toLowerCase().trim();
        for (let i = 0; i < subjectOrder.length; i++) {
          if (s === subjectOrder[i] || s.startsWith(subjectOrder[i])) {
            return i;
          }
        }
        return 999;
      };

      subjects.sort((a, b) => {
        const idxA = getSubjectOrderIndex(a);
        const idxB = getSubjectOrderIndex(b);
        if (idxA !== idxB) return idxA - idxB;
        return a.localeCompare(b);
      });

      // Calculate Total and Average
      dataList.forEach(student => {
        let total = 0;
        let count = 0;
        subjects.forEach(sub => {
          if (student.scores[sub] !== undefined && student.scores[sub] !== null) {
            total += student.scores[sub];
            count++;
          }
        });
        student.tongDiem = count > 0 ? Number(total.toFixed(2)) : null;
        student.tbm = count > 0 ? Number((total / count).toFixed(2)) : null;
      });

      // Rank by Total Score
      const sortedByTotal = [...dataList].filter(s => s.tongDiem !== null).sort((a, b) => b.tongDiem - a.tongDiem);
      let currentRank = 1;
      let currentScore = -1;
      
      sortedByTotal.forEach((student, index) => {
        if (student.tongDiem !== currentScore) {
          currentRank = index + 1;
          currentScore = student.tongDiem;
        }
        student.xt = currentRank;
      });

      // Calculate Blocks
      // Bảng điểm THCS không tính điểm các khối
      const blocks: Record<string, string[]> = {};

      // Sort by SBD
      dataList.sort((a, b) => {
        return a.sbd.localeCompare(b.sbd);
      });

      // --- Prepare Dashboard Data ---
      const scoreRanges = [
        { label: '< 1,0', min: 0, max: 1 },
        { label: '1,0 - < 2,0', min: 1, max: 2 },
        { label: '2,0 - < 3,0', min: 2, max: 3 },
        { label: '3,0 - < 4,0', min: 3, max: 4 },
        { label: '4,0 - < 5,0', min: 4, max: 5 },
        { label: '5,0 - < 6,0', min: 5, max: 6 },
        { label: '6,0 - < 7,0', min: 6, max: 7 },
        { label: '7,0 - < 8,0', min: 7, max: 8 },
        { label: '8,0 - < 9,0', min: 8, max: 9 },
        { label: '9,0 - < 10,0', min: 9, max: 10 },
        { label: '= 10,0', min: 10, max: 10.1 }
      ];

      const getScoreRangeLabel = (score: number) => {
        if (score < 1) return '< 1,0';
        if (score < 2) return '1,0 - < 2,0';
        if (score < 3) return '2,0 - < 3,0';
        if (score < 4) return '3,0 - < 4,0';
        if (score < 5) return '4,0 - < 5,0';
        if (score < 6) return '5,0 - < 6,0';
        if (score < 7) return '6,0 - < 7,0';
        if (score < 8) return '7,0 - < 8,0';
        if (score < 9) return '8,0 - < 9,0';
        if (score < 10) return '9,0 - < 10,0';
        return '= 10,0';
      };

      const overallDist: Record<string, ScoreRangeCount[]> = {};
      const overallAvg: Record<string, number> = {};

      const validDashboardDataList = dataList.filter(s => {
        if (s.lop === 'TSTD') return false;
        if (s.hoTen === 'Đặng Mai Minh Hoàng') return false;
        return true;
      });

      subjects.forEach(sub => {
        const dist = scoreRanges.map(r => ({ name: r.label, count: 0 }));
        let sum = 0;
        let count = 0;
        validDashboardDataList.forEach(s => {
          const score = s.scores[sub];
          if (typeof score === 'number') {
            const lbl = getScoreRangeLabel(score);
            const entry = dist.find(d => d.name === lbl);
            if (entry) entry.count++;
            sum += score;
            count++;
          }
        });
        overallDist[sub] = dist;
        overallAvg[sub] = count > 0 ? Number((sum / count).toFixed(2)) : 0;
      });

      const uniqueClasses = Array.from(new Set(dataList.map(s => s.lop).filter(l => l && l !== 'TSTD')));
      uniqueClasses.sort();

      const classDist: Record<string, Record<string, ScoreRangeCount[]>> = {};
      const classDistGrouped: Record<string, any[]> = {};
      
      uniqueClasses.forEach(cls => {
        classDist[cls] = {};
        const clsStudents = validDashboardDataList.filter(s => s.lop === cls);
        
        const grouped = scoreRanges.map(r => {
          const row: any = { name: r.label };
          subjects.forEach(sub => row[sub] = 0);
          return row;
        });

        subjects.forEach(sub => {
          const dist = scoreRanges.map(r => ({ name: r.label, count: 0 }));
          clsStudents.forEach(s => {
             const score = s.scores[sub];
             if (typeof score === 'number') {
               const lbl = getScoreRangeLabel(score);
               const entry = dist.find(d => d.name === lbl);
               if (entry) entry.count++;
               
               const groupEntry = grouped.find(d => d.name === lbl);
               if (groupEntry) groupEntry[sub]++;
             }
          });
          classDist[cls][sub] = dist;
        });
        
        classDistGrouped[cls] = grouped;
      });

      setDashboardData({
        subjects,
        classes: uniqueClasses,
        overall: overallDist,
        byClass: classDist,
        byClassGrouped: classDistGrouped,
        overallAvg: overallAvg,
        students: dataList,
        blocks: blocks
      });
      if (subjects.length > 0) setDashSubject(subjects[0]);
      if (uniqueClasses.length > 0) setDashClass(uniqueClasses[0]);
      // --------------------------------

      // 3. Generate Output Excel
      const outWorkbook = new ExcelJS.Workbook();
      const outSheet = outWorkbook.addWorksheet('Bảng Điểm Tổng Hợp');
      const topLimit = ['Khối 6', 'Khối 7', 'Khối 8', 'Khối 10', 'Khối 11'].includes(khoiLop) ? 15 : 5;
      const topSheetName = `Top ${topLimit}`;
      const top5Sheet = outWorkbook.addWorksheet(topSheetName);
      const tbdiemSheet = outWorkbook.addWorksheet('TBDIEM');
      const tksoSheet = outWorkbook.addWorksheet('TKSo');
      const tktruotSheet = outWorkbook.addWorksheet('TKTruot');
      const dstruotSheet = outWorkbook.addWorksheet('DSTruot');
      
      const classSheets: Record<string, ExcelJS.Worksheet> = {};
      const exportClasses = Array.from(new Set(dataList.map(s => s.lop))).filter(Boolean).sort();
      exportClasses.forEach(cls => {
        classSheets[cls as string] = outWorkbook.addWorksheet(cls as string);
      });

      // Define columns
      const columns = [
        { header: 'STT', key: 'stt', width: 5 },
        { header: 'Họ và tên', key: 'hoTen', width: 25 },
        { header: 'Ngày sinh', key: 'ngaySinh', width: 12 },
        { header: 'SBD', key: 'sbd', width: 12 },
        { header: 'Lớp', key: 'lop', width: 8 },
        ...subjects.map(sub => ({ header: sub, key: `sub_${sub}`, width: 6 })),
        { header: 'Tổng điểm', key: 'tongDiem', width: 10 },
        { header: 'XT', key: 'xt', width: 5 },
        { header: 'TBM', key: 'tbm', width: 6 }
      ];

      outSheet.columns = columns;

      // Title Row (Row 1)
      outSheet.insertRow(1, []);
      outSheet.insertRow(1, []);
      outSheet.mergeCells(1, 1, 1, columns.length);
      const titleCell = outSheet.getCell('A1');
      titleCell.value = `BẢNG ĐIỂM ${kyThi.toUpperCase()} ${khoiLop.toUpperCase()} - NĂM HỌC ${namHoc}`;
      titleCell.font = { bold: true, size: 14 };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      
      // Apply borders to row 1 and row 2
      for (let r = 1; r <= 2; r++) {
        const row = outSheet.getRow(r);
        for (let c = 1; c <= columns.length; c++) {
          row.getCell(c).border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
        }
      }

      // Header Row (Row 3)
      const headerRow = outSheet.getRow(3);
      columns.forEach((col, idx) => {
        const cell = headerRow.getCell(idx + 1);
        cell.value = col.header;
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD9D9D9' }
        };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      // Data Rows
      dataList.forEach((student, idx) => {
        const rowData: any = {
          stt: idx + 1,
          hoTen: student.hoTen,
          ngaySinh: student.ngaySinh,
          sbd: student.sbd,
          lop: student.lop,
          tongDiem: student.tongDiem,
          xt: student.xt,
          tbm: student.tbm
        };

        subjects.forEach(sub => {
          rowData[`sub_${sub}`] = student.scores[sub];
        });

        const row = outSheet.addRow(rowData);
        
        // Styling data cells
        for (let c = 1; c <= columns.length; c++) {
          const cell = row.getCell(c);
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          if (c !== 2) { // Center everything except name
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        }
      });

      // Add Average Row
      const avgRowData: any = {
        hoTen: 'ĐIỂM TRUNG BÌNH'
      };
      
      const colsToAverage = [...subjects, 'tongDiem', 'tbm'];
      
      const validListForOverallAvg = dataList.filter(s => {
        if (s.lop === 'TSTD') return false;
        if (s.hoTen === 'Đặng Mai Minh Hoàng') return false;
        return true;
      });

      colsToAverage.forEach(colKey => {
        let sum = 0;
        let count = 0;
        validListForOverallAvg.forEach((student: any) => {
          let val = null;
          if (subjects.includes(colKey)) {
            val = student.scores[colKey];
          } else {
            val = student[colKey];
          }
          if (val !== undefined && val !== null) {
            sum += val;
            count++;
          }
        });
        if (count > 0) {
          if (subjects.includes(colKey)) {
            avgRowData[`sub_${colKey}`] = Number((sum / count).toFixed(2));
          } else {
            avgRowData[colKey] = Number((sum / count).toFixed(2));
          }
        }
      });

      const avgRow = outSheet.addRow(avgRowData);
      avgRow.font = { bold: true };
      
      for (let c = 1; c <= columns.length; c++) {
        const cell = avgRow.getCell(c);
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        if (c !== 2) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      }

      // --- Top 5/15 Sheet ---
      
      top5Sheet.mergeCells('A1:L1');
      const top5Title = top5Sheet.getCell('A1');
      top5Title.value = `DANH SÁCH HỌC SINH - TOP ${topLimit} - ${kyThi.toUpperCase()} ${khoiLop.toUpperCase()} - NĂM HỌC ${namHoc}`;
      top5Title.font = { bold: true, size: 12 };
      top5Title.alignment = { vertical: 'middle' };
      top5Title.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
      };

      let currentRow = 3;

      const renderTop5Table = (title: string, scoreKey: string, subjectKeys: string[], limit: number) => {
        const sortedStudents = [...dataList]
          .filter(s => s[scoreKey] !== undefined && s[scoreKey] !== null)
          .sort((a, b) => b[scoreKey] - a[scoreKey])
          .slice(0, limit);

        if (sortedStudents.length === 0) return;

        const tableTitleCell = top5Sheet.getCell(`A${currentRow}`);
        tableTitleCell.value = title;
        tableTitleCell.font = { bold: true };
        currentRow++;

        const headers = ['STT', 'Họ và tên', 'Số BD', 'Lớp', ...subjectKeys, 'Điểm'];
        const headerRow = top5Sheet.getRow(currentRow);
        headers.forEach((headerText, idx) => {
          const cell = headerRow.getCell(idx + 1);
          cell.value = headerText;
          cell.font = { bold: true };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });
        currentRow++;

        sortedStudents.forEach((student, idx) => {
          const row = top5Sheet.getRow(currentRow);
          const rowData = [
            idx + 1,
            student.hoTen,
            student.sbd,
            student.lop,
            ...subjectKeys.map(sub => student.scores[sub]),
            student[scoreKey]
          ];
          
          rowData.forEach((val, colIdx) => {
            const cell = row.getCell(colIdx + 1);
            cell.value = val;
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            if (colIdx !== 1) {
              cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
          });
          currentRow++;
        });

        currentRow++;
      };

      renderTop5Table(`TOP ${topLimit} HỌC SINH CAO ĐIỂM NHẤT - TỔNG ĐIỂM`, 'tongDiem', subjects, topLimit);

      top5Sheet.columns.forEach((column, index) => {
        if (index === 0) column.width = 5;
        else if (index === 1) column.width = 25;
        else if (index === 2) column.width = 12;
        else if (index === 3) column.width = 8;
        else if (index < 4 + subjects.length) column.width = 6;
        else if (index === 4 + subjects.length) column.width = 8;
        else column.width = 8;
      });

      // --- TBDIEM Sheet ---
      const classes = Array.from(new Set(dataList.map(s => s.lop))).filter(Boolean).sort();
      
      const colsToAvg = ['tbm', ...subjects, 'tongDiem'];
      const classAvgs: any[] = [];
      
      const validStudentsForAvg = dataList.filter(s => {
        if (s.lop === 'TSTD') return false;
        if (s.hoTen === 'Đặng Mai Minh Hoàng') return false;
        return true;
      });
      
      classes.forEach(cls => {
        if (cls === 'TSTD') return;
        const classData = validStudentsForAvg.filter(s => s.lop === cls);
        const avgData: any = { lop: cls };
        
        colsToAvg.forEach(colKey => {
          let sum = 0;
          let count = 0;
          classData.forEach((student: any) => {
            let val = null;
            if (subjects.includes(colKey)) {
              val = student.scores[colKey];
            } else {
              val = student[colKey];
            }
            if (val !== undefined && val !== null) {
              sum += val;
              count++;
            }
          });
          if (count > 0) {
            avgData[colKey] = Number((sum / count).toFixed(2));
          } else {
            avgData[colKey] = null;
          }
        });
        classAvgs.push(avgData);
      });
      
      // Calculate ranks
      colsToAvg.forEach(colKey => {
        const sorted = [...classAvgs].filter(c => c[colKey] !== null).sort((a, b) => b[colKey] - a[colKey]);
        let currentRank = 1;
        let currentScore = -1;
        sorted.forEach((clsAvg, index) => {
          if (clsAvg[colKey] !== currentScore) {
            currentRank = index + 1;
            currentScore = clsAvg[colKey];
          }
          const originalCls = classAvgs.find(c => c.lop === clsAvg.lop);
          if (originalCls) {
            originalCls[`xt_${colKey}`] = currentRank;
          }
        });
      });
      
      // Title Row
      tbdiemSheet.insertRow(1, []);
      tbdiemSheet.insertRow(1, []);
      tbdiemSheet.mergeCells(1, 1, 1, 1 + colsToAvg.length * 2);
      const tbdiemTitle = tbdiemSheet.getCell('A1');
      tbdiemTitle.value = `BẢNG TỔNG HỢP ĐIỂM TRUNG BÌNH CÁC LỚP - ${kyThi.toUpperCase()} ${khoiLop.toUpperCase()} - NĂM HỌC ${namHoc}`;
      tbdiemTitle.font = { bold: true, size: 14 };
      tbdiemTitle.alignment = { vertical: 'middle' };
      
      // Header Row
      let tbdiemCurrentRow = 3;
      const tbdiemHeaderRow = tbdiemSheet.getRow(tbdiemCurrentRow);
      tbdiemHeaderRow.getCell(1).value = 'Lớp';
      tbdiemHeaderRow.getCell(1).font = { bold: true };
      tbdiemHeaderRow.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      tbdiemHeaderRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      
      let cIdx = 2;
      colsToAvg.forEach(colKey => {
        let headerText = colKey;
        if (colKey === 'tbm') headerText = 'TBM';
        else if (colKey === 'tongDiem') headerText = 'Tổng điểm';
        else if (subjects.includes(colKey)) headerText = colKey.charAt(0).toUpperCase() + colKey.slice(1);
        
        const cellVal = tbdiemHeaderRow.getCell(cIdx);
        cellVal.value = headerText;
        cellVal.font = { bold: true };
        cellVal.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellVal.alignment = { horizontal: 'center', vertical: 'middle' };
        
        const cellXt = tbdiemHeaderRow.getCell(cIdx + 1);
        cellXt.value = 'XT';
        cellXt.font = { bold: true };
        cellXt.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellXt.alignment = { horizontal: 'center', vertical: 'middle' };
        
        cIdx += 2;
      });
      tbdiemCurrentRow++;
      
      // Data Rows
      classAvgs.forEach(clsAvg => {
        const row = tbdiemSheet.getRow(tbdiemCurrentRow);
        row.getCell(1).value = clsAvg.lop;
        row.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        let colIdx = 2;
        colsToAvg.forEach(colKey => {
          const cellVal = row.getCell(colIdx);
          cellVal.value = clsAvg[colKey] !== null ? clsAvg[colKey] : '';
          cellVal.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          cellVal.alignment = { horizontal: 'center' };
          
          const cellXt = row.getCell(colIdx + 1);
          cellXt.value = clsAvg[`xt_${colKey}`] !== undefined ? clsAvg[`xt_${colKey}`] : '';
          cellXt.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          cellXt.alignment = { horizontal: 'center' };
          
          colIdx += 2;
        });
        tbdiemCurrentRow++;
      });
      
      // TB Khối Row
      const tbKhoiRow = tbdiemSheet.getRow(tbdiemCurrentRow);
      tbKhoiRow.getCell(1).value = 'TB Khối';
      tbKhoiRow.getCell(1).font = { bold: true };
      tbKhoiRow.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      
      let colIdxTB = 2;
      colsToAvg.forEach(colKey => {
        let sum = 0;
        let count = 0;
        validStudentsForAvg.forEach((student: any) => {
          let val = null;
          if (subjects.includes(colKey)) {
            val = student.scores[colKey];
          } else {
            val = student[colKey];
          }
          if (val !== undefined && val !== null) {
            sum += val;
            count++;
          }
        });
        
        const cellVal = tbKhoiRow.getCell(colIdxTB);
        cellVal.value = count > 0 ? Number((sum / count).toFixed(2)) : '';
        cellVal.font = { bold: true };
        cellVal.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellVal.alignment = { horizontal: 'center' };
        
        const cellXt = tbKhoiRow.getCell(colIdxTB + 1);
        cellXt.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        colIdxTB += 2;
      });
      
      // Auto-fit columns
      tbdiemSheet.columns.forEach((column, index) => {
        if (index === 0) column.width = 8;
        else if (index % 2 === 0) column.width = 5;
        else column.width = 8;
      });

      // --- Sheets for each class ---
      
      const getScoreRange = (score: number | null | undefined) => {
        if (score === null || score === undefined) return null;
        if (score < 1.0) return '<1,0';
        if (score < 2.0) return '1,0-<2,0';
        if (score < 3.0) return '2,0-<3,0';
        if (score < 4.0) return '3,0-<4,0';
        if (score < 5.0) return '4,0-<5,0';
        if (score < 6.0) return '5,0-<6,0';
        if (score < 7.0) return '6,0-<7,0';
        if (score < 8.0) return '7,0-<8,0';
        if (score < 9.0) return '8,0-<9,0';
        if (score < 10.0) return '9,0-<10,0';
        return '= 10,0';
      };
      
      const ranges = ['<1,0', '1,0-<2,0', '2,0-<3,0', '3,0-<4,0', '4,0-<5,0', '5,0-<6,0', '6,0-<7,0', '7,0-<8,0', '8,0-<9,0', '9,0-<10,0', '= 10,0'];

      const zip = new JSZip();

      for (const cls of classes) {
        const classSheet = classSheets[cls as string];
        const indWorkbook = new ExcelJS.Workbook();
        const indClassSheet = indWorkbook.addWorksheet(cls as string);
        
        const classData = dataList.filter(s => s.lop === cls).map(s => ({ ...s }));
        
        // Determine active subjects for this class
        const activeSubjects = subjects.filter(sub => {
          return classData.some(student => student.scores[sub] !== undefined && student.scores[sub] !== null);
        });

        // Filter columns
        const classColumns = columns.filter(col => {
          if (col.key?.startsWith('sub_')) {
            const sub = col.key.replace('sub_', '');
            return activeSubjects.includes(sub);
          }
          return true;
        });
        
        // Xếp thứ lại theo từng lớp
        const sortedByTotalCls = [...classData].filter(s => s.tongDiem !== null).sort((a, b) => b.tongDiem - a.tongDiem);
        let currentRank = 1;
        let currentScore = -1;
        sortedByTotalCls.forEach((student, index) => {
          if (student.tongDiem !== currentScore) {
            currentRank = index + 1;
            currentScore = student.tongDiem;
          }
          student.xt = currentRank;
        });
        
        const buildSheet = (sheet: ExcelJS.Worksheet) => {
          // Title Row
          sheet.insertRow(1, []);
          sheet.insertRow(1, []);
          sheet.mergeCells(1, 1, 1, classColumns.length);
          const titleCell = sheet.getCell('A1');
          titleCell.value = `BẢNG ĐIỂM ${kyThi.toUpperCase()} LỚP ${cls} - NĂM HỌC ${namHoc}`;
          titleCell.font = { bold: true, size: 14 };
          titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

          // Header Row
          const headerRow = sheet.getRow(3);
          classColumns.forEach((col, idx) => {
            const cell = headerRow.getCell(idx + 1);
            cell.value = col.header;
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          });

          // Data Rows
          let currentRow = 4;
          classData.forEach((student, idx) => {
            const row = sheet.getRow(currentRow);
            classColumns.forEach((col, cIdx) => {
              const cell = row.getCell(cIdx + 1);
              let val: any = '';
              if (col.key === 'stt') val = idx + 1;
              else if (col.key?.startsWith('sub_')) {
                const sub = col.key.replace('sub_', '');
                val = student.scores[sub] !== undefined ? student.scores[sub] : '';
              } else if (col.key === 'xt') {
                val = student.xt;
              } else if (col.key === 'tbm') {
                val = student.tbm;
              } else if (col.key === 'tongDiem') {
                val = student.tongDiem;
              } else if (col.key && student[col.key as keyof typeof student] !== undefined) {
                val = student[col.key as keyof typeof student];
              }
              
              cell.value = val !== undefined ? val : '';
              cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              if (cIdx !== 1) {
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
              }
            });
            currentRow++;
          });

          // Average Row
          const avgRow = sheet.getRow(currentRow);
          classColumns.forEach((col, cIdx) => {
            const cell = avgRow.getCell(cIdx + 1);
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            if (cIdx !== 1) {
              cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
          });
          
          avgRow.getCell(2).value = 'ĐIỂM TRUNG BÌNH';
          avgRow.font = { bold: true };
          
          const clsColsToAverage = [...activeSubjects, 'tongDiem', 'tbm'];
          clsColsToAverage.forEach(colKey => {
            let sum = 0;
            let count = 0;
            const validClassDataForAvg = classData.filter(s => {
              if (s.hoTen === 'Đặng Mai Minh Hoàng') return false;
              return true;
            });
            validClassDataForAvg.forEach((student: any) => {
              let val = null;
              if (activeSubjects.includes(colKey)) {
                val = student.scores[colKey];
              } else {
                val = student[colKey];
              }
              if (val !== undefined && val !== null) {
                sum += val;
                count++;
              }
            });
            if (count > 0) {
              const searchKey = activeSubjects.includes(colKey) ? `sub_${colKey}` : colKey;
              const colIdx = classColumns.findIndex(c => c.key === searchKey);
              if (colIdx !== -1) {
                avgRow.getCell(colIdx + 1).value = Number((sum / count).toFixed(2));
              }
            }
          });
          
          currentRow += 2;

          // Score Distribution Table
          const distTitleRow = sheet.getRow(currentRow);
          distTitleRow.getCell(1).value = 'Điểm';
          distTitleRow.getCell(1).font = { bold: true };
          distTitleRow.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          
          activeSubjects.forEach((sub, idx) => {
            const cell = distTitleRow.getCell(idx + 2);
            cell.value = sub.charAt(0).toUpperCase() + sub.slice(1);
            cell.font = { bold: true };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { horizontal: 'center' };
          });
          currentRow++;

          ranges.forEach(rName => {
            const row = sheet.getRow(currentRow);
            const cell1 = row.getCell(1);
            cell1.value = rName;
            cell1.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            
            activeSubjects.forEach((sub, idx) => {
              let count = 0;
              classData.forEach((student: any) => {
                if (getScoreRange(student.scores[sub]) === rName) {
                  count++;
                }
              });
              const cell = row.getCell(idx + 2);
              cell.value = count;
              cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              cell.alignment = { horizontal: 'center' };
            });
            currentRow++;
          });

          // Auto-fit columns for class sheet
          sheet.columns.forEach((column, index) => {
            if (classColumns[index]) {
              column.width = classColumns[index].width;
            } else {
              column.width = 10;
            }
          });
        };

        buildSheet(classSheet);
        buildSheet(indClassSheet);

        const indBuffer = await indWorkbook.xlsx.writeBuffer();
        zip.file(`Bang_diem_lop_${cls}.xlsx`, indBuffer);
      }

      // --- TKSo Sheet ---
      
      const tkRanges = [
        { label: 'Điểm 0', check: (s: number) => s === 0 },
        { label: '>= 0 ÷ < 1', check: (s: number) => s > 0 && s < 1 },
        { label: '>= 1 ÷ < 2', check: (s: number) => s >= 1 && s < 2 },
        { label: '>= 2 ÷ < 3', check: (s: number) => s >= 2 && s < 3 },
        { label: '>= 3 ÷ < 4', check: (s: number) => s >= 3 && s < 4 },
        { label: '>= 4 ÷ < 5', check: (s: number) => s >= 4 && s < 5 },
        { label: '>= 5 ÷ < 6', check: (s: number) => s >= 5 && s < 6 },
        { label: '>= 6 ÷ < 7', check: (s: number) => s >= 6 && s < 7 },
        { label: '>= 7 ÷ < 8', check: (s: number) => s >= 7 && s < 8 },
        { label: '>= 8 ÷ < 9', check: (s: number) => s >= 8 && s < 9 },
        { label: '>= 9 ÷ <= 10', check: (s: number) => s >= 9 && s <= 10 },
        { label: 'Điểm 10', check: (s: number) => s === 10 }
      ];
      
      // Header dòng 1
      tksoSheet.insertRow(1, []);
      tksoSheet.insertRow(1, []);
      const tksoTitle = tksoSheet.getCell('C1');
      tksoTitle.value = 'PHỔ ĐIỂM';
      tksoTitle.font = { bold: true };
      
      // Header dòng 2 & 3
      const tksoRow2 = tksoSheet.getRow(2);
      const tksoRow3 = tksoSheet.getRow(3);
      
      tksoRow3.getCell(1).value = 'Môn';
      tksoRow3.getCell(1).font = { bold: true };
      tksoRow3.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      tksoRow3.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      
      tksoRow3.getCell(2).value = 'Tổng số HS KS';
      tksoRow3.getCell(2).font = { bold: true };
      tksoRow3.getCell(2).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      tksoRow3.getCell(2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      
      let tksoCIdx = 3;
      tkRanges.forEach(r => {
        tksoSheet.mergeCells(2, tksoCIdx, 2, tksoCIdx + 1);
        const cellLabel = tksoRow2.getCell(tksoCIdx);
        cellLabel.value = r.label;
        cellLabel.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellLabel.alignment = { horizontal: 'center', vertical: 'middle' };
        tksoRow2.getCell(tksoCIdx + 1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        const cellSl = tksoRow3.getCell(tksoCIdx);
        cellSl.value = 'SL';
        cellSl.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellSl.alignment = { horizontal: 'center', vertical: 'middle' };
        
        const cellTl = tksoRow3.getCell(tksoCIdx + 1);
        cellTl.value = 'TL %';
        cellTl.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellTl.alignment = { horizontal: 'center', vertical: 'middle' };
        
        tksoCIdx += 2;
      });
      
      // Data rows
      let tksoCurrentRow = 4;
      const validStudents = dataList.filter(s => s.hoTen !== 'ĐIỂM TRUNG BÌNH' && s.lop !== 'TSTD');
      
      subjects.forEach(sub => {
        const scores = validStudents.map(s => s.scores[sub]).filter(s => s !== undefined && s !== null);
        if (scores.length === 0) return;
        
        const totalHs = scores.length;
        const row = tksoSheet.getRow(tksoCurrentRow);
        
        row.getCell(1).value = sub.charAt(0).toUpperCase() + sub.slice(1);
        row.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        row.getCell(2).value = totalHs;
        row.getCell(2).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        let cIdx = 3;
        tkRanges.forEach(r => {
          let count = 0;
          scores.forEach(s => {
            if (r.check(s)) count++;
          });
          const tl = totalHs > 0 ? Number(((count / totalHs) * 100).toFixed(2)) : 0;
          
          const cellSl = row.getCell(cIdx);
          cellSl.value = count;
          cellSl.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          cellSl.alignment = { horizontal: 'center' };
          
          const cellTl = row.getCell(cIdx + 1);
          cellTl.value = tl;
          cellTl.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          cellTl.alignment = { horizontal: 'center' };
          
          cIdx += 2;
        });
        
        tksoCurrentRow++;
      });
      
      // Auto-fit columns
      tksoSheet.columns.forEach((column, index) => {
        if (index === 0) column.width = 12;
        else if (index === 1) column.width = 15;
        else column.width = 8;
      });

      // --- DSTruot Sheet ---
      
      const dstruotCols = ['stt', 'hoTen', 'ngaySinh', 'sbd', 'lop', ...subjects, 'tongDiem', 'tbm'];
      const dstruotHeaders = ['STT', 'Họ và tên', 'Ngày sinh', 'Số BD', 'Lớp', ...subjects.map(s => s.charAt(0).toUpperCase() + s.slice(1)), 'Tổng điểm', 'TBM'];
      
      const dstruotStudents = dataList.filter(s => {
        if (s.hoTen === 'ĐIỂM TRUNG BÌNH') return false;
        return subjects.some(sub => s.scores[sub] !== undefined && s.scores[sub] !== null && s.scores[sub] < 5);
      });
      
      dstruotSheet.insertRow(1, []);
      dstruotSheet.insertRow(1, []);
      dstruotSheet.mergeCells(1, 1, 1, dstruotCols.length);
      const dstruotTitle = dstruotSheet.getCell('A1');
      dstruotTitle.value = `DANH SÁCH HỌC SINH CÓ ĐIỂM DƯỚI 5 - ${kyThi.toUpperCase()} ${khoiLop.toUpperCase()} - NĂM HỌC ${namHoc}`;
      dstruotTitle.font = { bold: true, size: 12 };
      
      const dstruotHeaderRow = dstruotSheet.getRow(3);
      dstruotHeaders.forEach((header, idx) => {
        const cell = dstruotHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      
      dstruotStudents.forEach((student, index) => {
        const row = dstruotSheet.getRow(index + 4);
        dstruotCols.forEach((col, idx) => {
          const cell = row.getCell(idx + 1);
          let val = null;
          if (col === 'stt') val = index + 1;
          else if (subjects.includes(col)) val = student.scores[col];
          else val = (student as any)[col];
          
          cell.value = val !== undefined && val !== null ? val : '';
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          if (col !== 'hoTen') {
            cell.alignment = { horizontal: 'center' };
          }
        });
      });
      
      dstruotSheet.columns.forEach((column, index) => {
        const header = dstruotHeaders[index];
        if (header === 'STT') column.width = 5;
        else if (header === 'Họ và tên') column.width = 25;
        else if (header === 'Ngày sinh') column.width = 12;
        else if (header === 'Số BD') column.width = 12;
        else if (header === 'Lớp') column.width = 8;
        else if (subjects.includes(header.toLowerCase())) column.width = 6;
        else if (header === 'Tổng điểm') column.width = 10;
        else if (header === 'TBM') column.width = 6;
        else column.width = 8;
      });

      // --- TKTruot Sheet ---
      
      const tktruotCols = ['lop', ...subjects, 'tongHsTruot'];
      const tktruotHeaders = ['Lớp', ...subjects.map(s => s.charAt(0).toUpperCase() + s.slice(1)), 'Tổng số HS trượt'];
      
      tktruotSheet.insertRow(1, []);
      tktruotSheet.insertRow(1, []);
      tktruotSheet.mergeCells(1, 1, 1, tktruotCols.length);
      const tktruotTitle = tktruotSheet.getCell('A1');
      tktruotTitle.value = `BẢNG THỐNG KÊ SỐ LƯỢNG HỌC SINH CÓ ĐIỂM DƯỚI 5 THEO LỚP - ${kyThi.toUpperCase()} ${khoiLop.toUpperCase()} - NĂM HỌC ${namHoc}`;
      tktruotTitle.font = { bold: true, size: 12 };
      
      const tktruotHeaderRow = tktruotSheet.getRow(3);
      tktruotHeaders.forEach((header, idx) => {
        const cell = tktruotHeaderRow.getCell(idx + 1);
        cell.value = header;
        cell.font = { bold: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      
      let tktruotCurrentRow = 4;
      const totalRowData: Record<string, number> = {};
      subjects.forEach(sub => totalRowData[sub] = 0);
      let totalHsTruotAll = 0;
      
      classes.forEach(cls => {
        if (cls === 'TSTD') return;
        
        const clsStudents = validStudentsForAvg.filter(s => s.lop === cls);
        
        let hsTruotCls = 0;
        clsStudents.forEach(student => {
          if (subjects.some(sub => student.scores[sub] !== undefined && student.scores[sub] !== null && student.scores[sub] < 5)) {
            hsTruotCls++;
          }
        });
        totalHsTruotAll += hsTruotCls;
        
        const row = tktruotSheet.getRow(tktruotCurrentRow);
        row.getCell(1).value = cls;
        row.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        subjects.forEach((sub, idx) => {
          let countSub = 0;
          clsStudents.forEach(student => {
            if (student.scores[sub] !== undefined && student.scores[sub] !== null && student.scores[sub] < 5) {
              countSub++;
            }
          });
          totalRowData[sub] += countSub;
          
          const cell = row.getCell(idx + 2);
          cell.value = countSub;
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          cell.alignment = { horizontal: 'center' };
        });
        
        const cellTotal = row.getCell(tktruotCols.length);
        cellTotal.value = hsTruotCls;
        cellTotal.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cellTotal.alignment = { horizontal: 'center' };
        
        tktruotCurrentRow++;
      });
      
      // TỔNG CỘNG row
      const tcRow = tktruotSheet.getRow(tktruotCurrentRow);
      tcRow.getCell(1).value = 'TỔNG CỘNG';
      tcRow.getCell(1).font = { bold: true };
      tcRow.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      
      subjects.forEach((sub, idx) => {
        const cell = tcRow.getCell(idx + 2);
        cell.value = totalRowData[sub];
        cell.font = { bold: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { horizontal: 'center' };
      });
      
      const tcTotalCell = tcRow.getCell(tktruotCols.length);
      tcTotalCell.value = totalHsTruotAll;
      tcTotalCell.font = { bold: true };
      tcTotalCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      tcTotalCell.alignment = { horizontal: 'center' };
      
      tktruotSheet.columns.forEach((column, index) => {
        if (index === 0) column.width = 8;
        else if (index === tktruotCols.length - 1) column.width = 15;
      });

      // Download main Excel file
      const buffer = await outWorkbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Bang_Diem_Tong_Hop.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);

      // Download ZIP file
      const zipContent = await zip.generateAsync({ type: 'blob' });
      const zipUrl = window.URL.createObjectURL(zipContent);
      const zipA = document.createElement('a');
      zipA.href = zipUrl;
      zipA.download = 'Danh_sach_cac_lop.zip';
      zipA.click();
      window.URL.revokeObjectURL(zipUrl);

    } catch (error) {
      console.error("Error processing files:", error);
      alert("Có lỗi xảy ra trong quá trình xử lý file. Vui lòng kiểm tra lại định dạng file.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#f0f2f5] text-[#1e293b] font-sans selection:bg-[#eff6ff]">
      {/* Header */}
      <header className="h-[80px] bg-white border-b-2 border-[#cbd5e1] flex items-center justify-center px-6 md:px-10 sticky top-0 z-10">
        <div className="text-[20px] md:text-[24px] font-[800] tracking-[-0.5px] text-[#1e40af] flex items-center justify-center">
          THỐNG KÊ ĐIỂM THI - THCS MỸ XÁ
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8 py-8 w-full">
        <div className="space-y-6">
              <div className="bg-white rounded-lg border border-[#cbd5e1] p-6 shadow-sm overflow-hidden mb-6">
                <div className="space-y-8">
                  {/* Bước 1 */}
                  <div>
                    <h3 className="text-[14px] font-[600] text-[#1e40af] mb-3">Bước 1: Chọn kỳ thi:</h3>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                       <div>
                         <label className="block text-[11px] font-[500] text-[#64748b] mb-1">Kỳ thi</label>
                         <select 
                           value={kyThi} 
                           onChange={(e) => setKyThi(e.target.value)} 
                           className="w-full border border-[#cbd5e1] rounded-md px-3 py-1.5 text-[13px] text-[#1e293b] focus:outline-none focus:ring-1 focus:ring-[#1e40af] bg-white cursor-pointer"
                         >
                           {kyThiOptions.map(opt => (
                             <option key={opt} value={opt}>{opt}</option>
                           ))}
                         </select>
                       </div>
                       <div>
                         <label className="block text-[11px] font-[500] text-[#64748b] mb-1">Năm học</label>
                         <select 
                           value={namHoc} 
                           onChange={(e) => setNamHoc(e.target.value)} 
                           className="w-full border border-[#cbd5e1] rounded-md px-3 py-1.5 text-[13px] text-[#1e293b] focus:outline-none focus:ring-1 focus:ring-[#1e40af] bg-white cursor-pointer"
                         >
                           {namHocOptions.map(opt => (
                             <option key={opt} value={opt}>{opt}</option>
                           ))}
                         </select>
                       </div>
                       <div>
                         <label className="block text-[11px] font-[500] text-[#64748b] mb-1">Khối lớp</label>
                         <select 
                           value={khoiLop} 
                           onChange={(e) => setKhoiLop(e.target.value)} 
                           className="w-full border border-[#cbd5e1] rounded-md px-3 py-1.5 text-[13px] text-[#1e293b] focus:outline-none focus:ring-1 focus:ring-[#1e40af] bg-white cursor-pointer"
                         >
                           {khoiLopOptions.map(opt => (
                             <option key={opt} value={opt}>{opt}</option>
                           ))}
                         </select>
                       </div>
                    </div>
                  </div>

                  <hr className="border-[#e2e8f0]" />

                  {/* Bước 2 */}
                  <div>
                    <h3 className="text-[14px] font-[600] text-[#1e40af] mb-3">Bước 2: Tải lên các tệp Excel</h3>
                    <div className={cn(
                      "rounded-lg border border-dashed transition-all duration-200 p-10 text-center cursor-pointer",
                      isDragging ? "border-[#1e40af] bg-[#eff6ff]" : "border-[#cbd5e1] bg-[#f8fafc] hover:border-[#94a3b8]"
                    )}
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                        onDrop={handleDrop}
                        onClick={() => fileInputRef.current?.click()}
                    >
                      <input
                        type="file"
                        multiple
                        accept=".xlsx, .xls"
                        className="hidden"
                        ref={fileInputRef}
                        onChange={handleFileSelect}
                      />
                      <UploadCloud className="w-8 h-8 text-[#94a3b8] mx-auto mb-3" />
                      <p className="text-[13px] text-[#475569]">
                        Nhấn để tải lên hoặc kéo và thả
                      </p>
                      <p className="text-[11px] text-[#94a3b8] mt-1">
                        Tải lên nhiều tệp Excel (.xlsx, .xls)
                      </p>
                    </div>

                    {files.length > 0 && (
                      <div className="mt-4">
                        <div className="flex items-center justify-between mb-2">
                          <span className="text-[12px] font-[500] text-[#1e293b]">{files.length} tệp được chọn</span>
                          <button 
                            onClick={() => setFiles([])}
                            className="text-red-500 hover:text-red-600 text-[12px] font-[500]"
                          >
                            Xóa tất cả
                          </button>
                        </div>
                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2">
                          {files.map((file, idx) => (
                            <div key={idx} className="flex items-center justify-between p-2 bg-white border border-[#cbd5e1] rounded-md group">
                              <div className="flex items-center gap-2 overflow-hidden">
                                <div className="w-1.5 h-1.5 rounded-full bg-[#10b981] shrink-0" />
                                <span className="text-[12px] text-[#1e293b] truncate">{file.name}</span>
                              </div>
                              <button
                                onClick={(e) => { e.stopPropagation(); removeFile(idx); }}
                                className="text-[#94a3b8] hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity p-0.5"
                                title="Xóa"
                              >
                                &times;
                              </button>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>

                  <hr className="border-[#e2e8f0]" />

                  {/* Bước 3 */}
                  <div>
                    <h3 className="text-[14px] font-[600] text-[#1e40af] mb-4">Bước 3: Xử lý dữ liệu</h3>
                    <div className="flex justify-center pb-2">
                      <button
                        onClick={processExcelFiles}
                        disabled={files.length === 0 || isProcessing}
                        className={cn(
                          "px-6 py-2.5 rounded text-[13px] font-[600] flex items-center gap-2 transition-all border-none relative overflow-hidden",
                          files.length === 0
                            ? "bg-[#94a3b8] text-white opacity-60 cursor-not-allowed" // inactive
                            : isProcessing
                              ? "bg-[#94a3b8] text-white opacity-80 cursor-wait pointer-events-none"
                              : "bg-[#94a3b8] text-white hover:bg-[#64748b] bg-[#8b9cb6] shadow-[0_2px_4px_rgba(0,0,0,0.1)]" // Default active style matching user image
                        )}
                      >
                        {isProcessing ? (
                          <>
                            <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                            Đang xử lý...
                          </>
                        ) : (
                          <>
                            <UploadCloud className="w-4 h-4 text-white opacity-90" />
                            Xử lý tệp tin
                          </>
                        )}
                        {/* Shimmer effect for generating */}
                        {isProcessing && (
                          <div className="absolute inset-0 bg-white/20 -skew-x-[20deg] w-[50%] animate-[shimmer_1.5s_infinite] -left-[50%]" />
                         )}
                      </button>
                    </div>
                  </div>
                </div>
              </div>
              
              {/* Dashboard Result */}
              {dashboardData && (
                <>
                  <div className="bg-white rounded-lg border border-[#cbd5e1] p-6 shadow-sm overflow-hidden mb-6">
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    
                    {/* Left: Phổ điểm môn học */}
                    <div className="border border-[#e2e8f0] rounded-lg p-4">
                      <div className="flex justify-between items-start mb-6">
                        <div>
                          <h3 className="font-[600] text-[#1e293b] text-[15px]">Phổ điểm theo môn học</h3>
                          {dashSubject && dashboardData.overallAvg[dashSubject] !== undefined && (
                            <div className="text-[14px] text-[#64748b] mt-1 font-[500]">
                              ĐIỂM TRUNG BÌNH: <span className="text-[#1e40af] font-[700]">{dashboardData.overallAvg[dashSubject]}</span>
                            </div>
                          )}
                        </div>
                        <select 
                           value={dashSubject} 
                           onChange={(e) => setDashSubject(e.target.value)}
                           className="border border-[#cbd5e1] rounded p-1.5 text-[13px] text-[#1e293b] outline-none focus:ring-1 focus:ring-[#1e40af]"
                        >
                           {dashboardData.subjects.map(s => <option key={s} value={s}>{s}</option>)}
                        </select>
                      </div>
                      
                      <div className="h-72">
                        {dashSubject && dashboardData.overall[dashSubject] ? (
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dashboardData.overall[dashSubject]} margin={{ top: 10, right: 10, left: -20, bottom: 20 }}>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} />
                              <XAxis dataKey="name" tick={{fontSize: 10, fill: '#64748b'}} interval={0} angle={-35} textAnchor="end" height={60} />
                              <YAxis tick={{fontSize: 11, fill: '#64748b'}} label={{ value: 'Số lượng', angle: -90, position: 'insideLeft', offset: 20, fontSize: 12, fill: '#64748b' }} />
                              <Tooltip cursor={{fill: '#f1f5f9'}} contentStyle={{fontSize: '12px', borderRadius: '8px', border: '1px solid #cbd5e1'}} />
                              <Bar dataKey="count" radius={[2, 2, 0, 0]}>
                                {dashboardData.overall[dashSubject].map((entry, index) => {
                                  // distinct colors mapping for bars
                                  const colors = ['#ef4444', '#f97316', '#f59e0b', '#eab308', '#84cc16', '#22c55e', '#10b981', '#14b8a6', '#06b6d4', '#3b82f6', '#8b5cf6'];
                                  return <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />;
                                })}
                              </Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        ) : (
                          <div className="h-full flex items-center justify-center text-[#94a3b8] text-[13px]">
                             Không có dữ liệu môn học để hiển thị dashboard.
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Right: Phổ điểm theo lớp */}
                    <div className="border border-[#e2e8f0] rounded-lg p-4 flex flex-col">
                      <div className="flex justify-between items-center mb-2">
                        <h3 className="font-[600] text-[#1e293b] text-[15px]">Phổ điểm theo lớp</h3>
                      </div>
                      
                      <div className="flex overflow-x-auto border-b border-[#e2e8f0] mb-6 pb-1 min-h-[40px] gap-2">
                         {dashboardData.classes.map(cls => (
                            <button
                              key={cls}
                              className={cn(
                                "px-3 py-1 whitespace-nowrap text-[13px] font-[500] bg-transparent border-b-2 transition-colors", 
                                dashClass === cls ? "border-[#3b82f6] text-[#3b82f6]" : "border-transparent text-[#64748b] hover:text-[#1e293b]"
                              )}
                              onClick={() => setDashClass(cls)}
                            >
                              Lớp {cls}
                            </button>
                         ))}
                      </div>

                      <div className="text-center font-[700] text-[15px] text-[#64748b] mb-4">
                        Biểu đồ điểm Lớp {dashClass}
                      </div>
                      
                      <div className="h-64 mt-auto">
                        {dashClass && dashboardData.byClassGrouped[dashClass] ? (
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dashboardData.byClassGrouped[dashClass]} margin={{ top: 10, right: 10, left: -20, bottom: 20 }}>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} />
                              <XAxis dataKey="name" tick={{fontSize: 10, fill: '#64748b'}} interval={0} angle={-35} textAnchor="end" height={60}>
                                <Label value="Thang điểm" offset={-15} position="insideBottom" fontSize={11} fill="#64748b" />
                              </XAxis>
                              <YAxis tick={{fontSize: 11, fill: '#64748b'}} label={{ value: 'Số lượng', angle: -90, position: 'insideLeft', offset: 20, fontSize: 12, fill: '#64748b' }} />
                              <Tooltip cursor={{fill: '#f1f5f9'}} contentStyle={{fontSize: '12px', borderRadius: '8px', border: '1px solid #cbd5e1'}} />
                              <Legend wrapperStyle={{fontSize: '12px', paddingTop: '20px'}} />
                              {dashboardData.subjects.map((sub, idx) => {
                                 const colors = ['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#06b6d4', '#84cc16'];
                                 return <Bar key={sub} dataKey={sub} name={sub} fill={colors[idx % colors.length]} radius={[2, 2, 0, 0]} />;
                              })}
                            </BarChart>
                          </ResponsiveContainer>
                        ) : (
                          <div className="h-full flex items-center justify-center text-[#94a3b8] text-[13px]">
                             Không có dữ liệu
                          </div>
                        )}
                      </div>
                    </div>

                  </div>
                </div>

                {/* Student Table List */}
                <div className="bg-white rounded-lg border border-[#cbd5e1] p-6 shadow-sm overflow-hidden min-h-[400px]">
                  <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
                    <h3 className="font-[600] text-[#1e293b] text-[15px]">Danh sách điểm học sinh</h3>
                    <div className="relative w-full sm:w-auto">
                      <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-[#94a3b8]" />
                      <input 
                        type="text" 
                        placeholder="Tìm kiếm theo họ tên hoặc SBD..." 
                        value={searchTerm} 
                        onChange={(e) => setSearchTerm(e.target.value)}
                        className="w-full sm:w-64 pl-9 pr-4 py-2 border border-[#cbd5e1] rounded-lg text-[13px] outline-none focus:ring-1 focus:ring-[#3b82f6] transition-shadow placeholder:text-[#94a3b8]"
                      />
                    </div>
                  </div>
                  
                  <div className="overflow-x-auto rounded-lg border border-[#e2e8f0]">
                    <table className="w-full text-left border-collapse text-[13px] whitespace-nowrap">
                      <thead>
                        <tr className="bg-[#f8fafc] text-[#475569] font-[600]">
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0]">STT</th>
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0]">Họ và tên</th>
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0]">Ngày sinh</th>
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0]">Số BD</th>
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0]">Lớp</th>
                          {dashboardData.subjects.map(s => (
                            <th key={s} className="px-4 py-3 border-b border-r border-[#e2e8f0] text-center capitalize">{s}</th>
                          ))}
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0] text-center">Tổng điểm</th>
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0] text-center">XT</th>
                          <th className="px-4 py-3 border-b border-r border-[#e2e8f0] text-center">TBM</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                           const filteredStudents = dashboardData.students.filter(s => {
                             if (!searchTerm) return true;
                             const searchLower = searchTerm.toLowerCase();
                             return (s.hoTen && s.hoTen.toLowerCase().includes(searchLower)) ||
                                    (s.sbd && s.sbd.toString().toLowerCase().includes(searchLower));
                           });
                           
                           if (filteredStudents.length === 0) {
                             return (
                               <tr>
                                 <td colSpan={100} className="px-4 py-8 text-center text-[#64748b]">
                                   Không tìm thấy học sinh nào phù hợp.
                                 </td>
                               </tr>
                             )
                           }
                           
                           return filteredStudents.slice(0, 100).map((s, idx) => ( // limit to 100
                             <tr key={s.sbd || idx} className="border-b border-[#e2e8f0] hover:bg-[#f1f5f9] transition-colors text-[#1e293b]">
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0] text-center">{idx + 1}</td>
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0] font-[500]">{s.hoTen}</td>
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0]">{s.ngaySinh}</td>
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0]">{s.sbd}</td>
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0]">{s.lop}</td>
                               {dashboardData.subjects.map(sub => (
                                 <td key={sub} className="px-4 py-2.5 border-r border-[#e2e8f0] text-center">{s.scores?.[sub] ?? ''}</td>
                               ))}
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0] text-center font-[600] text-[#1e40af]">{s.tongDiem ?? ''}</td>
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0] text-center font-[600] text-[#059669]">{s.xt ?? ''}</td>
                               <td className="px-4 py-2.5 border-r border-[#e2e8f0] text-center font-[600] text-[#d97706]">{s.tbm ?? ''}</td>
                             </tr>
                           ));
                        })()}
                      </tbody>
                    </table>
                  </div>
                  {dashboardData.students.length > 100 && (
                    <div className="mt-4 text-center text-[#64748b] text-[13px]">
                      {searchTerm ? `Hiển thị tối đa 100 kết quả tìm kiếm.` : 'Đang giới hạn hiển thị 100 học sinh đầu tiên để tối ưu hiệu suất. Vui lòng nhập ô tìm kiếm để xem thêm.'}
                    </div>
                  )}
                </div>
                </>
              )}
        </div>
      </main>
    </div>
  );
}
