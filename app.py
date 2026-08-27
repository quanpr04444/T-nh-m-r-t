import streamlit as st
import pandas as pd
from datetime import time, datetime, timedelta
import io

# --- THIẾT LẬP GIAO DIỆN WEB ---
st.set_page_config(page_title="Tính Mã Rớt Ca", layout="wide")
st.title("🚛 Công cụ tính Mã Rớt Ca Tuyến Chính")

# --- CẤU HÌNH ---
SPECIAL_90M_WAREHOUSES = ['024GW0', '229GW0', '221GW0', '292GW0', '028GW6', '028GW0']

DEFAULT_SHIFT_MAP = {
    '024GW0': {1: '04:30', 2: '11:30'}, '024GW3': {1: '04:30', 2: '12:15'},
    '024GW7': {1: '04:45', 2: '12:15'}, '024GW8': {1: '04:30', 2: '12:00'},
    '028GW0': {1: '04:00', 2: '10:00'}, '028GW5': {1: '04:00', 2: '12:00'},
    '028GW6': {1: '03:30', 2: '11:30'}, '205GW0': {1: '06:10'},
    '208GW0': {1: '04:30', 2: '11:00'}, '210GW0': {1: '05:25', 2: '10:00'},
    '212GW0': {1: '04:30', 2: '10:30'}, '214GW0': {1: '05:30', 2: '10:00'},
    '215GW0': {1: '00:01'}, '216GW0': {1: '05:30', 2: '09:40'},
    '218GW0': {1: '06:30', 2: '11:20'}, '221GW0': {1: '04:15', 2: '11:00'},
    '222GW2': {1: '05:30', 2: '11:40'}, '225GW1': {1: '03:00', 2: '05:00'},
    '229GW0': {1: '04:30', 2: '10:45'}, '236GW0': {1: '03:20', 2: '10:30'},
    '237GW0': {1: '03:00', 2: '08:30'}, '238GW0': {1: '04:00', 2: '08:40'},
    '251GW0': {1: '04:50', 2: '11:40'}, '256GW0': {1: '04:30', 2: '10:30'},
    '258GW0': {1: '04:00', 2: '10:30'}, '262GW0': {1: '05:15', 2: '10:30'},
    '269GW0': {1: '05:40', 2: '11:30'}, '272GW0': {1: '04:30', 2: '09:00'},
    '274GW0': {1: '05:30', 2: '11:00'}, '291GW0': {1: '03:30', 2: '09:00'},
    '292GW0': {1: '03:30', 2: '08:30'}, '296GW1': {1: '04:00', 2: '09:00'},
    '297GW0': {1: '04:30', 2: '09:30'}
}

# Sử dụng Session State để lưu file cấu hình ngầm trên môi trường Web
if 'shift_map' not in st.session_state:
    st.session_state.shift_map = DEFAULT_SHIFT_MAP.copy()


def update_shift_config(ca_phat_file):
    try:
        df_ca = pd.read_excel(ca_phat_file)
        req_cols = ['Mã TTTC đích đến', 'Bưu cục xe đến', 'Ca phân phối hàng', 'Thời gian phát hàng']

        if set(req_cols).issubset(df_ca.columns):
            mapping = df_ca[req_cols].dropna(
                subset=['Mã TTTC đích đến', 'Ca phân phối hàng', 'Thời gian phát hàng']).drop_duplicates()
            new_shift_map = {}
            for _, row in mapping.iterrows():
                code = str(row['Mã TTTC đích đến']).strip().upper()
                if not code or code == 'NAN': continue

                try:
                    shift = int(float(row['Ca phân phối hàng']))
                except:
                    continue

                t_val = row['Thời gian phát hàng']
                if isinstance(t_val, time):
                    time_str = t_val.strftime('%H:%M')
                else:
                    parts = str(t_val).strip().split(':')
                    if len(parts) >= 2:
                        time_str = f"{int(parts[0]):02d}:{int(parts[1]):02d}"
                    else:
                        time_str = str(t_val)

                if code not in new_shift_map:
                    new_shift_map[code] = {}
                new_shift_map[code][shift] = time_str

            st.session_state.shift_map = new_shift_map
            st.success("✅ Đã cập nhật và lưu hệ thống Ca Phát thành công!")
        else:
            st.warning("⚠️ File ca phát tải lên thiếu cột. Hệ thống tiếp tục dùng cấu hình đang có.")
    except Exception as e:
        st.error(f"❌ Lỗi đọc file ca phát: {e}")


def parse_time(t_str):
    parts = str(t_str).strip().split(':')
    return time(int(parts[0]), int(parts[1]))


def convert_df(df):
    """Hàm hỗ trợ nút Tải xuống Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='MaRot')
    processed_data = output.getvalue()
    return processed_data


# --- GIAO DIỆN CHÍNH ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. File Tra Hành Trình (Bắt buộc)")
    file_hanh_trinh = st.file_uploader("Kéo thả file Hành trình vào đây", type=['xlsx', 'xls'], key="ht")

    source_bc = st.text_input("3. Mã BC lấy hàng (Ví dụ: 024GW0)", value="024GW0").strip().upper()

with col2:
    st.subheader("2. Cập Nhật Thời Gian Ca Phát (Tùy chọn)")
    file_ca_phat = st.file_uploader("Kéo thả file Cập nhật ca phát vào đây (Chỉ cần tải 1 lần)", type=['xlsx', 'xls'],
                                    key="cp")
    if file_ca_phat is not None:
        if st.button("Cập nhật Ca Phát vào hệ thống"):
            update_shift_config(file_ca_phat)

st.markdown("---")

if st.button("🚀 TẠO BẢNG", type="primary"):
    if not file_hanh_trinh or not source_bc:
        st.error("⚠️ Vui lòng tải lên file Tra Hành Trình và nhập mã Bưu cục.")
    else:
        with st.spinner("Đang tính toán mã rớt..."):
            try:
                shift_release_map = st.session_state.shift_map
                df = pd.read_excel(file_hanh_trinh, sheet_name='Tra hành trình tuyến chính', dtype=str)

                for col in ['Bưu cục lấy hàng', 'Mã BC đích đến', 'Mã tuyến đường']:
                    if col in df.columns:
                        df[col] = df[col].fillna('TRỐNG').str.strip().str.upper().str.replace(r'\s+', '', regex=True)

                source_bc_clean = source_bc.replace(' ', '').replace('O', '0')
                mask = (df['Bưu cục lấy hàng'] == source_bc_clean) | \
                       (df['Bưu cục lấy hàng'].str.contains(source_bc_clean)) | \
                       (df['Mã tuyến đường'].str.startswith(source_bc_clean))
                routes = df[mask].copy()

                if routes.empty:
                    st.warning(f"⚠️ Không tìm thấy chuyến nào của bưu cục {source_bc}")
                else:
                    routes['Thời gian xe đi quy hoạch_dt'] = pd.to_datetime(routes['Thời gian xe đi quy hoạch'],
                                                                            errors='coerce')
                    routes['Thời gian quy hoạch xe đến_dt'] = pd.to_datetime(routes['Thời gian quy hoạch xe đến'],
                                                                             errors='coerce')
                    routes = routes.dropna(subset=['Thời gian xe đi quy hoạch_dt', 'Thời gian quy hoạch xe đến_dt'])

                    max_date = routes['Thời gian xe đi quy hoạch_dt'].max().date()
                    end_cutoff = datetime.combine(max_date, time(8, 0))
                    start_cutoff = end_cutoff - timedelta(days=1)

                    dest_names = routes.drop_duplicates(subset=['Mã BC đích đến'], keep='first').sort_values(
                        'Mã BC đích đến')
                    results = []

                    for _, drow in dest_names.iterrows():
                        dest_code = str(drow['Mã BC đích đến'])
                        if not dest_code or dest_code == 'TRỐNG': continue

                        dest_name = str(drow.get('Bưu cục hàng đến', dest_code)).strip()
                        sub = routes[routes['Mã BC đích đến'] == dest_code].copy()

                        extraction_mins = 90 if dest_code in SPECIAL_90M_WAREHOUSES else 60

                        shifts = shift_release_map.get(dest_code, {1: '08:00'})
                        s1_limit = parse_time(shifts.get(1, '08:00'))
                        s2_limit = parse_time(shifts[2]) if 2 in shifts else None

                        ca1_routes, ca2_routes = [], []
                        has_valid_time = False

                        for _, r in sub.iterrows():
                            try:
                                arr_dt = r['Thời gian quy hoạch xe đến_dt']
                                has_valid_time = True

                                effective_arr_dt = arr_dt + timedelta(minutes=extraction_mins)
                                d_date = effective_arr_dt.date()
                                s1_dt = datetime.combine(d_date, s1_limit)

                                if s2_limit:
                                    s2_dt = datetime.combine(d_date, s2_limit)
                                    if effective_arr_dt <= s1_dt:
                                        ca1_routes.append(r)
                                    elif s1_dt < effective_arr_dt <= s2_dt:
                                        ca2_routes.append(r)
                                    else:
                                        ca1_routes.append(r)
                                else:
                                    ca1_routes.append(r)
                            except:
                                continue

                        if not has_valid_time:
                            results.append(
                                {'Mã TTTC đích': dest_code, 'Tên trạm sau': dest_name, 'Ca 1 (Giờ đi)': 'LỖI',
                                 'Ca 2 (Giờ đi)': 'LỖI'})
                            continue

                        rot_ca1 = ""
                        if ca1_routes:
                            valid_ca1 = [r for r in ca1_routes if
                                         start_cutoff < r['Thời gian xe đi quy hoạch_dt'] <= end_cutoff]
                            if valid_ca1:
                                rot_ca1 = max(valid_ca1, key=lambda x: x['Thời gian xe đi quy hoạch_dt'])[
                                    'Thời gian xe đi quy hoạch_dt'].strftime("%H:%M")

                        rot_ca2 = ""
                        if ca2_routes:
                            rot_ca2 = max(ca2_routes, key=lambda x: x['Thời gian quy hoạch xe đến_dt'])[
                                'Thời gian xe đi quy hoạch_dt'].strftime("%H:%M")

                        results.append({'Mã TTTC đích': dest_code, 'Tên trạm sau': dest_name, 'Ca 1 (Giờ đi)': rot_ca1,
                                        'Ca 2 (Giờ đi)': rot_ca2})

                    df_res = pd.DataFrame(results)

                    st.success("🎉 Đã tính toán thành công!")
                    st.dataframe(df_res, use_container_width=True)

                    # Nút tải xuống file kết quả
                    excel_data = convert_df(df_res)
                    st.download_button(
                        label="📥 Tải xuống File Excel",
                        data=excel_data,
                        file_name=f'MaRot_{source_bc}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra trong quá trình tính toán: {e}")
