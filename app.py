import streamlit as st
import pandas as pd
import datetime
import itertools
import heapq

# ページ設定
st.set_page_config(
    page_title="三陸じぶんの旅パス",
    page_icon="🚃",
    layout="wide"
)

# ※ご自身の環境に合わせてファイル名を確認してください
FILE_PATH = "統合版・2026年三陸鉄道周辺ダイヤ_3.xlsx"

# ==========================================
# 1. データ読み込み＆探索ロジック
# ==========================================
def parse_time_to_min(val):
    if pd.isna(val) or val == '' or val == '-':
        return None
    if isinstance(val, (datetime.time, pd.Timestamp, datetime.datetime)):
        return val.hour * 60 + val.minute
    if isinstance(val, str):
        parts = val.strip().split(':')
        if len(parts) >= 2 and parts[0].isdigit():
            return int(parts[0]) * 60 + int(parts[1])
    return None

def min_to_str(m):
    if m is None:
        return "--:--"
    return f"{int(m)//60:02d}:{int(m)%60:02d}"

def is_service_available(row, target_dt):
    weekday = target_dt.weekday()
    is_holiday = weekday in [5, 6]
    
    service_type = str(row['運行区分']).strip() if pd.notna(row['運行区分']) else '毎日'
    if service_type == '平日' and is_holiday:
        return False
    if service_type == '土休日' and not is_holiday:
        return False
        
    ex_days = str(row['除外日']) if pd.notna(row['除外日']) else ''
    if '火曜' in ex_days and weekday == 1:
        if '7-8月除く' in ex_days and target_dt.month in [7, 8]:
            pass
        else:
            return False
            
    if pd.notna(row['運行開始日']) and str(row['運行開始日']).strip() != '':
        start_d = pd.to_datetime(row['運行開始日']).date()
        if target_dt < start_d:
            return False
    if pd.notna(row['運行終了日']) and str(row['運行終了日']).strip() != '':
        end_d = pd.to_datetime(row['運行終了日']).date()
        if target_dt > end_d:
            return False
            
    return True

@st.cache_data
def load_data(filepath, target_date_str):
    xls = pd.ExcelFile(filepath)
    target_dt = datetime.datetime.strptime(target_date_str, "%Y-%m-%d").date()
    all_legs = []
    
    sheet_line_map = {
        '三陸鉄道_北行': '三陸鉄道',
        '三陸鉄道_南行': '三陸鉄道',
        '龍泉洞バス_西行': '岩泉町民バス',
        '龍泉洞バス_東行': '岩泉町民バス',
        '浄土ヶ浜バス_往路': '岩手県北バス',
        '浄土ヶ浜バス_復路': '岩手県北バス',
        '宮古うみねこ丸便': '宮古うみねこ丸',
        '北山崎断崖クルーズ': '北山崎断崖クルーズ',
        '田野畑観光タクシー_行き': '田野畑観光タクシー',
        '田野畑観光タクシー_帰り': '田野畑観光タクシー',
        '普代村営バス_往路': '普代村営バス',
        '普代村営バス_復路': '普代村営バス',
    }
    
    for sname, line_name in sheet_line_map.items():
        if sname not in xls.sheet_names:
            continue
        df = pd.read_excel(filepath, sheet_name=sname)
        attr_cols = ['便ID', '便名', '運行区分', '運行開始日', '運行終了日', '除外日', '運行日']
        stop_cols = [c for c in df.columns if c not in attr_cols]
        
        for _, row in df.iterrows():
            if not is_service_available(row, target_dt):
                continue
                
            trip_id = row['便ID']
            trip_name = row['便名']
            
            stop_times = []
            for col in stop_cols:
                val = row[col]
                t_min = parse_time_to_min(val)
                if t_min is not None:
                    raw_stop = str(col).strip()
                    if raw_stop.endswith('_着'):
                        s_name = raw_stop[:-2]
                        kind = 'arr'
                    elif raw_stop.endswith('_発'):
                        s_name = raw_stop[:-2]
                        kind = 'dep'
                    else:
                        s_name = raw_stop
                        kind = 'both'
                    stop_times.append((s_name, kind, t_min))
            
            merged_stops = []
            idx = 0
            while idx < len(stop_times):
                s_name, kind, t = stop_times[idx]
                if kind == 'arr':
                    if idx + 1 < len(stop_times) and stop_times[idx+1][0] == s_name and stop_times[idx+1][1] == 'dep':
                        merged_stops.append((s_name, t, stop_times[idx+1][2]))
                        idx += 2
                    else:
                        merged_stops.append((s_name, t, t))
                        idx += 1
                elif kind == 'dep':
                    merged_stops.append((s_name, t, t))
                    idx += 1
                else:
                    merged_stops.append((s_name, t, t))
                    idx += 1
            
            for i in range(len(merged_stops)):
                for j in range(i + 1, len(merged_stops)):
                    s_from, arr_f, dep_f = merged_stops[i]
                    s_to, arr_t, dep_t = merged_stops[j]
                    if dep_f is not None and arr_t is not None and dep_f <= arr_t:
                        all_legs.append({
                            'line': line_name,
                            'trip_id': trip_id,
                            'trip_name': trip_name,
                            'from_stop': s_from,
                            'dep_time': dep_f,
                            'to_stop': s_to,
                            'arr_time': arr_t
                        })
                        
    df_trans = pd.read_excel(filepath, sheet_name='乗換設定')
    line_name_alias = {
        '龍泉洞バス': '岩泉町民バス',
        '浄土ヶ浜バス': '岩手県北バス'
    }
    df_trans['乗換元路線'] = df_trans['乗換元路線'].replace(line_name_alias)
    df_trans['乗換先路線'] = df_trans['乗換先路線'].replace(line_name_alias)
    
    trans_map = {}
    for _, r in df_trans.iterrows():
        k = (str(r['乗換元路線']).strip(), str(r['乗換元停留所・駅']).strip(),
             str(r['乗換先路線']).strip(), str(r['乗換先停留所・駅']).strip())
        trans_map[k] = int(r['所要時間（分）'])
        
    return all_legs, trans_map

@st.cache_data
def load_fare_data(filepath):
    sanriku_fares = {}
    other_fares = {}
    facility_fares = {}
    
    xls = pd.ExcelFile(filepath)
    if '三鉄_運賃三角表' in xls.sheet_names:
        df_sanriku = pd.read_excel(filepath, sheet_name='三鉄_運賃三角表', index_col=0)
        for row in df_sanriku.index:
            for col in df_sanriku.columns:
                if pd.notna(df_sanriku.at[row, col]):
                    try:
                        r_name = str(row).replace('\u3000', '').replace(' ', '')
                        c_name = str(col).replace('\u3000', '').replace(' ', '')
                        sanriku_fares[(r_name, c_name)] = int(df_sanriku.at[row, col])
                    except:
                        pass
                    
    if '他社_通常料金' in xls.sheet_names:
        df_other = pd.read_excel(filepath, sheet_name='他社_通常料金')
        for _, row in df_other.iterrows():
            operator = str(row['事業者']).strip()
            if operator == '田野畑観光乗合タクシー':
                operator = '田野畑観光タクシー'
                
            f_stop = str(row['出発地']).strip() if pd.notna(row['出発地']) else 'nan'
            t_stop = str(row['到着地']).strip() if pd.notna(row['到着地']) else 'nan'
            
            normal_f = float(row['運賃']) if pd.notna(row['運賃']) else 0
            group_f = float(row['団体']) if pd.notna(row['団体']) else normal_f
            
            if f_stop == 'nan' and t_stop == 'nan':
                facility_fares[operator] = {'normal': normal_f, 'group': group_f}
            else:
                other_fares[(operator, f_stop, t_stop)] = {'normal': normal_f, 'group': group_f}
                other_fares[(operator, t_stop, f_stop)] = {'normal': normal_f, 'group': group_f}
                
    return sanriku_fares, other_fares, facility_fares

def calculate_fares(history, sanriku_fares, other_fares, facility_fares):
    has_ryusendo = any(step.get('type') == 'stay' and step.get('spot') == '龍泉洞前' for step in history)
    
    normal_total = 0
    cost_total = 0
    breakdown = []
    
    if has_ryusendo:
        cost_total += 4000
        breakdown.append("🎟️ **【セット適用】岩泉龍泉洞１日フリーきっぷ: 原価 ¥4,000**")
        
    for step in history:
        if step['type'] == 'ride':
            line = step['line']
            f_stop = step['from_stop']
            t_stop = step['to_stop']
            
            n_fare = 0
            c_fare = 0
            
            if line == '三陸鉄道':
                n_fare = sanriku_fares.get((f_stop, t_stop), 0)
                if not has_ryusendo:
                    c_fare = n_fare
                breakdown.append(f"🚆 [三陸鉄道] {f_stop} ➔ {t_stop} : 通常 ¥{n_fare:,} / 原価 ¥{c_fare:,}")
            else:
                fare_info = other_fares.get((line, f_stop, t_stop))
                if fare_info:
                    n_fare = int(fare_info['normal'])
                    if has_ryusendo and line == '岩泉町民バス':
                        c_fare = 0
                    else:
                        c_fare = int(fare_info['group'])
                else:
                    n_fare = 0
                    c_fare = 0
                breakdown.append(f"🚌 [{line}] {f_stop} ➔ {t_stop} : 通常 ¥{n_fare:,} / 原価 ¥{c_fare:,}")
            
            normal_total += n_fare
            cost_total += c_fare
            
        elif step['type'] == 'stay':
            spot = step['spot']
            if spot == '龍泉洞前':
                f_info = facility_fares.get('龍泉洞')
                if f_info:
                    n_fare = int(f_info['normal'])
                    if not has_ryusendo:
                        c_fare = int(f_info['group'])
                    else:
                        c_fare = 0
                else:
                    n_fare = 0
                    c_fare = 0
                breakdown.append(f"🏰 [施設入場] {spot} : 通常 ¥{n_fare:,} / 原価 ¥{c_fare:,}")
                normal_total += n_fare
                cost_total += c_fare
                        
    # 販売価格 (1割増し・10円単位で四捨五入)
    sales_price = int(round(cost_total * 1.1 / 10) * 10)
    
    return int(normal_total), int(cost_total), sales_price, breakdown

def find_routes_point_to_point(legs, trans_map, start_stop, start_time_min, target_stop, max_transfers=4):
    counter = 0
    queue = [(start_time_min, 0, counter, start_stop, None, [])]
    best_time = {}
    found_routes = []
    
    while queue:
        cur_time, num_trans, _, cur_stop, cur_line, history = heapq.heappop(queue)
        
        if cur_stop == target_stop:
            found_routes.append((cur_time, history))
            if len(found_routes) >= 3:
                break
            continue
            
        if num_trans > max_transfers:
            continue
            
        state_key = (cur_stop, cur_line)
        if state_key in best_time and best_time[state_key] <= cur_time:
            continue
        best_time[state_key] = cur_time
        
        possible_starts = [(cur_stop, cur_line, cur_time, None)]
        for (f_line, f_stop, t_line, t_stop), dur in trans_map.items():
            if (cur_line is None or f_line == cur_line) and f_stop == cur_stop:
                trans_event = {
                    'type': 'transfer',
                    'from_stop': f_stop, 'from_line': f_line,
                    'to_stop': t_stop, 'to_line': t_line,
                    'duration': dur, 'start_time': cur_time, 'end_time': cur_time + dur
                }
                possible_starts.append((t_stop, t_line, cur_time + dur, trans_event))
                if t_stop == target_stop:
                    new_hist = list(history) + [trans_event]
                    counter += 1
                    heapq.heappush(queue, (cur_time + dur, num_trans, counter, t_stop, t_line, new_hist))
                
        for p_stop, p_line, ready_time, trans_info in possible_starts:
            for leg in legs:
                if leg['from_stop'] == p_stop and (p_line is None or leg['line'] == p_line):
                    if leg['dep_time'] >= ready_time:
                        if history and history[-1].get('type') == 'ride' and history[-1].get('trip_id') == leg['trip_id']:
                            continue
                        new_history = list(history)
                        if trans_info:
                            new_history.append(trans_info)
                        new_history.append({
                            'type': 'ride',
                            'line': leg['line'],
                            'trip_id': leg['trip_id'],
                            'trip_name': leg['trip_name'],
                            'from_stop': leg['from_stop'],
                            'dep_time': leg['dep_time'],
                            'to_stop': leg['to_stop'],
                            'arr_time': leg['arr_time']
                        })
                        next_num_trans = num_trans + (1 if cur_line is not None and cur_line != leg['line'] else 0)
                        counter += 1
                        heapq.heappush(queue, (leg['arr_time'], next_num_trans, counter, leg['to_stop'], leg['line'], new_history))

    return found_routes

def plan_tour(legs, trans_map, start_stop, start_time_str, goal_stop, spots_with_stay, sanriku_fares, other_fares, facility_fares):
    start_time_min = parse_time_to_min(start_time_str)
    all_perms = list(itertools.permutations(spots_with_stay))
    successful_plans = []
    
    for perm in all_perms:
        cur_stop = start_stop
        cur_time = start_time_min
        full_history = []
        is_possible = True
        
        for spot_name, stay_min in perm:
            routes = find_routes_point_to_point(legs, trans_map, cur_stop, cur_time, spot_name)
            if not routes:
                is_possible = False
                break
            
            arr_time, route_hist = routes[0]
            full_history.extend(route_hist)
            leave_time = arr_time + stay_min
            full_history.append({
                'type': 'stay',
                'spot': spot_name,
                'stay_min': stay_min,
                'arr_time': arr_time,
                'dep_time': leave_time
            })
            cur_stop = spot_name
            cur_time = leave_time
            
        if not is_possible:
            continue
            
        return_routes = find_routes_point_to_point(legs, trans_map, cur_stop, cur_time, goal_stop)
        if not return_routes:
            continue
            
        final_arr_time, ret_hist = return_routes[0]
        full_history.extend(ret_hist)
        
        total_duration = final_arr_time - start_time_min
        order_names = " → ".join([s[0] for s in perm])
        
        fares_calc = calculate_fares(full_history, sanriku_fares, other_fares, facility_fares)
        
        successful_plans.append({
            'order': order_names,
            'final_arr_time': final_arr_time,
            'total_duration': total_duration,
            'history': full_history,
            'fares': fares_calc[:3],
            'breakdown': fares_calc[3]
        })
        
    successful_plans.sort(key=lambda x: x['final_arr_time'])
    return successful_plans

# ==========================================
# 2. UI 表示
# ==========================================
st.title("🚃 三陸じぶんの旅パス")
st.caption("行きたい場所を選んで、ルート検索からチケット購入まで")

with st.expander("⚙️ **旅行条件・訪問地を設定する**", expanded=True):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        travel_date = st.date_input("出発日", datetime.date.today())
    with col_d2:
        start_time = st.time_input("出発希望時刻", datetime.time(8, 30))
        
    date_str = travel_date.strftime("%Y-%m-%d")
    start_time_str = start_time.strftime("%H:%M")

    station_options = ["宮古", "久慈", "盛", "釜石", "岩泉小本", "田野畑", "普代"]
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        start_station = st.selectbox("出発駅（起点）", station_options, index=0)
    with col_s2:
        goal_station = st.selectbox("到着駅（最終目的地）", station_options, index=0)

    st.markdown("##### 📍 訪問スポットを選択")
    available_spots = {
        "龍泉洞前": {"label": "龍泉洞（岩泉小本駅接続）", "default": 60},
        "奥浄土ヶ浜": {"label": "奥浄土ヶ浜（宮古駅接続）", "default": 40},
        "出崎ふ頭": {"label": "出崎ふ頭・うみねこ丸乗船場（宮古駅接続）", "default": 40},
        "北山崎展望台": {"label": "北山崎展望台（田野畑駅/普代駅接続）", "default": 45},
        "島越港": {"label": "島越港・断崖クルーズ（島越駅接続）", "default": 60},
    }

    selected_spots_with_stay = []
    for spot_key, info in available_spots.items():
        c_chk, c_num = st.columns([3, 2])
        with c_chk:
            checked = st.checkbox(info["label"], key=f"chk_{spot_key}")
        with c_num:
            if checked:
                stay = st.number_input(f"滞在(分)", min_value=15, max_value=240, value=info["default"], step=15, key=f"stay_{spot_key}")
                selected_spots_with_stay.append((spot_key, stay))

    search_btn = st.button("🔍 最適ルートを検索する", type="primary", use_container_width=True)

# 検索結果表示
if search_btn:
    if not selected_spots_with_stay:
        st.warning("⚠️ 訪問したい観光スポットを1つ以上選択してください。")
    else:
        with st.spinner("ダイヤと乗り継ぎ・運賃を最適化計算中..."):
            try:
                legs, trans_map = load_data(FILE_PATH, date_str)
                sanriku_fares, other_fares, facility_fares = load_fare_data(FILE_PATH)
                plans = plan_tour(legs, trans_map, start_station, start_time_str, goal_station, selected_spots_with_stay, sanriku_fares, other_fares, facility_fares)
                
                if not plans:
                    st.error(f"❌ {start_station}発 ➔ {goal_station}着 で当日中に移動できるルートが見つかりませんでした。出発時刻や滞在時間を調整してください。")
                else:
                    st.success(f"🎉 **{len(plans)} 件**のルートが見つかりました！")
                    
                    for idx, p in enumerate(plans, 1):
                        n_fare, c_fare, s_price = p['fares']
                        
                        with st.container(border=True):
                            st.markdown(f"### ⭐ プラン {idx}：{p['order']}")
                            
                            # --- 料金サマリ表示 ---
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("① 個別運賃の積上額", f"¥{n_fare:,}")
                            with col2:
                                st.metric("② 手配原価（セット適用）", f"¥{c_fare:,}")
                            with col3:
                                st.metric("③ 🎉 販売価格", f"¥{s_price:,}")
                                
                            st.info(f"⏱ **総所要時間**: {p['total_duration']//60}時間{p['total_duration']%60}分 ｜ **区間**: {start_station} ({start_time_str}発) ➔ {goal_station} (**{min_to_str(p['final_arr_time'])}着**)")
                            
                            with st.expander("💴 運賃計算の内訳を見る（0円の場合は設定不一致）"):
                                for b in p['breakdown']:
                                    st.write(b)
                            
                            st.markdown("##### 📍 ルート詳細")
                            for step in p['history']:
                                if step['type'] == 'ride':
                                    st.markdown(f"🚆 **[{step['line']}]** `{step['from_stop']}` (**{min_to_str(step['dep_time'])}発**) ➔ `{step['to_stop']}` (**{min_to_str(step['arr_time'])}着**)")
                                elif step['type'] == 'transfer':
                                    st.caption(f" 🚶 **徒歩・乗換 {step['duration']}分**: {step['from_stop']} ➔ {step['to_stop']}")
                                elif step['type'] == 'stay':
                                    st.success(f"★ **【観光・滞在】 {step['spot']}** （**{min_to_str(step['arr_time'])} 〜 {min_to_str(step['dep_time'])}** / {step['stay_min']}分間）")
                    
                    st.divider()
                    if st.button("🔄 次の検索（条件を再設定する）", use_container_width=True):
                        st.rerun()
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")