import streamlit as st
import pandas as pd
import datetime
import itertools
import heapq
import re
import streamlit.components.v1 as components

# ページ設定
st.set_page_config(
    page_title="三陸海岸・じぶんの旅パス",
    page_icon="🚃",
    layout="wide"
)

# ファイル名
FILE_PATH = "統合版・2026年三陸鉄道周辺ダイヤ_4_2.xlsx"

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

def parse_duration_to_min(val):
    if pd.isna(val):
        return 60
    if isinstance(val, (datetime.time, datetime.datetime)):
        return val.hour * 60 + val.minute
    if isinstance(val, pd.Timestamp):
        return val.hour * 60 + val.minute
    if isinstance(val, (int, float)):
        if val < 1.0:
            total_min = int(round(val * 24 * 60))
            return total_min if total_min > 0 else 60
        return int(val)
    if isinstance(val, str):
        val = val.strip()
        if ':' in val:
            parts = val.split(':')
            return int(parts[0]) * 60 + int(parts[1])
        try:
            return int(float(val))
        except:
            pass
    return 60

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
    
    sheet_line_map = {}
    if '路線マスタ' in xls.sheet_names:
        df_master = pd.read_excel(filepath, sheet_name='路線マスタ')
        col_s = df_master.columns[0]
        col_l = df_master.columns[1]
        for _, r in df_master.iterrows():
            if pd.notna(r[col_s]) and pd.notna(r[col_l]):
                sheet_line_map[str(r[col_s]).strip()] = str(r[col_l]).strip()
                
    # 路線マスタにないクルーズ便を補完
    if '北山崎断崖クルーズ' in xls.sheet_names and '北山崎断崖クルーズ' not in sheet_line_map:
        sheet_line_map['北山崎断崖クルーズ'] = '北山崎断崖クルーズ'
    
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
                            'line': line_name, 'trip_id': trip_id, 'trip_name': trip_name,
                            'from_stop': s_from, 'dep_time': dep_f, 'to_stop': s_to, 'arr_time': arr_t
                        })
                        
    df_trans = pd.read_excel(filepath, sheet_name='乗換設定')
    trans_map = {}
    for _, r in df_trans.iterrows():
        f_line = str(r['乗換元路線']).strip()
        f_stop = str(r['乗換元停留所・駅']).strip()
        t_line = str(r['乗換先路線']).strip()
        t_stop = str(r['乗換先停留所・駅']).strip()
        dur = int(r['所要時間（分）'])
        
        trans_map[(f_line, f_stop, t_line, t_stop)] = dur
        
        reverse_k = (t_line, t_stop, f_line, f_stop)
        if reverse_k not in trans_map:
            trans_map[reverse_k] = dur
        
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
                        
    if '施設料金' in xls.sheet_names:
        df_fac = pd.read_excel(filepath, sheet_name='施設料金')
        for _, row in df_fac.iterrows():
            base_name = str(row['施設・アクティビティ名']).strip()
            district = str(row['地区']).strip() if pd.notna(row['地区']) else 'その他'
            nearest_stop = str(row['最寄り停留所']).strip() if pd.notna(row['最寄り停留所']) else 'nan'
            
            normal_adult = float(row['通常料金']) if pd.notna(row['通常料金']) else 0
            group_adult = float(row['団体・割引料金']) if pd.notna(row['団体・割引料金']) else normal_adult
            
            normal_child = float(row['小人']) if ('小人' in df_fac.columns and pd.notna(row['小人'])) else (normal_adult / 2.0)
            group_child = float(row['小人・団体']) if ('小人・団体' in df_fac.columns and pd.notna(row['小人・団体'])) else normal_child
            
            name = base_name
            if name in facility_fares and nearest_stop != 'nan':
                name = f"{base_name}（{nearest_stop}経由）"
            
            is_fixed_raw = str(row['所要時間固定']).strip() if '所要時間固定' in df_fac.columns and pd.notna(row['所要時間固定']) else ''
            is_fixed = (is_fixed_raw == '固定')
            
            fixed_duration = 60
            if '所要時間' in df_fac.columns and pd.notna(row['所要時間']):
                fixed_duration = parse_duration_to_min(row['所要時間'])
                    
            remark = str(row['備考']) if '備考' in df_fac.columns and pd.notna(row['備考']) else ''
            description = str(row['説明']) if '説明' in df_fac.columns and pd.notna(row['説明']) else ''

            facility_fares[name] = {
                'district': district,
                'normal_adult': normal_adult, 
                'group_adult': group_adult,
                'normal_child': normal_child,
                'group_child': group_child,
                'nearest_stop': nearest_stop,
                'is_fixed': is_fixed,
                'fixed_duration': fixed_duration,
                'remark': remark,
                'description': description
            }
                    
    if '他社_通常料金' in xls.sheet_names:
        df_other = pd.read_excel(filepath, sheet_name='他社_通常料金')
        for _, row in df_other.iterrows():
            operator = str(row['事業者']).strip()
            if operator == '田野畑観光乗合タクシー':
                operator = '田野畑観光タクシー'
                
            f_stop = str(row['出発地']).strip() if pd.notna(row['出発地']) else 'nan'
            t_stop = str(row['到着地']).strip() if pd.notna(row['到着地']) else 'nan'
            
            normal_adult = float(row['運賃']) if pd.notna(row['運賃']) else 0
            group_adult = float(row['団体']) if pd.notna(row['団体']) else normal_adult
            
            normal_child = float(row['小人']) if ('小人' in df_other.columns and pd.notna(row['小人'])) else (normal_adult / 2.0)
            group_child = normal_child
            
            capacity_str = str(row['乗車定員']) if ('乗車定員' in df_other.columns and pd.notna(row['乗車定員'])) else ''
            
            if f_stop != 'nan' and t_stop != 'nan' and f_stop != t_stop:
                other_fares[(operator, f_stop, t_stop)] = {
                    'normal_adult': normal_adult, 'group_adult': group_adult,
                    'normal_child': normal_child, 'group_child': group_child,
                    'capacity': capacity_str
                }
                other_fares[(operator, t_stop, f_stop)] = {
                    'normal_adult': normal_adult, 'group_adult': group_adult,
                    'normal_child': normal_child, 'group_child': group_child,
                    'capacity': capacity_str
                }
                
    return sanriku_fares, other_fares, facility_fares

def calculate_fares(history, sanriku_fares, other_fares, facility_fares, num_adults, num_children):
    has_ryusendo = any(step.get('type') == 'stay' and step.get('activity') == '龍泉洞' for step in history)
    
    normal_adult_total = 0
    normal_child_total = 0
    cost_adult_total = 0
    cost_child_total = 0
    
    breakdown = []
    
    if has_ryusendo:
        cost_adult_total += 4000 * num_adults
        cost_child_total += 2000 * num_children
        detail_txt = []
        if num_adults > 0:
            detail_txt.append(f"大人 ¥4,000 × {num_adults}名 = ¥{4000 * num_adults:,}")
        if num_children > 0:
            detail_txt.append(f"小人 ¥2,000 × {num_children}名 = ¥{2000 * num_children:,}")
        breakdown.append(f"🎟️ **【セット適用】岩泉龍泉洞１日フリーきっぷ: 原価合計 ¥{4000 * num_adults + 2000 * num_children:,}（{' ＋ '.join(detail_txt)}）**")
        
    for step in history:
        if step['type'] == 'ride':
            line = step['line']
            f_stop = step['from_stop']
            t_stop = step['to_stop']
            
            if line == '三陸鉄道':
                n_a = sanriku_fares.get((f_stop, t_stop), 0)
                n_c = ((n_a + 19) // 20) * 10 if n_a > 0 else 0
                c_a = 0 if has_ryusendo else n_a
                c_c = 0 if has_ryusendo else n_c
                
                sum_n = n_a * num_adults + n_c * num_children
                sum_c = c_a * num_adults + c_c * num_children
                
                parts = []
                if num_adults > 0:
                    parts.append(f"大人 ¥{n_a:,}×{num_adults}")
                if num_children > 0:
                    parts.append(f"小人 ¥{n_c:,}×{num_children}")
                breakdown.append(f"🚆 [三陸鉄道] {f_stop} ➔ {t_stop} : 通常 ¥{sum_n:,} ({', '.join(parts)}) / 原価 ¥{sum_c:,}")
                
                normal_adult_total += n_a * num_adults
                normal_child_total += n_c * num_children
                cost_adult_total += c_a * num_adults
                cost_child_total += c_c * num_children
            else:
                f_info = other_fares.get((line, f_stop, t_stop), {})
                n_a = int(f_info.get('normal_adult', 0))
                n_c = int(f_info.get('normal_child', 0))
                
                if has_ryusendo and line == '岩泉町民バス':
                    c_a = 0
                    c_c = 0
                else:
                    c_a = int(f_info.get('group_adult', n_a))
                    c_c = int(f_info.get('group_child', n_c))
                    
                capacity_str = f_info.get('capacity', '')
                total_people = num_adults + num_children
                
                if '1台' in capacity_str or '人' in capacity_str:
                    nums = re.findall(r'\d+', capacity_str)
                    cap = int(nums[0]) if nums else 4
                    vehicles = (total_people + cap - 1) // cap if total_people > 0 else 0
                    
                    sum_n = n_a * vehicles
                    sum_c = c_a * vehicles
                    
                    normal_adult_total += sum_n if num_adults > 0 else 0
                    normal_child_total += 0
                    cost_adult_total += sum_c if num_adults > 0 else 0
                    cost_child_total += 0
                    
                    breakdown.append(f"🚕 [{line}] {f_stop} ➔ {t_stop} : 通常 ¥{sum_n:,} (車両{vehicles}台分) / 原価 ¥{sum_c:,}")
                else:
                    sum_n = n_a * num_adults + n_c * num_children
                    sum_c = c_a * num_adults + c_c * num_children
                    
                    parts = []
                    if num_adults > 0:
                        parts.append(f"大人 ¥{n_a:,}×{num_adults}")
                    if num_children > 0:
                        parts.append(f"小人 ¥{n_c:,}×{num_children}")
                    breakdown.append(f"🚌 [{line}] {f_stop} ➔ {t_stop} : 通常 ¥{sum_n:,} ({', '.join(parts)}) / 原価 ¥{sum_c:,}")
                    
                    normal_adult_total += n_a * num_adults
                    normal_child_total += n_c * num_children
                    cost_adult_total += c_a * num_adults
                    cost_child_total += c_c * num_children
                
        elif step['type'] == 'stay':
            act_name = step.get('activity')
            if act_name:
                f_info = facility_fares.get(act_name, {})
                n_a = int(f_info.get('normal_adult', 0))
                n_c = int(f_info.get('normal_child', 0))
                c_a = int(f_info.get('group_adult', n_a))
                c_c = int(f_info.get('group_child', n_c))
                
                if act_name == '龍泉洞' and has_ryusendo:
                    c_a = 0
                    c_c = 0
                    
                if "サッパ船" in act_name:
                    total_people = num_adults + num_children
                    if total_people <= 2:
                        sum_n = 7600
                        sum_c = 7600
                        
                        share_a = n_a * num_adults
                        share_c = n_c * num_children
                        base_sum = share_a + share_c
                        if base_sum > 0:
                            alloc_a = int(round(7600 * (share_a / base_sum)))
                            alloc_c = 7600 - alloc_a
                        else:
                            alloc_a = 7600 if num_adults > 0 else 0
                            alloc_c = 7600 if num_children > 0 else 0
                            
                        normal_adult_total += alloc_a
                        normal_child_total += alloc_c
                        cost_adult_total += alloc_a
                        cost_child_total += alloc_c
                        
                        parts = []
                        if num_adults > 0:
                            parts.append(f"大人 ¥{alloc_a:,}")
                        if num_children > 0:
                            parts.append(f"小人 ¥{alloc_c:,}")
                        breakdown.append(f"🎫 [体験・入場] {act_name} (特定料金適用) : 特定 ¥7,600 ({', '.join(parts)}) / 原価 ¥7,600")
                    else:
                        sum_n = n_a * num_adults + n_c * num_children
                        sum_c = c_a * num_adults + c_c * num_children
                        normal_adult_total += n_a * num_adults
                        normal_child_total += n_c * num_children
                        cost_adult_total += c_a * num_adults
                        cost_child_total += c_c * num_children
                        
                        parts = []
                        if num_adults > 0:
                            parts.append(f"大人 ¥{n_a:,}×{num_adults}")
                        if num_children > 0:
                            parts.append(f"小人 ¥{n_c:,}×{num_children}")
                        breakdown.append(f"🎫 [体験・入場] {act_name} : 通常 ¥{sum_n:,} ({', '.join(parts)}) / 原価 ¥{sum_c:,}")
                else:
                    sum_n = n_a * num_adults + n_c * num_children
                    sum_c = c_a * num_adults + c_c * num_children
                    normal_adult_total += n_a * num_adults
                    normal_child_total += n_c * num_children
                    cost_adult_total += c_a * num_adults
                    cost_child_total += c_c * num_children
                    
                    if sum_n > 0 or sum_c > 0:
                        parts = []
                        if num_adults > 0:
                            parts.append(f"大人 ¥{n_a:,}×{num_adults}")
                        if num_children > 0:
                            parts.append(f"小人 ¥{n_c:,}×{num_children}")
                        breakdown.append(f"🎫 [体験・入場] {act_name} : 通常 ¥{sum_n:,} ({', '.join(parts)}) / 原価 ¥{sum_c:,}")
                    else:
                        breakdown.append(f"🎫 [景勝地] {act_name} : 入場無料")

    normal_total = normal_adult_total + normal_child_total
    cost_total = cost_adult_total + cost_child_total
    sales_price_adult = int(round(cost_adult_total * 1.1 / 10) * 10) if num_adults > 0 else 0
    sales_price_child = int(round(cost_child_total * 1.1 / 10) * 10) if num_children > 0 else 0
    sales_price_total = sales_price_adult + sales_price_child

    fare_dict = {
        'normal_total': normal_total,
        'normal_adult': normal_adult_total,
        'normal_child': normal_child_total,
        'cost_total': cost_total,
        'cost_adult': cost_adult_total,
        'cost_child': cost_child_total,
        'sales_total': sales_price_total,
        'sales_adult': sales_price_adult,
        'sales_child': sales_price_child
    }
    return fare_dict, breakdown

def find_routes_point_to_point(legs, trans_map, start_stop, start_time_min, target_stop, max_transfers=4):
    counter = 0
    queue = [(start_time_min, 0, counter, start_stop, None, [])]
    best_time = {}
    found_routes = []
    
    while queue:
        cur_time, num_trans, _, cur_stop, cur_line, history = heapq.heappop(queue)
        
        if cur_stop == target_stop:
            found_routes.append((cur_time, history))
            if len(found_routes) >= 2:
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
                
        if cur_stop in ["久慈", "久慈駅"]:
            target_equiv = "久慈駅" if cur_stop == "久慈" else "久慈"
            trans_event = {
                'type': 'transfer',
                'from_stop': cur_stop, 'from_line': cur_line,
                'to_stop': target_equiv, 'to_line': None,
                'duration': 5, 'start_time': cur_time, 'end_time': cur_time + 5
            }
            possible_starts.append((target_equiv, None, cur_time + 5, trans_event))

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

def plan_tour(legs, trans_map, start_stop, start_time_str, goal_stop, spots_with_stay, sanriku_fares, other_fares, facility_fares, num_adults, num_children):
    start_time_min = parse_time_to_min(start_time_str)
    all_perms = list(itertools.permutations(spots_with_stay))
    successful_plans = []
    
    for perm in all_perms:
        cur_stop = start_stop
        cur_time = start_time_min
        full_history = []
        is_possible = True
        
        for spot_key, stay_min, act_name in perm:
            target_stop = spot_key
            if act_name and act_name in facility_fares:
                nearest = facility_fares[act_name].get('nearest_stop')
                if nearest and nearest != 'nan':
                    target_stop = nearest
                    
            # 船・クルーズ等の時刻表を持つアクティビティに対する強制上書き補正
            act_line_target = None
            if act_name:
                if "宮古うみねこ丸" in act_name:
                    act_line_target = "宮古うみねこ丸"
                    if "出崎ふ頭" in act_name:
                        target_stop = "出崎ふ頭"
                    elif "浄土ヶ浜" in act_name:
                        target_stop = "宮古うみねこ丸・浄土ヶ浜"
                elif "北山崎断崖クルーズ" in act_name:
                    act_line_target = "北山崎断崖クルーズ"
                    target_stop = "島越港"
            
            routes = find_routes_point_to_point(legs, trans_map, cur_stop, cur_time, target_stop)
            if not routes:
                is_possible = False
                break
            
            arr_time, route_hist = routes[0]
            full_history.extend(route_hist)
            
            if act_line_target:
                # 該当する便のダイヤを検索
                best_leg = None
                for leg in legs:
                    if leg['line'] == act_line_target and leg['from_stop'] == target_stop:
                        if leg['dep_time'] >= arr_time:
                            if "全体周遊" in act_name and leg['from_stop'] != leg['to_stop']: continue
                            if "移動" in act_name and leg['from_stop'] == leg['to_stop']: continue
                            if "北山崎断崖クルーズ" in act_name and leg['from_stop'] != leg['to_stop']: continue
                            
                            if best_leg is None or leg['dep_time'] < best_leg['dep_time']:
                                best_leg = leg
                
                if best_leg:
                    leave_time = best_leg['arr_time']
                    full_history.append({
                        'type': 'stay',
                        'spot': target_stop,
                        'activity': act_name,
                        'stay_min': leave_time - arr_time,
                        'arr_time': arr_time,
                        'dep_time': leave_time,
                        'timetable_ride': True,
                        'ride_dep': best_leg['dep_time'],
                        'ride_arr': best_leg['arr_time'],
                        'ride_to_stop': best_leg['to_stop']
                    })
                    cur_stop = best_leg['to_stop']
                    cur_time = leave_time
                else:
                    is_possible = False
                    break
            else:
                leave_time = arr_time + stay_min
                full_history.append({
                    'type': 'stay',
                    'spot': target_stop,
                    'activity': act_name,
                    'stay_min': stay_min,
                    'arr_time': arr_time,
                    'dep_time': leave_time
                })
                cur_stop = target_stop
                cur_time = leave_time
            
        if not is_possible:
            continue
            
        return_routes = find_routes_point_to_point(legs, trans_map, cur_stop, cur_time, goal_stop)
        if not return_routes:
            continue
            
        final_arr_time, ret_hist = return_routes[0]
        full_history.extend(ret_hist)
        
        total_duration = final_arr_time - start_time_min
        
        spot_names_only = [s[2] if s[2] else s[0] for s in perm]
        order_names = " ➔ ".join(spot_names_only)
        
        fare_dict, breakdown = calculate_fares(full_history, sanriku_fares, other_fares, facility_fares, num_adults, num_children)
        
        successful_plans.append({
            'order': order_names,
            'final_arr_time': final_arr_time,
            'total_duration': total_duration,
            'history': full_history,
            'fares': fare_dict,
            'breakdown': breakdown
        })
        
    successful_plans.sort(key=lambda x: x['final_arr_time'])
    return successful_plans

# ==========================================
# 2. 状態管理・ルーティング
# ==========================================
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'search'
if 'confirm_plan' not in st.session_state:
    st.session_state.confirm_plan = None
if 'plans_cache' not in st.session_state:
    st.session_state.plans_cache = None
if 'auth_step' not in st.session_state:
    st.session_state.auth_step = 'input_email'
if 'user_info' not in st.session_state:
    st.session_state.user_info = {}
if 'my_tickets' not in st.session_state:
    st.session_state.my_tickets = []
if 'active_ticket_id' not in st.session_state:
    st.session_state.active_ticket_id = None

st.title("🚃 三陸海岸・じぶんの旅パス")

# ==========================================
# 画面描画
# ==========================================

# ------------------------------------------
# 画面1: 認証・ユーザー情報入力 (モックアップ)
# ------------------------------------------
if st.session_state.current_page == 'auth':
    st.markdown("## 🔒 ユーザー情報の入力と認証")
    name = st.text_input("氏名", value=st.session_state.user_info.get('name', ''))
    email = st.text_input("メールアドレス", value=st.session_state.user_info.get('email', ''))

    if st.session_state.auth_step == 'input_email':
        if st.button("認証コードを送信する", type="primary"):
            if name and email:
                st.session_state.user_info = {'name': name, 'email': email}
                st.session_state.auth_step = 'verify_code'
                st.rerun()
            else:
                st.error("氏名とメールアドレスを入力してください。")
        if st.button("⬅️ プラン確認に戻る"):
            st.session_state.current_page = 'confirm'
            st.rerun()
            
    elif st.session_state.auth_step == 'verify_code':
        st.success(f"{st.session_state.user_info['email']} 宛に認証コードを送信しました。")
        st.info("※テスト用ダミー: 「1234」と入力して進んでください。")
        code = st.text_input("4桁の認証コード", max_chars=4)
        if st.button("認証して次へ", type="primary"):
            if code == "1234":
                st.session_state.current_page = 'payment'
                st.rerun()
            else:
                st.error("認証コードが正しくありません。（「1234」を入力してください）")
        if st.button("⬅️ メールアドレスを修正する"):
            st.session_state.auth_step = 'input_email'
            st.rerun()

# ------------------------------------------
# 画面2: クレジットカード決済 (モックアップ)
# ------------------------------------------
elif st.session_state.current_page == 'payment':
    st.markdown("## 💳 クレジットカード決済")
    f = st.session_state.confirm_plan['fares']
    
    st.info(f"**決済金額（販売価格合計）: ¥{f['sales_total']:,}**")
    
    st.text_input("カード番号 (ダミー)")
    col_cc1, col_cc2 = st.columns(2)
    with col_cc1:
        st.text_input("有効期限 (MM/YY)")
    with col_cc2:
        st.text_input("セキュリティコード (CVC)")

    if st.button("決済を完了してパスを発行する", type="primary", use_container_width=True):
        tickets = []
        meta = st.session_state.get('search_meta', {})
        n_adults = meta.get('num_adults', 1)
        n_children = meta.get('num_children', 0)
        has_ryusendo = any(step.get('type') == 'stay' and step.get('activity') == '龍泉洞' for step in st.session_state.confirm_plan['history'])
        try:
            sanriku_fares_ref, other_fares_ref, facility_fares_ref = load_fare_data(FILE_PATH)
        except:
            sanriku_fares_ref, other_fares_ref, facility_fares_ref = {}, {}, {}

        for i, step in enumerate(st.session_state.confirm_plan['history']):
            ticket_id = f"ticket_{i}_{int(datetime.datetime.now().timestamp())}"
            amount_str = ""
            
            if step['type'] == 'ride':
                line = step['line']
                f_stop = step['from_stop']
                t_stop = step['to_stop']
                if line == '三陸鉄道':
                    n_a = sanriku_fares_ref.get((f_stop, t_stop), 0)
                    n_c = ((n_a + 19) // 20) * 10 if n_a > 0 else 0
                    sum_n = n_a * n_adults + n_c * n_children
                    amount_str = f"¥{sum_n:,}"
                else:
                    f_info = other_fares_ref.get((line, f_stop, t_stop), {})
                    n_a = int(f_info.get('normal_adult', 0))
                    n_c = int(f_info.get('normal_child', 0))
                    cap_str = f_info.get('capacity', '')
                    tot_p = n_adults + n_children
                    if '1台' in cap_str or '人' in cap_str:
                        nums = re.findall(r'\d+', cap_str)
                        cap = int(nums[0]) if nums else 4
                        veh = (tot_p + cap - 1) // cap if tot_p > 0 else 0
                        amount_str = f"¥{n_a * veh:,}"
                    else:
                        amount_str = f"¥{(n_a * n_adults + n_c * n_children):,}"
                        
                tickets.append({
                    "id": ticket_id,
                    "type": "ride",
                    "title": f"🚆 {step['line']}",
                    "provider": step['line'],
                    "section": f"{step['from_stop']} ➔ {step['to_stop']}",
                    "time": f"{min_to_str(step['dep_time'])}発 ➔ {min_to_str(step['arr_time'])}着",
                    "amount": amount_str if amount_str else "---",
                    "status": "未使用",
                    "adults": n_adults,
                    "children": n_children
                })
                
            elif step['type'] == 'stay' and step.get('activity'):
                act = step.get('activity')
                f_info = facility_fares_ref.get(act, {})
                n_a = int(f_info.get('normal_adult', 0))
                n_c = int(f_info.get('normal_child', 0))
                tot_p = n_adults + n_children
                if "サッパ船" in act and tot_p <= 2:
                    amount_str = "¥7,600"
                else:
                    sum_n = n_a * n_adults + n_c * n_children
                    amount_str = f"¥{sum_n:,}" if sum_n > 0 else "無料"
                    
                if step.get('timetable_ride'):
                    time_str = f"{min_to_str(step['ride_dep'])}発 ➔ {min_to_str(step['ride_arr'])}着"
                    section_str = f"{step['spot']} ➔ {step.get('ride_to_stop', step['spot'])}"
                else:
                    time_str = f"利用予定: {min_to_str(step['arr_time'])} 〜 {min_to_str(step['dep_time'])}"
                    section_str = "施設・体験入場"
                    
                tickets.append({
                    "id": ticket_id,
                    "type": "activity",
                    "title": f"🎫 {act}",
                    "provider": act,
                    "section": section_str,
                    "time": time_str,
                    "amount": amount_str,
                    "status": "未使用",
                    "adults": n_adults,
                    "children": n_children
                })
                
        st.session_state.my_tickets = tickets
        st.session_state.active_ticket_id = None
        st.session_state.current_page = 'pass'
        st.rerun()
        
    if st.button("⬅️ 戻る"):
        st.session_state.current_page = 'auth'
        st.rerun()

# ------------------------------------------
# 画面3: デジタル乗車券・入場券 (マイパス)
# ------------------------------------------
elif st.session_state.current_page == 'pass':
    st.markdown("## 📱 デジタルパス (チケット一覧)")
    meta = st.session_state.get('search_meta', {})
    
    # 【ビューB: 個別のチケット詳細画面】
    if st.session_state.active_ticket_id is not None:
        t = next((tk for tk in st.session_state.my_tickets if tk['id'] == st.session_state.active_ticket_id), None)
        if t:
            if st.button("⬅️ チケット一覧に戻る"):
                st.session_state.active_ticket_id = None
                st.rerun()
                
            # 時計表示 (スクリーンショット防止のため1/10秒まで表示)
            clock_html = """
            <div style="text-align: center; margin: 15px 0;">
                <div id="clock" style="font-size: 3.5em; font-weight: bold; color: #111; font-family: 'Courier New', Courier, monospace; letter-spacing: 2px;"></div>
            </div>
            <script>
                function updateTime() {
                    const now = new Date();
                    const h = String(now.getHours()).padStart(2, '0');
                    const m = String(now.getMinutes()).padStart(2, '0');
                    const s = String(now.getSeconds()).padStart(2, '0');
                    const ms = String(Math.floor(now.getMilliseconds() / 100)); // 1/10秒
                    document.getElementById('clock').textContent = `${h}:${m}:${s}.${ms}`;
                }
                setInterval(updateTime, 100);
                updateTime();
            </script>
            """
            components.html(clock_html, height=100)
            
            with st.container(border=True):
                st.markdown(f"### {t['title']}")
                st.markdown(f"**利用交通機関・施設:** {t['provider']}")
                st.markdown(f"**区間/対象:** {t['section']}")
                
                a_count = t.get('adults', meta.get('num_adults', 1))
                c_count = t.get('children', meta.get('num_children', 0))
                st.markdown(f"**利用人数:** 大人 {a_count}名 / 小人 {c_count}名")
                
                st.markdown(f"**金額:** {t.get('amount', '---')}")
                
                st.divider()
                st.warning("⚠️ **「使用確認」ボタンは係員の前で押してください。**")
                
                if t['status'] == "未使用":
                    if st.button("使用確認", type="primary", use_container_width=True):
                        for i, tk in enumerate(st.session_state.my_tickets):
                            if tk['id'] == t['id']:
                                st.session_state.my_tickets[i]['status'] = "使用済み"
                                break
                        st.session_state.active_ticket_id = None
                        st.rerun()
                else:
                    st.error("このチケットは既に使用済みです。")
                    
    # 【ビューA: チケット一覧画面】
    else:
        st.success(f"**{st.session_state.user_info.get('name', 'お客様')}** 様のチケット一覧です。利用するチケットの「使用する」ボタンを押してください。")
        
        for i, t in enumerate(st.session_state.my_tickets):
            with st.container(border=True):
                col_info, col_action = st.columns([3, 1])
                with col_info:
                    st.markdown(f"#### {t['title']}")
                    st.markdown(f"**{t['section']}**")
                    if "発 ➔" in t['time']:
                        st.caption(f"🕒 {t['time']} （他の便にも乗車・乗船できます。）")
                    else:
                        st.caption(f"🕒 {t['time']} （当日の営業時間内であればいつでも利用可能です。）")
                with col_action:
                    if t['status'] == "未使用":
                        if st.button("使用する", key=f"open_btn_{t['id']}", type="primary", use_container_width=True):
                            st.session_state.active_ticket_id = t['id']
                            st.rerun()
                    else:
                        st.markdown("<div style='text-align: center; color: gray; margin-top: 10px; padding: 10px; background-color: #f0f0f0; border-radius: 5px;'><b>使用済み</b></div>", unsafe_allow_html=True)
                        
        st.divider()
        if st.button("🔄 検索画面に戻る (データをリセット)"):
            st.session_state.current_page = 'search'
            st.session_state.confirm_plan = None
            st.session_state.plans_cache = None
            st.session_state.my_tickets = []
            st.session_state.active_ticket_id = None
            st.rerun()

# ------------------------------------------
# 画面4: 購入確認画面
# ------------------------------------------
elif st.session_state.current_page == 'confirm' and st.session_state.confirm_plan is not None:
    p = st.session_state.confirm_plan
    f = p['fares']
    meta = st.session_state.get('search_meta', {})
    
    st.markdown("## 🛍️ ご購入内容の確認")
    
    start_st = meta.get('start_station', '')
    goal_st = meta.get('goal_station', '')
    travel_dt_str = meta.get('travel_date_str', str(datetime.date.today()))
    start_tm_str = meta.get('start_time_str', '06:30')
    
    st.markdown(
        f"""
        <div style="background-color: #e6f4ea; border: 2px solid #34a853; border-radius: 10px; padding: 15px 20px; margin-bottom: 20px;">
            <h4 style="color: #137333; margin-top: 0;">🗺️ 選択されたプラン：{p['order']}</h4>
            <p style="margin-bottom: 0; color: #202124; font-size: 1.05em;">
                <b>📍 出発・到着:</b> {start_st} ➔ {goal_st} &nbsp;&nbsp;|&nbsp;&nbsp; <b>📅 出発日:</b> {travel_dt_str} &nbsp;&nbsp;|&nbsp;&nbsp; <b>⏰ 出発時刻:</b> {start_tm_str}発
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        st.metric("① 通常チケット積み上げ合計", f"¥{f['normal_total']:,}")
        st.caption(f"大人: ¥{f['normal_adult']:,} / 小人: ¥{f['normal_child']:,}")
    with col_c2:
        st.metric("② 手配原価合計", f"¥{f['cost_total']:,}")
        st.caption(f"大人: ¥{f['cost_adult']:,} / 小人: ¥{f['cost_child']:,}")
    with col_c3:
        st.metric("③ 🎉 販売価格合計", f"¥{f['sales_total']:,}")
        st.caption(f"大人: ¥{f['sales_adult']:,} / 小人: ¥{f['sales_child']:,}")
        
    st.info(f"⏱ **総所要時間**: {p['total_duration']//60}時間{p['total_duration']%60}分")
    
    st.markdown("---")
    st.markdown("#### 🎫 このパスに含まれる交通機関・施設・アクティビティ")
    st.markdown("このパスには、次の交通機関・施設・アクティビティの運賃が含まれています。")
    
    try:
        sanriku_fares_ref, other_fares_ref, facility_fares_ref = load_fare_data(FILE_PATH)
    except:
        sanriku_fares_ref, other_fares_ref, facility_fares_ref = {}, {}, {}

    n_adults = meta.get('num_adults', 1)
    n_children = meta.get('num_children', 0)

    line_details = {}
    spots_details = []

    for step in p['history']:
        if step['type'] == 'ride':
            line = step['line']
            f_stop = step['from_stop']
            t_stop = step['to_stop']
            
            if line == '宮古うみねこ丸':
                continue
                
            if line == '三陸鉄道':
                n_a = sanriku_fares_ref.get((f_stop, t_stop), 0)
                n_c = ((n_a + 19) // 20) * 10 if n_a > 0 else 0
                sum_n = n_a * n_adults + n_c * n_children
                parts = []
                if n_adults > 0:
                    parts.append(f"大人 ¥{n_a:,}×{n_adults}")
                if n_children > 0:
                    parts.append(f"小人 ¥{n_c:,}×{n_children}")
                desc = f"{f_stop} ➔ {t_stop} : 通常 ¥{sum_n:,} ({', '.join(parts)})"
                if line not in line_details:
                    line_details[line] = []
                line_details[line].append(desc)
            else:
                f_info = other_fares_ref.get((line, f_stop, t_stop), {})
                n_a = int(f_info.get('normal_adult', 0))
                n_c = int(f_info.get('normal_child', 0))
                capacity_str = f_info.get('capacity', '')
                total_p = n_adults + n_children
                
                if '1台' in capacity_str or '人' in capacity_str:
                    nums = re.findall(r'\d+', capacity_str)
                    cap = int(nums[0]) if nums else 4
                    vehicles = (total_p + cap - 1) // cap if total_p > 0 else 0
                    sum_n = n_a * vehicles
                    desc = f"{f_stop} ➔ {t_stop} : 通常 ¥{sum_n:,} (車両{vehicles}台分)"
                else:
                    sum_n = n_a * n_adults + n_c * n_children
                    parts = []
                    if n_adults > 0:
                        parts.append(f"大人 ¥{n_a:,}×{n_adults}")
                    if n_children > 0:
                        parts.append(f"小人 ¥{n_c:,}×{n_children}")
                    desc = f"{f_stop} ➔ {t_stop} : 通常 ¥{sum_n:,} ({', '.join(parts)})"
                    
                if line not in line_details:
                    line_details[line] = []
                line_details[line].append(desc)
                
        elif step['type'] == 'stay':
            act_name = step.get('activity')
            if act_name:
                f_info = facility_fares_ref.get(act_name, {})
                n_a = int(f_info.get('normal_adult', 0))
                n_c = int(f_info.get('normal_child', 0))
                
                if "サッパ船" in act_name and (n_adults + n_children) <= 2:
                    sum_n = 7600
                    share_a = n_a * n_adults
                    share_c = n_c * n_children
                    base_sum = share_a + share_c
                    if base_sum > 0:
                        alloc_a = int(round(7600 * (share_a / base_sum)))
                        alloc_c = 7600 - alloc_a
                    else:
                        alloc_a = 7600 if n_adults > 0 else 0
                        alloc_c = 7600 if n_children > 0 else 0
                    parts = []
                    if n_adults > 0:
                        parts.append(f"大人 ¥{alloc_a:,}")
                    if n_children > 0:
                        parts.append(f"小人 ¥{alloc_c:,}")
                    desc = f"特定 ¥7,600 ({', '.join(parts)})"
                else:
                    sum_n = n_a * n_adults + n_c * n_children
                    if sum_n > 0:
                        parts = []
                        if n_adults > 0:
                            parts.append(f"大人 ¥{n_a:,}×{n_adults}")
                        if n_children > 0:
                            parts.append(f"小人 ¥{n_c:,}×{n_children}")
                        desc = f"通常 ¥{sum_n:,} ({', '.join(parts)})"
                    else:
                        desc = "入場無料"
                        
                spots_details.append((act_name, desc))

    st.markdown("**🚆 含まれる交通機関（路線）:**")
    all_active_lines = set(line_details.keys())
    for step in p['history']:
        if step['type'] == 'ride':
            all_active_lines.add(step['line'])
            
    lines_html = "<ul>"
    for line in sorted(list(all_active_lines)):
        if line == '宮古うみねこ丸':
            lines_html += f"<li><b>{line}</b>: 施設・アクティビティに含まれます。</li>"
        else:
            lines_html += f"<li><b>{line}</b><ul>"
            for d in line_details.get(line, []):
                lines_html += f"<li>{d}</li>"
            lines_html += "</ul></li>"
    lines_html += "</ul>"
    st.markdown(lines_html, unsafe_allow_html=True)
        
    st.markdown("**🏛️ 含まれる施設・アクティビティ:**")
    spots_html = "<ul>"
    for spot_name, spot_desc in spots_details:
        spots_html += f"<li><b>{spot_name}</b>: {spot_desc}</li>"
    spots_html += "</ul>"
    st.markdown(spots_html, unsafe_allow_html=True)
        
    st.markdown("---")
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("⬅️ 検索結果に戻る", use_container_width=True):
            st.session_state.current_page = 'search'
            st.rerun()
    with col_btn2:
        if st.button("🎉 この内容で申し込む（購入手続きへ）", type="primary", use_container_width=True):
            st.session_state.current_page = 'auth'
            st.session_state.auth_step = 'input_email'
            st.rerun()

# ------------------------------------------
# 画面5: ルート検索 (デフォルト)
# ------------------------------------------
elif st.session_state.current_page == 'search':
    st.caption("行きたい場所を選んで、ルート検索からチケット購入まで")
    
    with st.expander("⚙️ **旅行条件・訪問地を設定する**", expanded=(st.session_state.plans_cache is None)):
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            travel_date = st.date_input("出発日", datetime.date.today())
        with col_d2:
            start_time = st.time_input("出発希望時刻", datetime.time(6, 30))
            
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            num_adults = st.number_input("👤 大人人数（中学生以上）", min_value=0, max_value=10, value=1, step=1)
        with col_p2:
            num_children = st.number_input("🧒 小人人数（小学生）", min_value=0, max_value=10, value=0, step=1)
            
        date_str = travel_date.strftime("%Y-%m-%d")
        start_time_str = start_time.strftime("%H:%M")

        try:
            xls_obj = pd.ExcelFile(FILE_PATH)
            if '出発駅・到着駅一覧' in xls_obj.sheet_names:
                df_st = pd.read_excel(FILE_PATH, sheet_name='出発駅・到着駅一覧')
                station_options = df_st[df_st.columns[0]].dropna().astype(str).tolist()
            else:
                station_options = ["宮古", "久慈", "釜石", "盛"]
        except Exception:
            station_options = ["宮古", "久慈", "釜石", "盛"]

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            start_station = st.selectbox("出発駅（起点）", station_options, index=0)
        with col_s2:
            goal_station = st.selectbox("到着駅（最終目的地）", station_options, index=0)

        st.markdown("##### 📍 訪問スポットを選択")
        
        try:
            _, _, facility_fares_raw = load_fare_data(FILE_PATH)
        except:
            facility_fares_raw = {}

        district_groups = {}
        for act_name, info in facility_fares_raw.items():
            dist = info['district']
            if dist not in district_groups:
                district_groups[dist] = []
            district_groups[dist].append(act_name)

        selected_spots_with_stay = []

        for dist, acts in district_groups.items():
            st.markdown(f"**📌 【 {dist} 地区 】**")
            for act_name in acts:
                info = facility_fares_raw[act_name]
                display_name = act_name

                is_fixed = info['is_fixed']
                fixed_duration = info['fixed_duration']
                description = info['description']
                
                if is_fixed:
                    c_chk, c_dummy = st.columns([3, 2])
                    with c_chk:
                        checked = st.checkbox(display_name, key=f"chk_{act_name}")
                    with c_dummy:
                        if checked:
                            if description and description != 'nan':
                                st.caption(f"💡 {description}")
                            st.caption(f"⏱ 標準所要時間: {fixed_duration}分")
                            selected_spots_with_stay.append((act_name, fixed_duration, act_name))
                else:
                    c_chk, c_num = st.columns([3, 2])
                    with c_chk:
                        checked = st.checkbox(display_name, key=f"chk_{act_name}")
                    with c_num:
                        if checked:
                            if description and description != 'nan':
                                st.caption(f"💡 {description}")
                            
                            stay = st.number_input("滞在(分)", min_value=15, max_value=240, value=fixed_duration, step=15, key=f"stay_{act_name}")
                            selected_spots_with_stay.append((act_name, stay, act_name))

        search_btn = st.button("🔍 最適ルートを検索する", type="primary", use_container_width=True)

    if search_btn:
        if num_adults + num_children == 0:
            st.warning("⚠️ 利用人数（大人または小人）を1名以上設定してください。")
        elif not selected_spots_with_stay:
            st.warning("⚠️ 訪問したい観光スポットを1つ以上選択してください。")
        else:
            with st.spinner("ダイヤと乗り継ぎ・運賃を最適化計算中..."):
                try:
                    legs, trans_map = load_data(FILE_PATH, date_str)
                    sanriku_fares, other_fares, facility_fares = load_fare_data(FILE_PATH)
                    
                    plans = plan_tour(legs, trans_map, start_station, start_time_str, goal_station, selected_spots_with_stay, sanriku_fares, other_fares, facility_fares, num_adults, num_children)
                    st.session_state.plans_cache = plans
                    st.session_state.search_meta = {
                        'start_station': start_station,
                        'goal_station': goal_station,
                        'travel_date_str': date_str,
                        'start_time_str': start_time_str,
                        'num_adults': num_adults,
                        'num_children': num_children
                    }
                    st.rerun()
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

    if st.session_state.plans_cache is not None:
        plans = st.session_state.plans_cache
        meta = st.session_state.get('search_meta', {})
        
        if not plans:
            st.error("❌ 条件に一致するルートが見つかりませんでした。出発時刻や滞在時間を調整してください。")
        else:
            st.success(f"🎉 **{len(plans)} 件**のルートが見つかりました！（大人: {meta.get('num_adults', 1)}名, 小人: {meta.get('num_children', 0)}名）")
            
            for idx, p in enumerate(plans, 1):
                f = p['fares']
                
                with st.container(border=True):
                    st.markdown(f"### ⭐ プラン {idx}：{p['order']}")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("① 個別積上合計", f"¥{f['normal_total']:,}")
                        st.caption(f"内訳: 大人 ¥{f['normal_adult']:,} / 小人 ¥{f['normal_child']:,}")
                    with col2:
                        st.metric("② 手配原価合計", f"¥{f['cost_total']:,}")
                        st.caption(f"内訳: 大人 ¥{f['cost_adult']:,} / 小人 ¥{f['cost_child']:,}")
                    with col3:
                        st.metric("③ 🎉 販売価格合計", f"¥{f['sales_total']:,}")
                        st.caption(f"内訳: 大人 ¥{f['sales_adult']:,} / 小人 ¥{f['sales_child']:,}")
                        
                    st.info(f"⏱ **総所要時間**: {p['total_duration']//60}時間{p['total_duration']%60}分 ｜ **区間**: {meta.get('start_station', '')} ({meta.get('start_time_str', '')}発) ➔ {meta.get('goal_station', '')} (**{min_to_str(p['final_arr_time'])}着**)")
                    
                    with st.expander("💴 運賃・アクティビティ計算の内訳（大人・小人別）を見る"):
                        for b in p['breakdown']:
                            st.write(b)
                            
                    diff = f['sales_total'] - f['normal_total']
                    if diff >= 200:
                        st.markdown(
                            """
                            <div style="background-color: #ffebee; border: 3px solid #e53935; border-radius: 8px; padding: 15px; margin: 20px 0;">
                                <div style="color: #c62828; font-size: 1.15em; font-weight: bold; margin-bottom: 8px;">
                                    ⚠️ ご確認：このプランは、個別にチケットを購入された方が安いですがよろしいですか。
                                </div>
                                <div style="color: #333333; font-size: 1em; line-height: 1.5;">
                                    このプランのメリットは、その都度、決済する必要がないことです。内容をご確認の上、よろしければ「購入確認へ進む」を押してください。
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                    elif abs(f['normal_total'] - f['sales_total']) < 200:
                        st.markdown(
                            """
                            <div style="background-color: #fff8e1; border: 3px solid #ffb300; border-radius: 8px; padding: 15px; margin: 20px 0;">
                                <div style="color: #ff8f00; font-size: 1.15em; font-weight: bold;">
                                    ⚠️ ご確認：このプランは、個別にチケットを購入された場合との差は200円未満ですが、よろしいですか。
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                            
                    if st.button(f"🛒 このプランを選択して購入確認へ進む (プラン {idx})", key=f"btn_confirm_{idx}", type="primary"):
                        st.session_state.confirm_plan = p
                        st.session_state.current_page = 'confirm'
                        st.rerun()
                    
                    st.markdown("##### 📍 ルート詳細")
                    for step in p['history']:
                        if step['type'] == 'ride':
                            st.markdown(f"🚆 **[{step['line']}]** `{step['from_stop']}` (**{min_to_str(step['dep_time'])}発**) ➔ `{step['to_stop']}` (**{min_to_str(step['arr_time'])}着**)")
                        elif step['type'] == 'transfer':
                            st.caption(f" 🚶 **徒歩・乗換 {step['duration']}分**: {step['from_stop']} ➔ {step['to_stop']}")
                        elif step['type'] == 'stay':
                            if step['activity']:
                                if step.get('timetable_ride'):
                                    st.markdown(
                                        f"""
                                        <div style="background-color: #f0f8ff; border: 2px solid #4682b4; border-radius: 8px; padding: 10px 15px; margin: 10px 0;">
                                            ⭐ <b>【施設・体験】 {step['activity']}</b><br>
                                            <span style="color: #333333; font-size: 0.9em;">🕒 港到着 {min_to_str(step['arr_time'])} ➔ 乗船 {min_to_str(step['ride_dep'])} 〜 {min_to_str(step['ride_arr'])}</span>
                                        </div>
                                        """,
                                        unsafe_allow_html=True
                                    )
                                else:
                                    st.markdown(
                                        f"""
                                        <div style="background-color: #f0f8ff; border: 2px solid #4682b4; border-radius: 8px; padding: 10px 15px; margin: 10px 0;">
                                            ⭐ <b>【施設・体験】 {step['activity']}</b><br>
                                            <span style="color: #333333; font-size: 0.9em;">🕒 滞在時間: {min_to_str(step['arr_time'])} 〜 {min_to_str(step['dep_time'])} （{step['stay_min']}分間）</span>
                                        </div>
                                        """,
                                        unsafe_allow_html=True
                                    )
                            else:
                                st.markdown(
                                    f"""
                                    <div style="background-color: #f5fffa; border: 2px solid #2e8b57; border-radius: 8px; padding: 10px 15px; margin: 10px 0;">
                                        🌿 <b>【観光・滞在】 {step['spot']}</b><br>
                                        <span style="color: #333333; font-size: 0.9em;">🕒 滞在時間: {min_to_str(step['arr_time'])} 〜 {min_to_str(step['dep_time'])} （{step['stay_min']}分間）</span>
                                    </div>
                                    """,
                                    unsafe_allow_html=True
                                )
            
            st.divider()
            if st.button("🔄 条件をリセットして再検索する", use_container_width=True):
                st.session_state.plans_cache = None
                st.session_state.search_meta = {}
                st.rerun()