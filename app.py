import streamlit as st
import pandas as pd
import datetime
import itertools
import heapq
import re

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
    has_ryusendo = any(step.get('type'] == 'stay' and step.get('activity') == '龍泉洞' for step in history)
    
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
                        # 1名または2名利用時は最低7600円
                        sum_n = 7600
                        sum_c = 7600
                        
                        # 大人・小人への按分または計上
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
                        breakdown.append(f"🎫 [体験・入場] {act_name} (1〜2名最低料金適用) : 通常 ¥7,600 ({', '.join(parts)}) / 原価 ¥7,600")
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
            
            routes = find_routes_point_to_point(legs, trans_map, cur_stop, cur_time, target_stop)
            if not routes:
                is_possible = False
                break
            
            arr_time, route_hist = routes[0]
            full_history.extend(route_hist)
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
# 2. UI 表示
# ==========================================
st.title("🚃 三陸海岸・じぶんの旅パス")
st.caption("行きたい場所を選んで、ルート検索からチケット購入まで")

with st.expander("⚙️ **旅行条件・訪問地を設定する**", expanded=True):
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
                
                if not plans:
                    st.error(f"❌ {start_station}発 ➔ {goal_station}着 で当日中に移動できるルートが見つかりませんでした。出発時刻や滞在時間を調整してください。")
                else:
                    st.success(f"🎉 **{len(plans)} 件**のルートが見つかりました！（大人: {num_adults}名, 小人: {num_children}名）")
                    
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
                                
                            st.info(f"⏱ **総所要時間**: {p['total_duration']//60}時間{p['total_duration%60']}分 ｜ **区間**: {start_station} ({start_time_str}発) ➔ {goal_station} (**{min_to_str(p['final_arr_time'])}着**)")
                            
                            with st.expander("💴 運賃・アクティビティ計算の内訳（大人・小人別）を見る"):
                                for b in p['breakdown']:
                                    st.write(b)
                            
                            st.markdown("##### 📍 ルート詳細")
                            for step in p['history']:
                                if step['type'] == 'ride':
                                    st.markdown(f"🚆 **[{step['line']}]** `{step['from_stop']}` (**{min_to_str(step['dep_time'])}発**) ➔ `{step['to_stop']}` (**{min_to_str(step['arr_time'])}着**)")
                                elif step['type'] == 'transfer':
                                    st.caption(f" 🚶 **徒歩・乗換 {step['duration']}分**: {step['from_stop']} ➔ {step['to_stop']}")
                                elif step['type'] == 'stay':
                                    if step['activity']:
                                        st.success(f"★ **【施設・体験】 {step['activity']}** （**{min_to_str(step['arr_time'])} 〜 {min_to_str(step['dep_time'])}** / {step['stay_min']}分間）")
                                    else:
                                        st.success(f"★ **【観光・滞在】 {step['spot']}** （**{min_to_str(step['arr_time'])} 〜 {min_to_str(step['dep_time'])}** / {step['stay_min']}分間）")
                    
                    st.divider()
                    if st.button("🔄 次の検索（条件を再設定する）", use_container_width=True):
                        st.rerun()
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")