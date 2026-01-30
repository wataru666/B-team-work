import streamlit as st
import pandas as pd
# burantiテスト
# セッションステートでデータを管理
if 'grade_data' not in st.session_state:
    st.session_state.grade_data = pd.DataFrame(columns=['科目', '成績'])

if 'teaching_style' not in st.session_state:
    st.session_state.teaching_style = pd.DataFrame(columns=['科目', '授業スタイル'])

st.title("授業効率化システム")

# ページ選択
page = st.sidebar.selectbox("ページを選択", ["メインメニュー", "編集画面", "比較・閲覧画面"])

# メインメニュー
if page == "メインメニュー":
    st.write("成績データと授業スタイルデータを管理します。")

    st.write("サイドバーからページを選択してください。")

# 編集画面
elif page == "編集画面":
    st.header("編集画面")
    
    # タブでセイセキ管理とスタイル管理を分ける
    tab1, tab2 = st.tabs(["成績管理", "スタイル管理"])
    
    with tab1:
        st.subheader("成績データの確認")

        # 成績管理.xlsxの全シート名を取得
        try:
            excel_file = pd.ExcelFile('成績管理.xlsx')
            sheet_names = excel_file.sheet_names
            
            st.success("成績データを読み込みました。")
        except Exception as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
            sheet_names = []
        
        # 絞り込み機能：シート名（科目）・西暦・期間
        st.subheader("実施日の絞り込み")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            selected_sheet = st.selectbox("科目", sheet_names, key="grade_subject")
        with col2:
            year = st.number_input("西暦", min_value=2000, max_value=2100, value=2024)
        with col3:
            period = st.selectbox("期間", ["前期", "後期"], key="grade_period")
        
        search_keyword = f"{year}年{period}"
        
        # 選択したシートを読み込み
        if selected_sheet:
            try:
                excel_data = pd.read_excel('成績管理.xlsx', sheet_name=selected_sheet)
                
                if not excel_data.empty and len(excel_data.columns) > 1:
                    # B列（2番目の列）から絞り込み
                    b_column = excel_data.iloc[:, 1]
                    filtered_indices = b_column[b_column.astype(str).str.contains(search_keyword, na=False)].index
                    
                    if len(filtered_indices) > 0:
                        # 絞り込み結果の行を抽出（B, C, D, E列 - 科目は不要）
                        filtered_result = excel_data.iloc[filtered_indices, 1:5].copy()
                        
                        # カラム名を設定
                        filtered_result.columns = ['実施日', '学籍番号', '点数', 'かかった時間']
                        
                        # 点数に「点」、時間に「分」の単位を追加
                        filtered_result['点数'] = filtered_result['点数'].astype(str) + '点'
                        filtered_result['かかった時間'] = filtered_result['かかった時間'].astype(str) + '分'
                        
                        st.success(f"'{selected_sheet}' - '{search_keyword}' に該当する実施日:")
                        st.dataframe(filtered_result)
                    else:
                        st.info(f"'{selected_sheet}' - '{search_keyword}' に該当するデータはありません。")
                
                st.divider()
                st.subheader(f"シート '{selected_sheet}' の全データ")
                st.dataframe(excel_data)
            except Exception as e:
                st.error(f"シートの読み込みに失敗しました: {e}")
    
    with tab2:
        st.subheader("スタイルデータの確認・更新")
        
        # スタイル管理.xlsxの全シート名を取得
        try:
            style_file = pd.ExcelFile('スタイル管理.xlsx')
            style_sheet_names = style_file.sheet_names
            
            st.success("スタイルデータを読み込みました。")
        except Exception as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
            style_sheet_names = []
        
        # シート選択
        if style_sheet_names:
            col1, col2 = st.columns(2)
            
            with col1:
                selected_style_sheet = st.selectbox("科目", style_sheet_names, key="style_subject")
            
            # スタイル名を取得（E2、F2以降の2行目の列から）
           # --- 全シートからスタイル名を抽出する ---
            style_names = []

            try:
                style_file = pd.ExcelFile('スタイル管理.xlsx')
                style_sheet_names = style_file.sheet_names

                for sheet in style_sheet_names:
                    df = pd.read_excel('スタイル管理.xlsx', sheet_name=sheet, header=None)

                    # 2行目（index=1）が存在するか確認
                    if len(df) > 1 and len(df.columns) > 4:
                        # E列(4)〜右端まで
                        for col_idx in range(4, len(df.columns)):
                            value = df.iloc[1, col_idx]
                            if pd.notna(value) and str(value).strip() != "":
                                style_names.append(str(value).strip())

                # 重複削除
                style_names = list(dict.fromkeys(style_names))

                st.write(f"全シートから抽出されたスタイル名: {style_names}")

            except Exception as e:
                st.error(f"スタイル名の取得に失敗しました: {e}")
            
            with col2:
                if style_names:
                    selected_style_name = st.selectbox("スタイル名", [""] + style_names, key="style_name")
                else:
                    st.info("スタイル名が見つかりません。スタイル管理ファイルを確認してください。")
                    selected_style_name = ""
            
            try:
                style_data = pd.read_excel('スタイル管理.xlsx', sheet_name=selected_style_sheet)
                
                # スタイル名が選択されている場合のみ処理
                if selected_style_name:
                    st.subheader(f"スタイル '{selected_style_name}' のデータ")
                else:
                    st.subheader(f"シート '{selected_style_sheet}' のスタイルデータ")
                
                # 計算・更新ボタン
                if st.button("平均点数と平均時間を計算・更新"):
                    # 成績管理.xlsxから対応するシートを取得
                    try:
                        grade_data = pd.read_excel('成績管理.xlsx', sheet_name=selected_style_sheet)
                        
                        # B2以降の実施日を取得（インデックス1以降）
                        if len(style_data) > 1 and len(style_data.columns) >= 4:
                            for idx in range(1, len(style_data)):
                                implementation_date = str(style_data.iloc[idx, 1]).strip()  # B列（スタイル管理）
                                
                                if pd.isna(style_data.iloc[idx, 1]) or implementation_date == 'nan':
                                    continue
                                
                                # 成績管理から実施日が完全に一致するデータを抽出（B2以降を対象）
                                matching_rows = []
                                for grade_idx in range(1, len(grade_data)):
                                    grade_date_str = str(grade_data.iloc[grade_idx, 1]).strip()
                                    if grade_date_str == implementation_date:
                                        matching_rows.append(grade_idx)
                                
                                if len(matching_rows) > 0:
                                    # D列（点数）とE列（かかった時間）の平均を計算
                                    scores = grade_data.iloc[matching_rows, 3].astype(float).values
                                    times = grade_data.iloc[matching_rows, 4].astype(float).values
                                    
                                    avg_score = scores.mean()
                                    avg_time = times.mean()
                                    
                                    # スタイル管理のC列、D列に入力
                                    style_data.iloc[idx, 2] = round(avg_score, 2)
                                    style_data.iloc[idx, 3] = round(avg_time, 2)
                            
                            # 更新したデータをスタイル管理.xlsxに保存
                            with pd.ExcelWriter('スタイル管理.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                style_data.to_excel(writer, sheet_name=selected_style_sheet, index=False)
                            
                            st.success(f"'{selected_style_sheet}' の平均点数と平均時間を計算・更新しました。")
                        else:
                            st.error("スタイル管理のデータ形式が正しくありません。")
                    except Exception as e:
                        st.error(f"計算・更新に失敗しました: {e}")
                
                st.subheader(f"シート '{selected_style_sheet}' のスタイルデータ")
                
                # 科目列を除いた表示用データ（B, C, D列のみ）
                display_data = style_data.iloc[:, 1:4].copy()
                display_data.columns = ['実施日', '平均点数', '平均時間']
                st.dataframe(display_data)
            except Exception as e:
                st.error(f"シートの読み込みに失敗しました: {e}")
    
    # 従来の手動入力フォーム（オプション）
    st.divider()
    st.subheader("手動入力（オプション）")
    with st.form("grade_form"):
        subject = st.text_input("科目")
        grade = st.number_input("成績", min_value=0, max_value=100)
        submitted = st.form_submit_button("追加")
        if submitted:
            new_row = pd.DataFrame({'科目': [subject], '成績': [grade]})
            st.session_state.grade_data = pd.concat([st.session_state.grade_data, new_row], ignore_index=True)
            st.success("成績データを追加しました。")

    st.subheader("授業スタイルデータの編集")
    with st.form("style_form"):
        subject_style = st.text_input("科目 (スタイル)")
        style = st.text_area("授業スタイル")
        submitted_style = st.form_submit_button("追加")
        if submitted_style:
            new_row_style = pd.DataFrame({'科目': [subject_style], '授業スタイル': [style]})
            st.session_state.teaching_style = pd.concat([st.session_state.teaching_style, new_row_style], ignore_index=True)
            st.success("授業スタイルデータを追加しました。")

# 比較・閲覧画面
elif page == "比較・閲覧画面":
    st.header("比較・閲覧画面")
    st.subheader("成績データ")
    st.dataframe(st.session_state.grade_data)

    st.subheader("授業スタイルデータ")
    st.dataframe(st.session_state.teaching_style)

    # 比較機能（例: 成績の平均）
    if not st.session_state.grade_data.empty:
        avg_grade = st.session_state.grade_data['成績'].mean()
        st.write(f"成績の平均: {avg_grade:.2f}")

    # --- 追加: 成績管理.xlsxから年度・期間を抽出して比較グラフを表示 ---
    try:
        excel_file = pd.ExcelFile('成績管理.xlsx')
        sheet_names = excel_file.sheet_names

        # B列に含まれる「YYYY年(前期|後期)」を抽出
        import re

        period_set = set()
        for sheet in sheet_names:
            try:
                df = pd.read_excel('成績管理.xlsx', sheet_name=sheet, dtype=str)
                if df.shape[1] > 1:
                    bcol = df.iloc[:, 1].astype(str).fillna('')
                    for s in bcol:
                        m = re.search(r"(\d{4})年(前期|後期)", s)
                        if m:
                            period_set.add(f"{m.group(1)}年{m.group(2)}")
            except Exception:
                continue

        period_options = sorted(list(period_set))

        st.subheader("年度・期間で比較 (複数選択可)")
        selected_periods = st.multiselect("年度・期間を選択", period_options, key="compare_periods")

        if selected_periods:
            results = []
            for token in selected_periods:
                scores_acc = []
                times_acc = []
                for sheet in sheet_names:
                    try:
                        df = pd.read_excel('成績管理.xlsx', sheet_name=sheet, dtype=str)
                        if df.shape[1] > 4:
                            bcol = df.iloc[:, 1].astype(str).fillna('')
                            mask = bcol.str.contains(token, na=False)
                            if mask.any():
                                # D列(インデックス3)=点数, E列(インデックス4)=かかった時間
                                score_vals = pd.to_numeric(df.loc[mask, df.columns[3]], errors='coerce')
                                time_vals = pd.to_numeric(df.loc[mask, df.columns[4]], errors='coerce')
                                scores_acc.extend(score_vals.dropna().tolist())
                                times_acc.extend(time_vals.dropna().tolist())
                    except Exception:
                        continue

                avg_score = float('nan') if len(scores_acc) == 0 else sum(scores_acc) / len(scores_acc)
                avg_time = float('nan') if len(times_acc) == 0 else sum(times_acc) / len(times_acc)
                results.append({'period': token, 'avg_score': avg_score, 'avg_time': avg_time})

            if results:
                df_res = pd.DataFrame(results)

                # 年度順にソート
                def _sort_key(x):
                    m = re.match(r"(\d{4})年(前期|後期)", x)
                    if not m:
                        return (0, 0)
                    year = int(m.group(1))
                    term = 1 if m.group(2) == '前期' else 2
                    return (year, term)

                df_res['sort_key'] = df_res['period'].map(_sort_key)
                df_res = df_res.sort_values('sort_key').set_index('period')
                df_res = df_res[['avg_score', 'avg_time']]

                st.subheader("平均点数と平均時間の推移")
                st.line_chart(df_res)
            else:
                st.info("選択した期間に該当するデータが見つかりませんでした。")
    except FileNotFoundError:
        st.info("成績管理.xlsx が見つかりません。編集画面から読み込んでください。")
    except Exception as e:
        st.error(f"比較データの作成に失敗しました: {e}")

