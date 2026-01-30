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
        
        if style_sheet_names:
            # ============================================
            # スタイル管理セクション
            # ============================================
            st.subheader("スタイル管理")
            
            # 1. 科目・西暦・期間の選択
            col1, col2, col3 = st.columns(3)
            with col1:
                selected_style_sheet = st.selectbox("科目", style_sheet_names, key="style_subject")
            with col2:
                style_year = st.number_input("西暦", min_value=2000, max_value=2100, value=2024, key="style_year")
            with col3:
                style_period = st.selectbox("期間", ["前期", "後期"], key="style_period")
            
            # 2. スタイル名を取得（E2、F2以降の2行目の列から）
            style_names = []
            try:
                style_file = pd.ExcelFile('スタイル管理.xlsx')
                style_sheet_names_all = style_file.sheet_names

                for sheet in style_sheet_names_all:
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

            except Exception as e:
                st.error(f"スタイル名の取得に失敗しました: {e}")
            
            # 3. スタイル名ごとに適応/適応しないを選択できるUIを作成
            st.write("**各スタイルの適用状況を選択:**")
            
            style_selections = {}
            
            # 既存スタイル名の選択
            if style_names:
                for style_name in style_names:
                    col_name, col_radio = st.columns([3, 1])
                    with col_name:
                        st.write(f"**{style_name}**")
                    with col_radio:
                        choice = st.radio("", ["適応", "適応しない"], key=f"style_choice_{style_name}", horizontal=True, label_visibility="collapsed")
                        style_selections[style_name] = "〇" if choice == "適応" else "✕"
            
            # 4. 新規スタイル名の入力欄と適応/適応しない選択
            st.write("**新規スタイル:**")
            col_new_name, col_new_radio = st.columns([3, 1])
            with col_new_name:
                new_style_name = st.text_input("新規スタイル名を入力", key="new_style_name_input")
            with col_new_radio:
                if new_style_name.strip():
                    new_style_choice = st.radio("", ["適応", "適応しない"], key="new_style_choice", horizontal=True, label_visibility="collapsed")
                    style_selections[new_style_name.strip()] = "〇" if new_style_choice == "適応" else "✕"
            
            # 5. 決定ボタン
            if st.button("スタイルを一括適用", key="apply_style_button"):
                try:
                    if not style_selections:
                        st.error("適用するスタイルを選択してください。")
                        st.stop()
                    
                    # スタイル管理.xlsxを読み込み
                    style_data = pd.read_excel('スタイル管理.xlsx', sheet_name=selected_style_sheet, header=None)
                    
                    search_keyword = f"{style_year}年{style_period}"
                    
                    # 選択した時期が存在するかチェック
                    period_exists = False
                    if len(style_data) > 2:
                        for row_idx in range(2, len(style_data)):
                            b_value = str(style_data.iloc[row_idx, 1]).strip() if pd.notna(style_data.iloc[row_idx, 1]) else ""
                            if search_keyword in b_value:
                                period_exists = True
                                break
                    
                    if not period_exists:
                        st.error(f"選択した科目 '{selected_style_sheet}' に '{search_keyword}' のデータが見つかりません。")
                    else:
                        # 各スタイルを処理
                        for style_name, marker in style_selections.items():
                            # 既存スタイル名から列インデックスを探す
                            style_col_idx = None
                            
                            if style_name in style_names:
                                # 既存スタイルの場合
                                if len(style_data) > 1 and len(style_data.columns) > 4:
                                    for col_idx in range(4, len(style_data.columns)):
                                        value = style_data.iloc[1, col_idx]
                                        if pd.notna(value) and str(value).strip() == style_name:
                                            style_col_idx = col_idx
                                            break
                            else:
                                # 新規スタイルの場合は、2行目に列を追加
                                next_col_idx = len(style_data.columns)
                                # 新しい列をデータフレームに追加
                                for row_idx in range(len(style_data)):
                                    if row_idx == 1:
                                        # 2行目（index=1）にスタイル名を設定
                                        style_data.loc[row_idx, next_col_idx] = style_name
                                    else:
                                        # その他の行は初期化
                                        style_data.loc[row_idx, next_col_idx] = None
                                style_col_idx = next_col_idx
                                
                                # 追加した列にデータを書き込む前に新しい列数を確認
                                # すべてのデータ行に✕を初期化
                                for row_idx in range(2, len(style_data)):
                                    style_data.iloc[row_idx, style_col_idx] = "✕"
                            
                            if style_col_idx is not None:
                                # B列に search_keyword を含む行を探して、該当する列にマーカーを書き込み
                                if len(style_data) > 1:
                                    for row_idx in range(2, len(style_data)):  # 3行目以降（実データ部分）
                                        b_value = str(style_data.iloc[row_idx, 1]).strip() if pd.notna(style_data.iloc[row_idx, 1]) else ""
                                        
                                        if search_keyword in b_value:
                                            # 該当行のスタイル列に値を書き込み
                                            style_data.iloc[row_idx, style_col_idx] = marker
                        
                        # 更新したデータをExcelに保存
                        with pd.ExcelWriter('スタイル管理.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            style_data.to_excel(writer, sheet_name=selected_style_sheet, index=False, header=False)
                        
                        st.success("選択したスタイルを一括適用しました。")
                
                except Exception as e:
                    st.error(f"スタイルの適用に失敗しました: {e}")
            
            # ============================================
            # スタイル管理表の確認・表示
            # ============================================
            st.divider()
            st.subheader("スタイルデータの確認・更新")
            
            selected_style_sheet_view = st.selectbox("科目（表示用）", style_sheet_names, key="style_subject_view")
            
            try:
                style_data = pd.read_excel('スタイル管理.xlsx', sheet_name=selected_style_sheet_view)
                
                st.subheader(f"シート '{selected_style_sheet_view}' のスタイルデータ")
                
                # 計算・更新ボタン
                if st.button("平均点数と平均時間を計算・更新"):
                    # 成績管理.xlsxから対応するシートを取得
                    try:
                        grade_data = pd.read_excel('成績管理.xlsx', sheet_name=selected_style_sheet_view)
                        
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
                                style_data.to_excel(writer, sheet_name=selected_style_sheet_view, index=False)
                            
                            st.success(f"'{selected_style_sheet_view}' の平均点数と平均時間を計算・更新しました。")
                        else:
                            st.error("スタイル管理のデータ形式が正しくありません。")
                    except Exception as e:
                        st.error(f"計算・更新に失敗しました: {e}")
                
                # 科目列を除いた表示用データ（B, C, D列のみ）
                # 2行目以降（ヘッダー行を除いたデータ行）のみを取得
                display_data = style_data.iloc[2:, 1:4].copy()
                display_data.columns = ['実施日', '平均点数', '平均時間']
                display_data = display_data.reset_index(drop=True)
                
                st.dataframe(display_data)
            except Exception as e:
                st.error(f"シートの読み込みに失敗しました: {e}")

# 比較・閲覧画面
elif page == "比較・閲覧画面":
    st.header("比較・閲覧画面")
    
    st.subheader("スタイル管理データの比較")
    
    try:
        style_file = pd.ExcelFile('スタイル管理.xlsx')
        style_sheet_names = style_file.sheet_names
        
        # 1. 科目を選択
        selected_subject = st.selectbox("科目を選択", style_sheet_names, key="compare_subject")
        
        # 2. 2つの時期を入力
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**時期1**")
            year1 = st.number_input("西暦 1", min_value=2000, max_value=2100, value=2024, key="compare_year1")
            period1 = st.selectbox("期間 1", ["前期", "後期"], key="compare_period1")
        
        with col2:
            st.write("**時期2**")
            year2 = st.number_input("西暦 2", min_value=2000, max_value=2100, value=2024, key="compare_year2")
            period2 = st.selectbox("期間 2", ["前期", "後期"], key="compare_period2")
        
        # 比較ボタン
        if st.button("比較"):
            try:
                # スタイル管理.xlsxから選択されたシートを読み込み
                style_data = pd.read_excel('スタイル管理.xlsx', sheet_name=selected_subject, header=None)
                
                search_keyword1 = f"{year1}年{period1}"
                search_keyword2 = f"{year2}年{period2}"
                
                # 時期1のデータを抽出
                avg_score1 = None
                avg_time1 = None
                found1 = False
                
                for row_idx in range(2, len(style_data)):
                    date_str = str(style_data.iloc[row_idx, 1]).strip() if pd.notna(style_data.iloc[row_idx, 1]) else ""
                    if search_keyword1 in date_str:
                        found1 = True
                        score_val = style_data.iloc[row_idx, 2]
                        time_val = style_data.iloc[row_idx, 3]
                        avg_score1 = float(score_val) if pd.notna(score_val) else None
                        avg_time1 = float(time_val) if pd.notna(time_val) else None
                        break
                
                # 時期2のデータを抽出
                avg_score2 = None
                avg_time2 = None
                found2 = False
                
                for row_idx in range(2, len(style_data)):
                    date_str = str(style_data.iloc[row_idx, 1]).strip() if pd.notna(style_data.iloc[row_idx, 1]) else ""
                    if search_keyword2 in date_str:
                        found2 = True
                        score_val = style_data.iloc[row_idx, 2]
                        time_val = style_data.iloc[row_idx, 3]
                        avg_score2 = float(score_val) if pd.notna(score_val) else None
                        avg_time2 = float(time_val) if pd.notna(time_val) else None
                        break
                
                # 結果を表示
                if not found1:
                    st.error(f"時期1: '{selected_subject}' の '{search_keyword1}' のデータが見つかりません。")
                elif not found2:
                    st.error(f"時期2: '{selected_subject}' の '{search_keyword2}' のデータが見つかりません。")
                else:
                    st.success("データ取得成功")
                    
                    # 結果テーブルを作成
                    result_data = {
                        "項目": ["平均点数", "平均時間"],
                        f"{search_keyword1}": [avg_score1, avg_time1],
                        f"{search_keyword2}": [avg_score2, avg_time2],
                    }
                    
                    # 変化量を計算（%）
                    if avg_score1 is not None and avg_score2 is not None:
                        score_change = ((avg_score2 - avg_score1) / avg_score1 * 100) if avg_score1 != 0 else 0
                    else:
                        score_change = None
                    
                    if avg_time1 is not None and avg_time2 is not None:
                        time_change = ((avg_time2 - avg_time1) / avg_time1 * 100) if avg_time1 != 0 else 0
                    else:
                        time_change = None
                    
                    result_data["変化量(%)"] = [f"{score_change:.2f}%" if score_change is not None else "N/A",
                                               f"{time_change:.2f}%" if time_change is not None else "N/A"]
                    
                    result_df = pd.DataFrame(result_data)
                    
                    st.subheader("比較結果")
                    st.dataframe(result_df, use_container_width=True)
                    
                    # 解釈を表示
                    st.subheader("変化量の解釈")
                    if score_change is not None:
                        if score_change > 0:
                            st.write(f"✅ **平均点数**: {score_change:.2f}% 上昇しました")
                        elif score_change < 0:
                            st.write(f"❌ **平均点数**: {abs(score_change):.2f}% 低下しました")
                        else:
                            st.write(f"➡️ **平均点数**: 変わりません")
                    
                    if time_change is not None:
                        if time_change > 0:
                            st.write(f"⏱️ **平均時間**: {time_change:.2f}% 増加しました（効率が低下）")
                        elif time_change < 0:
                            st.write(f"⏱️ **平均時間**: {abs(time_change):.2f}% 減少しました（効率が向上）")
                        else:
                            st.write(f"➡️ **平均時間**: 変わりません")
                            
            except Exception as e:
                st.error(f"比較に失敗しました: {e}")
        
        # グラフは常に表示（比較ボタンを押さなくても表示）
        st.divider()
        st.subheader("すべての時期の推移グラフ")
        
        try:
            # スタイル管理.xlsxから選択されたシートを読み込み
            style_data = pd.read_excel('スタイル管理.xlsx', sheet_name=selected_subject, header=None)
            
            # 選択された科目のすべてのデータを集計
            all_periods_data = {}
            
            for row_idx in range(2, len(style_data)):
                date_str = str(style_data.iloc[row_idx, 1]).strip() if pd.notna(style_data.iloc[row_idx, 1]) else ""
                
                # 時期情報を抽出（YYYY年前期/後期形式）
                if "年" in date_str:
                    period_key = date_str.split("(")[0] if "(" in date_str else date_str
                    
                    # 既に同じ時期があればスキップ（最初の出現のみ使用）
                    if period_key not in all_periods_data:
                        score_val = style_data.iloc[row_idx, 2]
                        time_val = style_data.iloc[row_idx, 3]
                        
                        all_periods_data[period_key] = {
                            "平均点数": float(score_val) if pd.notna(score_val) else None,
                            "平均時間": float(time_val) if pd.notna(time_val) else None
                        }
            
            if all_periods_data:
                # DataFrameに変換
                graph_df = pd.DataFrame(all_periods_data).T
                graph_df = graph_df.dropna(how='all')
                
                # 時期でソート（年度順）
                def sort_by_period(x):
                    import re
                    m = re.match(r"(\d{4})年(前期|後期)", x)
                    if not m:
                        return (0, 0)
                    year = int(m.group(1))
                    term = 1 if m.group(2) == '前期' else 2
                    return (year, term)
                
                graph_df = graph_df.sort_index(key=lambda x: x.map(sort_by_period))
                
                # グラフを2つに分ける
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**📊 平均点数 (点)**")
                    if "平均点数" in graph_df.columns:
                        score_chart = graph_df[["平均点数"]].copy()
                        st.line_chart(score_chart, color=["#1f77b4"])
                        
                        # テーブルで数値も表示
                        score_table = pd.DataFrame({
                            "時期": score_chart.index,
                            "平均点数 (点)": score_chart["平均点数"].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")
                        })
                        st.dataframe(score_table, use_container_width=True, hide_index=True)
                
                with col2:
                    st.write("**⏱️ 平均時間 (分)**")
                    if "平均時間" in graph_df.columns:
                        time_chart = graph_df[["平均時間"]].copy()
                        st.line_chart(time_chart, color=["#ff7f0e"])
                        
                        # テーブルで数値も表示
                        time_table = pd.DataFrame({
                            "時期": time_chart.index,
                            "平均時間 (分)": time_chart["平均時間"].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")
                        })
                        st.dataframe(time_table, use_container_width=True, hide_index=True)
            else:
                st.info("データが見つかりません。")
        except Exception as e:
            st.error(f"グラフ表示に失敗しました: {e}")
        
        # ============================================
        # スタイル毎の比較セクション
        # ============================================
        st.divider()
        st.subheader("スタイル毎の結果比較")
        
        try:
            # スタイル比較用に別途科目を選択
            selected_subject_for_style = st.selectbox("科目を選択（スタイル比較用）", style_sheet_names, key="compare_subject_style")
            
            # 選択された科目のデータを読み込み
            style_data_compare = pd.read_excel('スタイル管理.xlsx', sheet_name=selected_subject_for_style, header=None)
            
            # スタイル名を取得（E列以降）
            available_styles = []
            if len(style_data_compare.columns) > 4:
                for col_idx in range(4, len(style_data_compare.columns)):
                    style_name = style_data_compare.iloc[1, col_idx]
                    if pd.notna(style_name) and str(style_name).strip() != "":
                        available_styles.append(str(style_name).strip())
            
            if available_styles:
                st.write("**このスタイルが適応した時期と適応しない時期での平均点数・平均時間の比較:**")
                selected_style = st.selectbox("比較するスタイルを選択", available_styles, key="style_for_compare")
                
                if st.button("スタイル比較"):
                    try:
                        # 選択されたスタイルの列インデックスを探す
                        style_col_idx = None
                        for col_idx in range(4, len(style_data_compare.columns)):
                            if str(style_data_compare.iloc[1, col_idx]).strip() == selected_style:
                                style_col_idx = col_idx
                                break
                        
                        if style_col_idx is None:
                            st.error("スタイルが見つかりません。")
                        else:
                            # スタイルが適応した時期（〇）とそうでない時期（✕または未入力）を分類
                            adapted_scores = []
                            adapted_times = []
                            not_adapted_scores = []
                            not_adapted_times = []
                            
                            adapted_periods = []
                            not_adapted_periods = []
                            
                            for row_idx in range(2, len(style_data_compare)):
                                style_value = str(style_data_compare.iloc[row_idx, style_col_idx]).strip()
                                score_val = style_data_compare.iloc[row_idx, 2]
                                time_val = style_data_compare.iloc[row_idx, 3]
                                period_val = str(style_data_compare.iloc[row_idx, 1]).strip() if pd.notna(style_data_compare.iloc[row_idx, 1]) else ""
                                
                                # スコアと時間が有効かチェック
                                if pd.notna(score_val) and pd.notna(time_val):
                                    try:
                                        score = float(score_val)
                                        time = float(time_val)
                                        
                                        if style_value == "〇":
                                            adapted_scores.append(score)
                                            adapted_times.append(time)
                                            adapted_periods.append(period_val)
                                        else:
                                            not_adapted_scores.append(score)
                                            not_adapted_times.append(time)
                                            not_adapted_periods.append(period_val)
                                    except ValueError:
                                        continue
                            
                            # 平均を計算
                            avg_adapted_score = sum(adapted_scores) / len(adapted_scores) if adapted_scores else None
                            avg_adapted_time = sum(adapted_times) / len(adapted_times) if adapted_times else None
                            avg_not_adapted_score = sum(not_adapted_scores) / len(not_adapted_scores) if not_adapted_scores else None
                            avg_not_adapted_time = sum(not_adapted_times) / len(not_adapted_times) if not_adapted_times else None
                            
                            # 結果テーブルを作成
                            if adapted_scores or not_adapted_scores:
                                comparison_data = {
                                    "項目": ["平均点数", "平均時間"],
                                    f"{selected_style}（適応）": [
                                        f"{avg_adapted_score:.2f}" if avg_adapted_score is not None else "N/A",
                                        f"{avg_adapted_time:.2f}" if avg_adapted_time is not None else "N/A"
                                    ],
                                    f"{selected_style}（非適応）": [
                                        f"{avg_not_adapted_score:.2f}" if avg_not_adapted_score is not None else "N/A",
                                        f"{avg_not_adapted_time:.2f}" if avg_not_adapted_time is not None else "N/A"
                                    ]
                                }
                                
                                # 効果を計算（%）
                                if avg_adapted_score is not None and avg_not_adapted_score is not None:
                                    score_effect = ((avg_adapted_score - avg_not_adapted_score) / avg_not_adapted_score * 100) if avg_not_adapted_score != 0 else 0
                                else:
                                    score_effect = None
                                
                                if avg_adapted_time is not None and avg_not_adapted_time is not None:
                                    time_effect = ((avg_adapted_time - avg_not_adapted_time) / avg_not_adapted_time * 100) if avg_not_adapted_time != 0 else 0
                                else:
                                    time_effect = None
                                
                                comparison_data["効果(%)"] = [
                                    f"{score_effect:.2f}%" if score_effect is not None else "N/A",
                                    f"{time_effect:.2f}%" if time_effect is not None else "N/A"
                                ]
                                
                                comparison_df = pd.DataFrame(comparison_data)
                                
                                st.subheader(f"'{selected_style}' の効果分析")
                                st.dataframe(comparison_df, use_container_width=True)
                                
                                # 効果の解釈
                                st.subheader("効果の解釈")
                                if score_effect is not None:
                                    if score_effect > 0:
                                        st.write(f"✅ **平均点数**: {score_effect:.2f}% 向上しました（スタイル適応の効果あり）")
                                    elif score_effect < 0:
                                        st.write(f"❌ **平均点数**: {abs(score_effect):.2f}% 低下しました（スタイル適応で悪化）")
                                    else:
                                        st.write(f"➡️ **平均点数**: 変わりません")
                                
                                if time_effect is not None:
                                    if time_effect > 0:
                                        st.write(f"⏱️ **平均時間**: {time_effect:.2f}% 増加しました（所要時間が増加）")
                                    elif time_effect < 0:
                                        st.write(f"⏱️ **平均時間**: {abs(time_effect):.2f}% 削減できました（時間効率が向上）")
                                    else:
                                        st.write(f"➡️ **平均時間**: 変わりません")
                            else:
                                st.info("比較するデータがありません。")
                    
                    except Exception as e:
                        st.error(f"スタイル比較に失敗しました: {e}")
            else:
                st.info("比較可能なスタイルがありません。")
        
        except Exception as e:
            st.error(f"スタイル比較セクションでエラーが発生しました: {e}")
    
    except FileNotFoundError:
        st.error("スタイル管理.xlsx が見つかりません。")
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")

