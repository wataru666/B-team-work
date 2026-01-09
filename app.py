import streamlit as st
import pandas as pd

# セッションステートでデータを管理
if 'grade_data' not in st.session_state:
    st.session_state.grade_data = pd.DataFrame(columns=['科目', '成績'])

if 'teaching_style' not in st.session_state:
    st.session_state.teaching_style = pd.DataFrame(columns=['科目', '授業スタイル'])

st.title("授業効率化システム")

# ページ選択
page = st.sidebar.selectbox("ページを選択", ["メインメニュー", "編集画面", "比較・閲覧画面"])

if page == "メインメニュー":
    st.write("成績データと授業スタイルデータを管理します。")

    st.write("サイドバーからページを選択してください。")

elif page == "編集画面":
    st.header("編集画面")
    st.subheader("成績データの編集")

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