import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io

# データの保存先（一時保存用）
CSV_FILE = "evaluations_data.csv"

def save_data(params):
    """URLパラメータからデータを抽出し、CSVに追記する"""
    # フォームのname属性に対応するデータを取得
    new_data = {
        "送信日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "クラス名": params.get("className", ""),
        "評価者名": params.get("evaluatorName", ""),
        "評価対象者名": params.get("targetName", ""),
        "Q1_車いす点検": params.get("q1", ""),
        "Q2_杖歩行器点検": params.get("q2", ""),
        "Q3_事前説明": params.get("q3", ""),
        "Q4_声かけのトーン": params.get("q4", ""),
        "Q5_状態確認": params.get("q5", ""),
        "Q6_車いす展開": params.get("q6", ""),
        "Q7_段差越え": params.get("q7", ""),
        "Q8_スロープ": params.get("q8", ""),
        "Q9_杖の調整": params.get("q9", ""),
        "Q10_介助位置": params.get("q10", ""),
        "Q11_安全確保": params.get("q11", ""),
        "Q12_利用者体験": params.get("q12", ""),
        "コメント": params.get("comments", "")
    }

    df_new = pd.DataFrame([new_data])
    
    # 既存のCSVがあれば追記、なければ新規作成
    if os.path.exists(CSV_FILE):
        df_new.to_csv(CSV_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
    else:
        df_new.to_csv(CSV_FILE, mode='w', header=True, index=False, encoding='utf-8-sig')

def main():
    st.set_page_config(page_title="実技評価 集計システム", layout="wide")
    
    # Streamlit 1.30以降の新しいクエリパラメータ取得方法
    params = st.query_params

    # パラメータに 'className' が含まれていれば、学生からのフォーム送信と判定
    if "className" in params:
        try:
            save_data(params)
            st.success("✅ 評価データの送信が完了しました！")
            st.info("この画面（タブ）は閉じてください。")
            
            # 再読み込み時の二重送信を防ぐためパラメータをクリア
            st.query_params.clear()
        except Exception as e:
            st.error("送信エラーが発生しました。教員に報告してください。")
            
    # パラメータがない場合は、教員用のダッシュボードを表示
    else:
        st.title("📋 実技評価 集計ダッシュボード")
        
        if os.path.exists(CSV_FILE):
            df = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
            st.write(f"**現在の提出件数: {len(df)} 件**")
            st.dataframe(df)
            
            # Excel出力用のバイナリデータ作成
            towrite = io.BytesIO()
            df.to_excel(towrite, index=False, engine='openpyxl')
            towrite.seek(0)
            
            # 既存のExcelマクロ等に流し込みやすいようxlsx形式でダウンロード
            st.download_button(
                label="📥 集計データをExcel(.xlsx)でダウンロード",
                data=towrite,
                file_name=f"実技評価集計_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # データリセット機能（誤操作防止のためエクスパンダー内に配置）
            with st.expander("⚙️ 管理メニュー（データのリセット）"):
                if st.button("全データを削除する", type="primary"):
                    os.remove(CSV_FILE)
                    st.rerun()
        else:
            st.info("現在、提出されたデータはありません。")

if __name__ == "__main__":
    main()
