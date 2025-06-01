import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="紙タオル ランニングコスト比較", layout="centered")

st.title("🧻 紙タオル ランニングコスト比較アプリ")

st.markdown("""
※ 実使用に基づく5日間以上のデータから平均を算出しています。

なお、すべての製品に対して同一条件で比較を行っており、
信頼性向上のため今後も継続的にデータ取得を進めていきます。
""")

# GitHub用：相対パスでExcelファイルを参照
excel_path = "使用量調査.xlsx"

# Excelファイル読み込み
@st.cache_data
def load_data():
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path, engine="openpyxl")
    else:
        st.error("Excelファイルが見つかりません。'使用量調査.xlsx' をこのアプリと同じフォルダに配置してください。")
        st.stop()

    df.columns = df.columns.str.strip()  # 列名の空白除去
    required_columns = ["商品コード", "略称", "推定使用枚数", "事務所人数", "枚数", "入数"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        st.error(f"Excelに必要な列が見つかりません：{', '.join(missing)}")
        st.stop()

    df_valid = df.dropna(subset=["推定使用枚数", "事務所人数"])
    df_valid = df_valid[df_valid["事務所人数"] > 0]  # 0除算防止
    df_valid["1人あたり使用枚数"] = df_valid["推定使用枚数"] / df_valid["事務所人数"]
    usage_by_product = df_valid.groupby("略称")["1人あたり使用枚数"].mean().to_dict()
    pack_size_by_product = df_valid.groupby("略称")["枚数"].first().to_dict()
    packs_per_case_by_product = df_valid.groupby("略称")["入数"].first().to_dict()

    product_code_map = df_valid.groupby("略称")["商品コード"].first().to_dict()

    return usage_by_product, pack_size_by_product, packs_per_case_by_product, product_code_map

try:
    usage_by_product, pack_size_by_product, packs_per_case_by_product, product_code_map = load_data()
except Exception as e:
    st.error(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
    st.stop()

# 入力：対象製品選択
with st.sidebar:
    st.header("📋 比較製品を選択")
    if not usage_by_product:
        st.error("使用可能な略称データがありません。")
        st.stop()
    product_choices = [key for key in usage_by_product.keys() if key != "新エルナ"]
    target_product = st.selectbox("比較対象製品を選んでください", product_choices)
    monthly_cases = st.number_input("現在の出荷ケース数（月間）", value=50)
    st.markdown("### 単価入力（200枚あたり）")
    new_price_per_pack = st.number_input("新エルナ 単価", value=79.0, step=0.1, format="%.1f")
    target_price_per_pack = st.number_input(f"{target_product} 単価", value=70.0, step=0.1, format="%.1f")

# 製品情報（新エルナも略称から平均使用枚数を取得）
products = {
    "新エルナ": {
        "daily_usage": usage_by_product.get("新エルナ", 6.71),
        "pack_size": pack_size_by_product.get("新エルナ", 200),
        "packs_per_case": packs_per_case_by_product.get("新エルナ", 35),
        "price_per_pack": new_price_per_pack
    },
    target_product: {
        "daily_usage": usage_by_product.get(target_product, 0),
        "pack_size": pack_size_by_product.get(target_product, 200),
        "packs_per_case": packs_per_case_by_product.get(target_product, 40),
        "price_per_pack": target_price_per_pack
    }
}

# 計算処理（ケース単価＝単価×入数）
def calculate_cost(product):
    unit_price = product["price_per_pack"] / product["pack_size"]
    daily_cost = product["daily_usage"] * unit_price
    case_price = product["price_per_pack"] * product["packs_per_case"]
    return unit_price, daily_cost, case_price

new_unit, new_daily, new_case = calculate_cost(products["新エルナ"])
target_unit, target_daily, target_case = calculate_cost(products[target_product])

# 月間コスト比較
new_required_cases = monthly_cases * (products["新エルナ"]["daily_usage"] / products[target_product]["daily_usage"])
new_monthly_cost = new_required_cases * new_case
target_monthly_cost = monthly_cases * target_case
diff = target_monthly_cost - new_monthly_cost
rate = (diff / target_monthly_cost) * 100

# 結果表示
st.subheader("📊 1人1日あたりのコスト")
df_table = pd.DataFrame({
    "製品": ["新エルナ", target_product],
    "使用枚数": [f"{products['新エルナ']['daily_usage']:.2f}", f"{products[target_product]['daily_usage']:.2f}"],
    "単価（◯枚）": [f"{new_price_per_pack:.1f}", f"{target_price_per_pack:.1f}"],
    "枚数/パック": [products["新エルナ"]["pack_size"], products[target_product]["pack_size"]],
    "1人1日コスト (円)": [f"{new_daily:.2f}", f"{target_daily:.2f}"]
})

# 表示スタイル指定（製品：左寄せ、それ以外：中央）
css_style = """
<style>
    table td:nth-child(n+2), table th:nth-child(n+2) {
        text-align: center !important;
    }
    table td:first-child, table th:first-child {
        text-align: left !important;
    }
</style>
"""
st.markdown(css_style, unsafe_allow_html=True)
st.table(df_table)

st.subheader("📦 月間コスト比較")
st.write(f"{target_product}：{monthly_cases:.2f}ケース × {target_price_per_pack:.0f}円 × {products[target_product]['packs_per_case']}パック = {target_monthly_cost:.0f}円")
st.write(f"新エルナ：約{new_required_cases:.2f}ケース × {new_price_per_pack:.0f}円 × {products['新エルナ']['packs_per_case']}パック = {new_monthly_cost:.0f}円")

if diff > 0:
    st.success(f"差額：{diff:.0f}円（約{rate:.1f}% 削減の見込み）")
    st.markdown("✅ **新エルナはコスト削減につながる可能性があります。**")

    if product_code_map.get(target_product) == 9962860:
        st.markdown("📝 天候や使用状況による多少の変動はありますが、")
        st.markdown("**傾向として『新エルナは使用枚数が明らかに減っている』ことが確認されています。**")
        st.markdown("📝 使用枚数の削減により、**発注回数や保管スペースの圧縮、交換頻度の低減**なども期待できます。")
    else:
        st.markdown("📝 使用枚数の削減により、**発注回数や保管スペースの圧縮、交換頻度の低減**なども期待できます。")
else:
    st.warning(f"差額：{diff:.0f}円（約{rate:.1f}% 増加）")
    st.markdown("⚠️ **新エルナは削減効果が見られません。使用条件をご確認ください。**")

st.caption("ver 4.3 - 識別を商品コードベースに変更")
