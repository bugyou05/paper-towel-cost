import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="紙タオル ランニングコスト比較（Excel連携）", layout="centered")

st.title("🧻 紙タオル ランニングコスト比較アプリ（Excel製品連携版）")

st.markdown("""
このアプリでは、Excelファイルから有効データを抽出し、
略称ごとに平均使用枚数を元にコスト比較を行います。

※ 推定使用枚数と事務所人数の両方が揃ったデータのみ使用
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

    df_valid = df.dropna(subset=["推定使用枚数", "事務所人数"])
    df_valid = df_valid[df_valid["事務所人数"] > 0]  # 0除算防止
    df_valid["1人あたり使用枚数"] = df_valid["推定使用枚数"] / df_valid["事務所人数"]
    usage_by_product = df_valid.groupby("略称")["1人あたり使用枚数"].mean().to_dict()
    pack_size_by_product = df_valid.groupby("略称")["枚数"].first().to_dict()

    return usage_by_product, pack_size_by_product

try:
    usage_by_product, pack_size_by_product = load_data()
except Exception as e:
    st.error(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
    st.stop()

# 入力：対象製品選択
with st.sidebar:
    st.header("📋 比較製品を選択")
    if not usage_by_product:
        st.error("使用可能な略称データがありません。")
        st.stop()
    target_product = st.selectbox("比較対象製品を選んでください", list(usage_by_product.keys()))
    monthly_cases = st.number_input("現在の出荷ケース数（月間）", value=50)
    st.markdown("### 単価入力（200枚あたり）")
    new_price_per_pack = st.number_input("新エルナ 単価", value=79.0, format="%.1f")
    target_price_per_pack = st.number_input(f"{target_product} 単価", value=70.0, format="%.1f")

# 製品情報
products = {
    "新エルナ": {
        "daily_usage": 6.71,
        "pack_size": 200,
        "packs_per_case": 35,
        "price_per_pack": new_price_per_pack
    },
    target_product: {
        "daily_usage": usage_by_product[target_product],
        "pack_size": pack_size_by_product.get(target_product, 200),
        "packs_per_case": 40,
        "price_per_pack": target_price_per_pack
    }
}

# 計算処理
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
st.table(pd.DataFrame({
    "製品": ["新エルナ", target_product],
    "使用枚数": [round(products["新エルナ"]["daily_usage"], 1), round(products[target_product]["daily_usage"], 1)],
    "単価（◯枚）": [round(new_price_per_pack, 1), round(target_price_per_pack, 1)],
    "枚数/パック": [products["新エルナ"]["pack_size"], products[target_product]["pack_size"]],
    "1人1日コスト (円)": [round(new_daily, 2), round(target_daily, 2)]
}))

st.subheader("📦 月間コスト比較")
st.write(f"{target_product}：{monthly_cases:.2f}ケース × {target_case:.0f}円 = {target_monthly_cost:.0f}円")
st.write(f"新エルナ：約{new_required_cases:.2f}ケース × {new_case:.0f}円 = {new_monthly_cost:.0f}円")

if diff > 0:
    st.success(f"差額：{diff:.0f}円（約{rate:.1f}% 削減の見込み）")
    st.markdown("✅ **新エルナはコスト削減につながる可能性があります。**")
    st.markdown("📝 使用枚数の削減により、発注回数や保管スペースの削減、交換頻度の低減なども期待できます。")
else:
    st.warning(f"差額：{diff:.0f}円（約{rate:.1f}% 増加）")
    st.markdown("⚠️ **新エルナは削減効果が見られません。使用条件をご確認ください。**")

st.caption("ver 3.7.0 - 表示形式調整、枚数列参照")
