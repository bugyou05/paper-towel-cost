60
61
62
63
64
65
66
67
68
69
70
71
72
73
74
75
76
77
78
79
80
81
82
83
84
85
86
87
88
89
90
91
92
93
94
95
96
97
98
99
100
101
102
103
104
105
106
107
108
109
110
111
112
113
114
115
116
117
118
119
120
121
122
123
124
125
126
127
128
129
130
131
132
133
134
135
136
137
138
139
140
141
import streamlit as st
        st.error("使用可能な略符データがありません。")
        st.stop()

    product_choices = [key for key in usage_by_product.keys() if key != "新エルナ"]
    product_choices_display = [f"{p}（{origin_by_product.get(p, '?')}）" for p in product_choices]
    product_choice_map = dict(zip(product_choices_display, product_choices))
    selected_display = st.selectbox("比較対象製品を選んでください", product_choices_display)
    target_product = product_choice_map[selected_display]

    monthly_cases = st.number_input("現在の出荷ケース数（月間）", value=50)
    st.markdown("### 単価入力（200枚あたり）")
    new_price_per_pack = st.number_input("新エルナ 単価", value=79.0, step=0.1, format="%.1f")
    target_price_per_pack = st.number_input(f"{target_product} 単価", value=70.0, step=0.1, format="%.1f")

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

def calculate_case_uses(product):
    total_sheets_per_case = product["pack_size"] * product["packs_per_case"]
    return total_sheets_per_case / product["daily_usage"]

# 正しい必要ケース数（使える回数ベース）
target_case_uses = calculate_case_uses(products[target_product])
new_case_uses = calculate_case_uses(products["新エルナ"])

total_required_uses = target_case_uses * monthly_cases
new_required_cases = total_required_uses / new_case_uses

new_monthly_cost = new_required_cases * new_price_per_pack * products["新エルナ"]["packs_per_case"]
target_monthly_cost = monthly_cases * target_price_per_pack * products[target_product]["packs_per_case"]
diff = target_monthly_cost - new_monthly_cost
rate = (diff / target_monthly_cost) * 100

st.subheader("📊 1人1日あたりのコスト")
df_table = pd.DataFrame({
    "製品": ["新エルナ", f"{target_product}（{origin_by_product.get(target_product, '?')}）"],
    "使用枚数": [f"{products['新エルナ']['daily_usage']:.2f}", f"{products[target_product]['daily_usage']:.2f}"],
    "単価（○枚）": [f"{new_price_per_pack:.1f}", f"{target_price_per_pack:.1f}"],
    "枚数/パック": [products["新エルナ"]["pack_size"], products[target_product]["pack_size"]],
    "1人1日コスト (円)": [
        f"{products['新エルナ']['daily_usage'] * new_price_per_pack / products['新エルナ']['pack_size']:.2f}",
        f"{products[target_product]['daily_usage'] * target_price_per_pack / products[target_product]['pack_size']:.2f}"
    ]
})

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
else:
    st.warning(f"差額：{diff:.0f}円（約{rate:.1f}% 増加）")
    st.markdown("⚠️ **新エルナは削減効果が見られません。使用条件をご確認ください。**")

st.caption("ver 4.6 - 使える回数ベースでケース数を算出する方式に修正")