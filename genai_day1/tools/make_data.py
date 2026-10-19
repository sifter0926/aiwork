"""종합 예제용 가상 데이터 생성 + 분석 수치 계산 (슬라이드·노트북 공통 기준값)."""
import json
import sys
from datetime import date, timedelta

import numpy as np
import pandas as pd

seed = int(sys.argv[1]) if len(sys.argv) > 1 else 7
out_dir = sys.argv[2] if len(sys.argv) > 2 else "."
rng = np.random.default_rng(seed)

days = [date(2026, 9, 1) + timedelta(d) for d in range(30)]
yo = "월화수목금토일"
rows = []
for i, d in enumerate(days):
    weekend = int(d.weekday() >= 5)
    temp = round(31.5 - 0.22 * i + rng.normal(0, 1.4), 1)
    rain = 0
    if rng.random() < 0.25:
        rain = int(rng.integers(3, 45))
    visitors = 250 + 85 * weekend - 2.6 * rain + 4 * (temp - 27) + rng.normal(0, 14)
    ice = -245 + 13.5 * temp + 0.32 * visitors + rng.normal(0, 11)
    hot = 135 - 3.6 * temp + 0.14 * visitors + rng.normal(0, 7)
    sales = visitors * 9.4 + rng.normal(0, 130)  # 천원
    rows.append({
        "날짜": d.isoformat(), "요일": yo[d.weekday()], "주말": weekend,
        "최고기온": temp, "강수량": rain, "방문객수": round(visitors),
        "아이스음료": round(ice), "따뜻한음료": round(hot), "매출액": round(sales),
    })
df = pd.DataFrame(rows)

num = ["최고기온", "강수량", "방문객수", "아이스음료", "따뜻한음료", "매출액"]
corr = df[num].corr().round(2)

x, y = df["최고기온"], df["아이스음료"]
slope, intercept = np.polyfit(x, y, 1)
r = x.corr(y)

X = np.column_stack([np.ones(len(df)), df["최고기온"], df["강수량"], df["주말"]])
beta, *_ = np.linalg.lstsq(X, df["방문객수"], rcond=None)
pred = X @ beta
ss_res = ((df["방문객수"] - pred) ** 2).sum()
ss_tot = ((df["방문객수"] - df["방문객수"].mean()) ** 2).sum()
r2_multi = 1 - ss_res / ss_tot
r2_simple_visit = df["최고기온"].corr(df["방문객수"]) ** 2

stats = {
    "corr": corr.to_dict(),
    "simple": {"slope": slope, "intercept": intercept, "r": r, "r2": r ** 2,
               "pred28": intercept + slope * 28, "xmin": x.min(), "xmax": x.max()},
    "multi": {"b0": beta[0], "b_temp": beta[1], "b_rain": beta[2], "b_weekend": beta[3],
              "r2": r2_multi, "r2_temp_only": r2_simple_visit},
    "desc": df[num].describe().round(1).to_dict(),
    "rainy_days": int((df["강수량"] > 0).sum()),
}
df.to_csv(f"{out_dir}/cafe_daily_202609.csv", index=False, encoding="utf-8-sig")
json.dump(stats, open(f"{out_dir}/stats.json", "w"), ensure_ascii=False, indent=1, default=float)
print(df.head(8).to_string())
print(corr.to_string())
print(json.dumps({k: stats[k] for k in ("simple", "multi", "rainy_days")}, indent=1, default=float))
