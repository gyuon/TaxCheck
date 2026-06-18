import sys
import pandas as pd
sys.path.insert(0, ".")
from constants import Col, Status


def _calc_adjacent_month_amounts(df_errors, df_raw):
    if df_raw is None or df_errors is None or df_errors.empty or df_raw.empty:
        return df_errors
    if Col.CODE1 not in df_raw.columns:
        return df_errors
    df_work = df_raw.copy()
    df_work["canonical"] = df_work[Col.CODE1]
    valid = df_work[df_work["canonical"] > 0].copy()
    valid["serial"] = valid[Col.YEAR] * 12 + valid[Col.MONTH]
    grouped = valid.groupby([Col.NAME, "serial", "canonical"])[Col.RAW_DEPOSIT].sum().reset_index()
    lookup = {}
    for _, row in grouped.iterrows():
        lookup[(row[Col.NAME], row["serial"], row["canonical"])] = int(row[Col.RAW_DEPOSIT])
    df_out = df_errors.copy()
    if Col.PREV_MONTH_AMT not in df_out.columns:
        df_out[Col.PREV_MONTH_AMT] = None
    if Col.NEXT_MONTH_AMT not in df_out.columns:
        df_out[Col.NEXT_MONTH_AMT] = None
    for idx, row in df_out.iterrows():
        status = str(row.get(Col.STATUS, ""))
        if status not in (Status.INSUFFICIENT, Status.EXCESS):
            continue
        fund_name = str(row.get(Col.FUND_NAME, ""))
        canonical_map = {"운영": 1, "협력": 2, "복지": 3}
        canonical = canonical_map.get(fund_name)
        if canonical is None:
            continue
        name = row[Col.NAME]
        cur_serial = int(row[Col.YEAR]) * 12 + int(row[Col.MONTH])
        prev_serial = cur_serial - 1
        next_serial = cur_serial + 1
        prev_val = lookup.get((name, prev_serial, canonical))
        next_val = lookup.get((name, next_serial, canonical))
        df_out.at[idx, Col.PREV_MONTH_AMT] = prev_val if prev_val is not None else 0
        df_out.at[idx, Col.NEXT_MONTH_AMT] = next_val if next_val is not None else 0
    return df_out


def make_raw(rows):
    return pd.DataFrame(rows)


def make_error(rows):
    return pd.DataFrame(rows, columns=[Col.NAME, Col.YEAR, Col.MONTH, Col.FUND_NAME, Col.CODE,
                                        Col.DEPOSIT, Col.STANDARD, Col.FORMULA, Col.STATUS, Col.DIFF,
                                        Col.PREV_MONTH_AMT, Col.NEXT_MONTH_AMT, Col.REMARKS])


def test(name, df_errors, df_raw, check_fn):
    result = _calc_adjacent_month_amounts(df_errors, df_raw)
    if check_fn(result):
        print(f"  [PASS] {name}")
        return True
    else:
        print(f"  [FAIL] {name}")
        return False


print("=" * 60)
print("TEST: _calc_adjacent_month_amounts (컬럼 분리 버전)")
print("=" * 60)

passed = 0
total = 0

total += 1
df_e = make_error([["UserA", 2026, 1, "운영", "11", 100000, 70000, "산출법", Status.EXCESS, 30000, None, None, ""]])
df_r = make_raw([
    {Col.NAME: "UserA", Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 50000},
    {Col.NAME: "UserA", Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.CODE2: 2, Col.RAW_DEPOSIT: 20000},
    {Col.NAME: "UserA", Col.YEAR: 2026, Col.MONTH: 2, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 60000},
])
if test("운영기금 초과 - 전후달 존재 (코드1=1 합산)", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [70000] and r[Col.NEXT_MONTH_AMT].tolist() == [60000] and r[Col.REMARKS].tolist() == [""]):
    passed += 1

total += 1
df_e = make_error([["UserA", 2026, 1, "협력", "21", 30000, 40000, "산출법", Status.INSUFFICIENT, -10000, None, None, ""]])
df_r = make_raw([
    {Col.NAME: "UserA", Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 2, Col.CODE2: 1, Col.RAW_DEPOSIT: 40000},
    {Col.NAME: "UserA", Col.YEAR: 2026, Col.MONTH: 2, Col.CODE1: 2, Col.CODE2: 1, Col.RAW_DEPOSIT: 50000},
])
if test("협력기금 부족 - 전후달 존재", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [40000] and r[Col.NEXT_MONTH_AMT].tolist() == [50000]):
    passed += 1

total += 1
df_e = make_error([["UserB", 2026, 3, "운영", "11", 80000, 70000, "산출법", Status.EXCESS, 10000, None, None, ""]])
df_r = make_raw([
    {Col.NAME: "UserB", Col.YEAR: 2026, Col.MONTH: 2, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 70000},
])
if test("운영기금 초과 - 이후달 없음 → 0", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [70000] and r[Col.NEXT_MONTH_AMT].tolist() == [0]):
    passed += 1

total += 1
df_e = make_error([["UserC", 2026, 12, "복지", "31", 0, 70000, "산출법", Status.INSUFFICIENT, -70000, None, None, ""]])
df_r = make_raw([
    {Col.NAME: "UserC", Col.YEAR: 2027, Col.MONTH: 1, Col.CODE1: 3, Col.CODE2: 1, Col.RAW_DEPOSIT: 70000},
])
if test("복지기금 부족 - 12월→1월 연도 경계", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [0] and r[Col.NEXT_MONTH_AMT].tolist() == [70000]):
    passed += 1

total += 1
df_e = make_error([
    ["UserD", 2026, 1, "운영", "11~17", 0, None, "", Status.UNPAID, None, None, None, "기존비고"],
])
df_r = make_raw([
    {Col.NAME: "UserD", Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 70000},
])
if test("미납은 전/익월 컬럼 수정 안함", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [None] and r[Col.NEXT_MONTH_AMT].tolist() == [None] and r[Col.REMARKS].tolist() == ["기존비고"]):
    passed += 1

total += 1
df_e = make_error([["UserE", 2026, 6, "운영", "11", 100000, 70000, "산출법", Status.EXCESS, 30000, None, None, ""]])
df_r = make_raw([
    {Col.NAME: "UserE", Col.YEAR: 2026, Col.MONTH: 6, Col.CODE1: 1, Col.CODE2: 3, Col.RAW_DEPOSIT: 100000},
])
if test("당월 원본 데이터는 전후달에서 제외", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [0] and r[Col.NEXT_MONTH_AMT].tolist() == [0]):
    passed += 1

total += 1
df_e = make_error([["UserF", 2026, 6, "운영", "11", 100000, 70000, "산출법", Status.EXCESS, 30000, None, None, ""]])
if test("df_raw=None → 변경 없음", df_e, None,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [None] and r[Col.NEXT_MONTH_AMT].tolist() == [None]):
    passed += 1

total += 1
df_e = make_error([["UserF", 2026, 6, "운영", "11", 100000, 70000, "산출법", Status.EXCESS, 30000, None, None, ""]])
if test("df_raw=빈 DF → 변경 없음", df_e, pd.DataFrame(),
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [None] and r[Col.NEXT_MONTH_AMT].tolist() == [None]):
    passed += 1

total += 1
df_e = make_error([["UserF", 2026, 6, "운영", "11", 100000, 70000, "산출법", Status.EXCESS, 30000, None, None, ""]])
df_r = make_raw([{Col.NAME: "UserF", Col.YEAR: 2026, Col.MONTH: 5, Col.RAW_DEPOSIT: 70000}])
if test("코드1 컬럼 없음 → 변경 없음", df_e, df_r,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [None] and r[Col.NEXT_MONTH_AMT].tolist() == [None]):
    passed += 1

total += 1
df_e = pd.DataFrame()
df_r = make_raw([{Col.NAME: "X", Col.YEAR: 2026, Col.MONTH: 1, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 70000}])
if test("df_errors 빈 DF", df_e, df_r,
       lambda r: r.empty):
    passed += 1

total += 1
df_e1 = make_error([
    ["UserA", 2026, 1, "운영", "11", 100000, 70000, "산출법", Status.EXCESS, 30000, None, None, ""],
    ["UserB", 2026, 3, "협력", "21", 30000, 40000, "산출법", Status.INSUFFICIENT, -10000, None, None, ""],
])
df_r1 = make_raw([
    {Col.NAME: "UserA", Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 50000},
    {Col.NAME: "UserA", Col.YEAR: 2026, Col.MONTH: 2, Col.CODE1: 1, Col.CODE2: 1, Col.RAW_DEPOSIT: 60000},
    {Col.NAME: "UserB", Col.YEAR: 2026, Col.MONTH: 2, Col.CODE1: 2, Col.CODE2: 1, Col.RAW_DEPOSIT: 40000},
    {Col.NAME: "UserB", Col.YEAR: 2026, Col.MONTH: 4, Col.CODE1: 2, Col.CODE2: 1, Col.RAW_DEPOSIT: 50000},
])
if test("다중 행 - 각각 다른 기금", df_e1, df_r1,
       lambda r: r[Col.PREV_MONTH_AMT].tolist() == [50000, 40000] and r[Col.NEXT_MONTH_AMT].tolist() == [60000, 50000]):
    passed += 1

print(f"\n{'=' * 60}")
print(f"결과: {passed}/{total} 통과")
print(f"{'=' * 60}")
