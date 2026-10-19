"""Kiểm tra dữ liệu do build_data.py sinh ra.

CÁCH CHẠY:
    py scripts/check_data.py            # kiểm tra + in bảng
    py scripts/check_data.py --sample   # thêm 10 mục mẫu để xem bằng mắt

LỖI (dừng, thoát mã 1): từ thiếu vi/pos; mã trùng; bài dưới 6 từ; ô xlsx kết thúc bằng khoảng trắng.
CẢNH BÁO (không dừng): hai từ cùng bài có cùng sameMeaningGroup hoặc cùng confusableGroup.
"""
import json
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
XLSX = ROOT / "data-src" / "tu_vung_A1_theo_chu_de_v2.xlsx"


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    errors, warns = [], []
    topics = json.loads((DATA / "topics.json").read_text(encoding="utf-8"))
    ids = {}
    rows = []
    all_words = []

    for t in topics:
        d = json.loads((DATA / f"t{t['id']:02d}.json").read_text(encoding="utf-8"))
        nw = nt = 0
        for les in d["lessons"]:
            if len(les["words"]) < 6:
                errors.append(f"Bài {les['id']} chỉ có {len(les['words'])} từ (<6)")
            sm, cf = {}, {}
            for w in les["words"]:
                nw += 1
                nt += 0 if w["typeable"] else 1
                all_words.append(w)
                if not w["vi"]:
                    errors.append(f"{w['id']}: thiếu nghĩa tiếng Việt")
                if not w["pos"]:
                    errors.append(f"{w['id']}: thiếu loại từ")
                if w["id"] in ids:
                    errors.append(f"Mã trùng: {w['id']}")
                ids[w["id"]] = 1
                if w["sameMeaningGroup"]:
                    sm.setdefault(w["sameMeaningGroup"], []).append(w["fr"])
                if w["confusableGroup"]:
                    cf.setdefault(w["confusableGroup"], []).append(w["fr"])
            for g, lst in list(sm.items()) + list(cf.items()):
                if len(lst) > 1:
                    warns.append(f"Bài {les['id']}: cùng nhóm {g}: {' | '.join(lst)}")
        rows.append((t["id"], t["vi"], len(d["lessons"]), nw, nt))
        if (nw, len(d["lessons"])) != (t["wordCount"], t["lessonCount"]):
            errors.append(f"topics.json lệch với t{t['id']:02d}.json")

    if XLSX.exists():
        import openpyxl
        wb = openpyxl.load_workbook(XLSX, data_only=True)
        for ws in wb.worksheets[2:]:
            for row in ws.iter_rows(min_row=2):
                for c in row[:5]:
                    if isinstance(c.value, str) and c.value != c.value.rstrip():
                        errors.append(f"xlsx {ws.title}!{c.coordinate}: kết thúc bằng khoảng trắng")

    print(f"{'#':>2}  {'Chủ điểm':<40}{'Bài':>5}{'Từ':>6}{'Không gõ':>10}")
    for r in rows:
        print(f"{r[0]:>2}  {r[1]:<40}{r[2]:>5}{r[3]:>6}{r[4]:>10}")
    print(f"{'':>2}  {'TỔNG':<40}{sum(r[2] for r in rows):>5}{sum(r[3] for r in rows):>6}{sum(r[4] for r in rows):>10}")

    if warns:
        print(f"\nCẢNH BÁO ({len(warns)}):")
        for w in warns:
            print("  -", w)
    if errors:
        print(f"\nLỖI ({len(errors)}):")
        for e in errors:
            print("  -", e)

    if "--sample" in sys.argv:
        pick = []
        def first(pred):
            for w in all_words:
                if pred(w) and w not in pick:
                    pick.append(w)
                    return
        first(lambda w: w["fr"] == "l'étudiant / l'étudiante")
        first(lambda w: w["fr"] == "le dentiste / la dentiste")
        first(lambda w: "…" in w["fr"] or "..." in w["fr"])
        first(lambda w: w["fr"] == "la France")
        first(lambda w: w["vi"] != w["viMain"])
        rnd = random.Random()
        while len(pick) < 10:
            w = rnd.choice(all_words)
            if w not in pick:
                pick.append(w)
        print("\n10 MỤC MẪU:")
        for w in pick:
            print(json.dumps(w, ensure_ascii=False))

    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
