"""patch_ui.py - ปรับ UI ของ app.py ทีเดียว (สำรองไฟล์เดิมเป็น app_backup.py ให้ก่อน)
วิธีใช้: วางไฟล์นี้ในโฟลเดอร์เดียวกับ app.py แล้วรัน  python patch_ui.py
"""
import os
import re
import shutil

SRC = "app.py"
if not os.path.exists(SRC):
    raise SystemExit("ไม่พบ app.py - ให้วาง patch_ui.py ไว้โฟลเดอร์เดียวกับ app.py")

shutil.copy(SRC, "app_backup.py")
with open(SRC, encoding="utf-8") as f:
    src = f.read()

done, skipped = [], []


def sub_region(name, pattern, new):
    global src
    out, n = re.subn(pattern, lambda m: new, src, count=1, flags=re.S)
    if n:
        src = out
        done.append(name)
    else:
        skipped.append(name)


def replace(name, old, new):
    global src
    if old in src:
        src = src.replace(old, new, 1)
        done.append(name)
    else:
        skipped.append(name)


# 1) ข้อความแบนเนอร์ให้ตรงความจริง + หมายเหตุความเป็นส่วนตัวใน footer
replace("แบนเนอร์", "⚡ ขับเคลื่อนด้วย Machine Learning", "⚡ ประมวลผลภาพอัตโนมัติ")
replace("footer",
        "พัฒนาด้วย Python · OpenCV · Streamlit — โปรเจควิชา Machine Learning",
        "พัฒนาด้วย Python · OpenCV · Streamlit — โปรเจควิชา Machine Learning<br>"
        "ไฟล์ที่อัปโหลดใช้ประมวลผลชั่วคราวและไม่ถูกจัดเก็บถาวร")

# 2) หัวข้อ "ผลการตรวจ" ก่อนแสดงผล
if 'step("ผลการตรวจ")' not in src:
    replace("หัวข้อผลการตรวจ", "        bad = df[df[",
            '        step("ผลการตรวจ")\n        bad = df[df[')

# 3) ลบปุ่มดาวน์โหลดเดิมที่อยู่ล่างสุด
sub_region("ลบปุ่มดาวน์โหลดเดิม", r'        st\.download_button\(.*?type="primary"\)\n', "")

# 4) การ์ดสรุป 2 แถว แถวละ 4 + ปุ่มดาวน์โหลดไว้ใต้การ์ด
METRICS = '''        kr = kr20(key, students)
        m = st.columns(4)
        m[0].metric("ผู้เข้าสอบ (คน)", len(df))
        m[1].metric("คะแนนเฉลี่ย", f"{df['คะแนน'].mean():.2f}")
        m[2].metric("คะแนนสูงสุด", int(df["คะแนน"].max()))
        m[3].metric("คะแนนต่ำสุด", int(df["คะแนน"].min()))
        m2 = st.columns(4)
        m2[0].metric("มัธยฐาน", f"{df['คะแนน'].median():.1f}")
        m2[1].metric("S.D.", f"{df['คะแนน'].std():.2f}" if len(df) > 1 else "-")
        m2[2].metric("ร้อยละเฉลี่ย", f"{df['ร้อยละ'].mean():.1f}%")
        m2[3].metric("ความเชื่อมั่น (KR-20)", f"{kr:.2f}" if kr is not None else "-",
                     help="ยิ่งใกล้ 1 ยิ่งเชื่อถือได้ (โดยทั่วไป ≥ 0.70 ถือว่ายอมรับได้) ต้องมีผู้สอบตั้งแต่ 2 คน")
        st.download_button("ดาวน์โหลดรายงาน Excel", res["xlsx"],
                           file_name=f"exam_result_{datetime.now():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")

'''
sub_region("การ์ดสรุป+ปุ่มดาวน์โหลด",
           r"        kr = kr20\(key, students\)\n.*?(?=        # 3 อันดับแรก)", METRICS)

# 5) แท็บผลรายบุคคล: ค้นหา + เรียงลำดับ
TAB1 = '''        with tab1:
            f1, f2 = st.columns([3, 1])
            q = f1.text_input("ค้นหา (ชื่อไฟล์ หรือ เลขประจำตัวสอบ)", key="search_box")
            order = f2.selectbox("เรียงตาม", ["อันดับ", "ชื่อไฟล์"])
            view = df
            if q:
                view = view[view["ไฟล์"].str.contains(q, case=False, na=False)
                            | view["เลขประจำตัวสอบ"].astype(str).str.contains(q, case=False, na=False)]
            if order == "ชื่อไฟล์":
                view = view.sort_values("ไฟล์")
            st.caption(f"แสดง {len(view)} จาก {len(df)} ฉบับ")

            def hl(row):
                return ["background-color: #FFF2CC; font-weight: 600" if row["อันดับ"] <= 3 else ""] * len(row)
            st.dataframe(view.style.apply(hl, axis=1), width="stretch", hide_index=True,
                         column_config={"ร้อยละ": st.column_config.ProgressColumn(
                             "ร้อยละ", min_value=0, max_value=100, format="%.1f")})
'''
sub_region("แท็บผลรายบุคคล", r"        with tab1:.*?(?=        with tab2:)", TAB1)

# 6) กล่องวิธีใช้งาน (พับได้) ไว้ก่อนขั้นตอนที่ 1
if "วิธีใช้งานและข้อแนะนำ" not in src:
    HELP = '''with st.expander("📖 วิธีใช้งานและข้อแนะนำ"):
    st.markdown(
        "**รูปแบบเฉลยที่รับได้**\\n"
        "- เรียงต่อกัน เช่น `ABCDEABCDE` (ตามลำดับข้อ)\\n"
        "- ระบุเลขข้อ เช่น `1-A 2-C 3-B`\\n\\n"
        "**การสแกนกระดาษคำตอบ**\\n"
        "- สแกนหรือถ่ายให้กระดาษตรง ไม่เอียง ไม่มีเงา ความละเอียดอย่างน้อย 150 dpi\\n"
        "- รองรับไฟล์ PNG, JPG, PDF และเลือกอัปโหลดได้หลายไฟล์พร้อมกัน\\n\\n"
        "**ความหมายของสัญลักษณ์**\\n"
        "- `-` = ไม่ได้ตอบข้อนั้น\\n"
        "- `*` = ฝนมากกว่า 1 ตัวเลือก (นับเป็นผิด)\\n"
        "- `?` ในรหัส = อ่านรหัสไม่ได้ ควรตรวจกับกระดาษจริง")

'''
    replace("วิธีใช้งาน", "c1, c2 = st.columns(2, gap=\"large\")", HELP + "c1, c2 = st.columns(2, gap=\"large\")")

# 7) ปุ่มล้างผล / เริ่มชุดใหม่
if "ล้างผล / เริ่มชุดใหม่" not in src:
    CLEAR = '''if st.session_state.get("result") and st.button("ล้างผล / เริ่มชุดใหม่"):
    del st.session_state["result"]
    st.rerun()

'''
    replace("ปุ่มล้างผล", 'res = st.session_state.get("result")',
            CLEAR + 'res = st.session_state.get("result")')

with open(SRC, "w", encoding="utf-8") as f:
    f.write(src)

# 8) บังคับธีมสว่าง (ไม่เขียนทับถ้ามีไฟล์อยู่แล้ว)
cfg_dir, cfg = ".streamlit", os.path.join(".streamlit", "config.toml")
if not os.path.exists(cfg):
    os.makedirs(cfg_dir, exist_ok=True)
    with open(cfg, "w", encoding="utf-8") as f:
        f.write('[theme]\nbase = "light"\nprimaryColor = "#1F3864"\n'
                'backgroundColor = "#F6F8FC"\nsecondaryBackgroundColor = "#FFFFFF"\n'
                'textColor = "#1A1A1A"\n')
    done.append("config.toml (ธีมสว่าง)")
else:
    print('พบ .streamlit/config.toml อยู่แล้ว ไม่แตะต้อง (ตรวจว่ามี base = "light")')

print("ปะสำเร็จ:", ", ".join(done) or "-")
if skipped:
    print("ข้าม (หาจุดไม่เจอ):", ", ".join(skipped))
print("สำรองไฟล์เดิมไว้ที่ app_backup.py")
