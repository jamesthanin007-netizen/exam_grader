"""
omr_reader.py  -  อ่านกระดาษคำตอบ OMR แบบ R1101 (170 ข้อ x 5 ตัวเลือก)
ขั้นตอน: โหลดภาพ -> หา timing marks -> warp ให้ตรงกับ template -> ตัดช่อง -> วัดความเข้ม
พิกัดทั้งหมดอิงภาพ A4 ที่ 200 DPI (1654 x 2339 px)
"""
import cv2
import numpy as np

# ---------------------------------------------------------------- layout ----
DPI = 200
PAGE_W, PAGE_H = 1654, 2339
PITCH = 43.3          # ระยะห่างวงกลม (ทั้งแนวตั้งและแนวนอน)
RADIUS = 14           # รัศมีวงกลม (px)

# anchor ของ timing marks บน template (x, y) เรียงเป็น TL, BL, BOTTOM-LEFT, BOTTOM-RIGHT
# = มาร์กบนสุดของแถวซ้าย, มาร์กล่างสุดของแถวซ้าย, มาร์กซ้ายสุด/ขวาสุดของแถวล่าง
TEMPLATE_ANCHORS = np.float32([[83, 228], [83, 2177], [169, 2263.5], [1555, 2263.5]])

# บล็อกคำตอบ: (ข้อแรก, จำนวนข้อ, x ของตัวเลือกแรก, y ของข้อแรก)
ANSWER_BLOCKS = [
    (1,   25, 213.0,  1051.0),
    (26,  25, 516.1,  1051.0),
    (51,  40, 862.4,   401.1),
    (91,  40, 1122.1,  401.1),
    (131, 40, 1381.7,  401.1),
]
N_CHOICES = 5   # 1-5 (= A-E)

# ช่องรหัส: (x คอลัมน์แรก, y แถวเลข 0, จำนวนคอลัมน์)  แต่ละคอลัมน์มีเลข 0-9
SUBJECT_CODE = (169.3, 574.3, 6)
EXAM_ID      = (472.5, 574.3, 8)


# ------------------------------------------------------------- alignment ----
def _find_anchors(gray):
    """หา timing marks แล้วคืนค่า 4 anchor ตามลำดับเดียวกับ TEMPLATE_ANCHORS"""
    H, W = gray.shape
    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, _, st, cen = cv2.connectedComponentsWithStats(bw)
    scale = W / PAGE_W
    left, bottom = [], []
    for i in range(1, n):
        x, y, w, h, a = st[i]
        if a < 150 * scale**2 or a / (w * h) < 0.85:   # ต้องเป็นสี่เหลี่ยมทึบ
            continue
        if x < W * 0.08 and w > 1.5 * h:
            left.append(cen[i])
        elif y > H * 0.94 and h > 1.5 * w:
            bottom.append(cen[i])
    if len(left) < 10 or len(bottom) < 10:
        raise RuntimeError("หา timing marks ไม่พอ - ตรวจว่าภาพเห็นขอบซ้ายและขอบล่างครบ")
    left = sorted(left, key=lambda p: p[1])
    bottom = sorted(bottom, key=lambda p: p[0])
    return np.float32([left[0], left[-1], bottom[0], bottom[-1]])


def align(image_bgr):
    """warp ภาพให้ตรงกับ template (PAGE_W x PAGE_H)"""
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    src = _find_anchors(gray)
    M, _ = cv2.findHomography(src, TEMPLATE_ANCHORS)
    return cv2.warpPerspective(image_bgr, M, (PAGE_W, PAGE_H), borderValue=(255, 255, 255))


# ------------------------------------------------------------ extraction ----
def _bubble_features(binary, gray, cx, cy):
    """ตัดช่องรอบ (cx, cy) แล้วคืนฟีเจอร์ [fill_ratio, mean_dark, ...]"""
    r = RADIUS
    x0, y0 = int(round(cx)) - r, int(round(cy)) - r
    patch_b = binary[y0:y0 + 2 * r, x0:x0 + 2 * r]
    patch_g = gray[y0:y0 + 2 * r, x0:x0 + 2 * r]
    mask = np.zeros(patch_b.shape, np.uint8)
    cv2.circle(mask, (r, r), r - 4, 255, -1)          # ตัดขอบวงกลมพิมพ์ออก
    inside = mask > 0
    fill = float(patch_b[inside].mean() / 255.0)       # สัดส่วนพิกเซลดำ
    dark = float(1.0 - patch_g[inside].mean() / 255.0) # ความเข้มเฉลี่ย
    return fill, dark


def extract_bubbles(aligned_bgr):
    """คืน dict: 'answers' (170,5,2), 'subject' (6,10,2), 'exam_id' (8,10,2)
    แต่ละช่องเป็น [fill_ratio, mean_darkness]  -> ใช้เป็นฟีเจอร์ให้โมเดล ML"""
    gray = cv2.cvtColor(aligned_bgr, cv2.COLOR_BGR2GRAY)
    # ใช้ threshold ต่ำเพื่อให้เส้นวงกลมสีม่วงพิมพ์ไม่นับเป็นรอยฝน
    _, binary = cv2.threshold(gray, 110, 255, cv2.THRESH_BINARY_INV)

    ans = np.zeros((170, N_CHOICES, 2), np.float32)
    for first, count, x0, y0 in ANSWER_BLOCKS:
        for row in range(count):
            for c in range(N_CHOICES):
                ans[first - 1 + row, c] = _bubble_features(
                    binary, gray, x0 + c * PITCH, y0 + row * PITCH)

    def grid(spec):
        x0, y0, ncol = spec
        out = np.zeros((ncol, 10, 2), np.float32)
        for col in range(ncol):
            for d in range(10):
                out[col, d] = _bubble_features(binary, gray, x0 + col * PITCH, y0 + d * PITCH)
        return out

    return {"answers": ans, "subject": grid(SUBJECT_CODE), "exam_id": grid(EXAM_ID)}


# --------------------------------------------------------- rule baseline ----
def pick_answers(feat, marked_thr=0.35, margin=0.15):
    """baseline ไม่ใช้ ML: เลือกช่องที่เข้มสุดถ้าเกินเกณฑ์
    คืน list ตัวอักษร A-E, '-' = ว่าง, '*' = ฝนหลายช่อง/กำกวม"""
    out = []
    for q in feat:
        fill = q[:, 0]
        order = np.argsort(fill)[::-1]
        top, second = fill[order[0]], fill[order[1]]
        if top < marked_thr:
            out.append("-")
        elif second > marked_thr and top - second < margin:
            out.append("*")
        else:
            out.append("ABCDE"[order[0]])
    return out


def pick_digits(feat, marked_thr=0.35):
    """อ่านรหัสจากตาราง (คอลัมน์ x 10 หลัก) -> string, '?' ถ้าอ่านไม่ได้"""
    s = ""
    for col in feat:
        fill = col[:, 0]
        s += str(int(np.argmax(fill))) if fill.max() >= marked_thr else "?"
    return s


def _load_image(path):
    """โหลดภาพจาก PNG/JPG หรือ PDF (หน้าแรก) -> ภาพ BGR"""
    if path.lower().endswith(".pdf"):
        import pymupdf
        page = pymupdf.open(path)[0]
        pix = page.get_pixmap(dpi=DPI, colorspace=pymupdf.csRGB)
        arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)
        return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    img = cv2.imread(path)
    if img is None:
        raise FileNotFoundError(f"เปิดไฟล์ภาพไม่ได้: {path}")
    return img


def read_sheet(path):
    img = _load_image(path)
    aligned = align(img)
    feat = extract_bubbles(aligned)
    return {
        "answers": pick_answers(feat["answers"]),
        "subject_code": pick_digits(feat["subject"]),
        "exam_id": pick_digits(feat["exam_id"]),
        "features": feat,
        "aligned": aligned,
    }


if __name__ == "__main__":
    import sys
    r = read_sheet(sys.argv[1])
    print("subject:", r["subject_code"], " exam id:", r["exam_id"])
    for i in range(0, 170, 10):
        print(f"{i+1:>3}-{i+10:<3}", " ".join(r["answers"][i:i + 10]))
