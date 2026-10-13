#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抓错题：转转学堂「武汉日考提升」错题 → 飞书留底表「错题集」 → 看板推送

用法: python3 抓错题.py
流程:
  1. 校验转转学堂登录态（失效时自动弹飞书扫码二维码，终端展示）
  2. 拉取最近 3 天武汉平板/笔记本日考提升培训的答题明细
  3. 按（日期+品类+知识点+题干）聚合错题，人员逗号连接
  4. 与留底表现有数据去重后，追加写入飞书留底表「错题集」
  5. 有新增时运行 update_dashboard.py，完成看板更新与群推送
"""
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import datetime

ZDIR = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(ZDIR)
sys.path.insert(0, REPO)
import update_dashboard as ud  # 复用飞书凭证/导出工具

COOKIE_FILE = os.path.join(ZDIR, "cookie.txt")
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36"
ZZ = "https://daxue.zhuanspirit.com/api"
STATION = "武汉分站"
EXAM_TYPE = "日考"
TITLE_RE = re.compile(r"武汉(平板|笔记本)日考提升(\d{2})(\d{2})")
DAYS_BACK = 3
SHEET_TITLE = "错题"  # 留底表中的「错题集」工作表


# ---------- 转转学堂 API ----------

def load_cookies():
    with open(COOKIE_FILE, encoding="utf-8") as f:
        return f.read().strip()


def zz_headers():
    return {"User-Agent": UA, "Cookie": load_cookies(),
            "Referer": "https://wxdaxue.zhuanspirit.com/"}


def zz_get(path, params=None):
    url = ZZ + path
    if params:
        url += "?" + "&".join(f"{k}={v}" for k, v in params.items())
    req = urllib.request.Request(url, headers=zz_headers())
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def zz_post(path, body):
    req = urllib.request.Request(ZZ + path, data=json.dumps(body).encode(),
                                 headers={**zz_headers(), "Content-Type": "application/json"},
                                 method="POST")
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def session_ok():
    try:
        return zz_get("/common/getStudyCountToday").get("code") == 0
    except Exception:
        return False


def qr_login():
    print("登录态已失效，开始飞书扫码登录 ...")
    r = subprocess.run(["node", os.path.join(ZDIR, "qr_login.js")], cwd=ZDIR)
    if r.returncode != 0 or not os.path.isfile(COOKIE_FILE):
        raise SystemExit("扫码登录未完成，本次终止（稍后再说一声「抓错题」重试）")


def fetch_trainings():
    """最近 DAYS_BACK 天的武汉日考提升培训（平板+笔记本）"""
    # 分品类搜索再按标题正则过滤（其他站点的推送会混在结果里，需过滤）
    today = datetime.now()
    out = []
    for kw in ("平板日考提升", "笔记本日考提升"):
        r = zz_post("/training/list", {"pageNum": 1, "pageSize": 10, "trainingTitle": kw})
        if r.get("code") != 0:
            raise SystemExit(f"培训列表获取失败: {r.get('message')}")
        for t in r["data"]["list"]:
            m = TITLE_RE.match(t["trainingTitle"])
            if not m:
                continue
            cat, mm, dd = m.group(1), int(m.group(2)), int(m.group(3))
            push = datetime.strptime(t["pushTime"], "%Y-%m-%d %H:%M:%S")
            d = datetime(push.year, mm, dd)
            if (today - d).days > DAYS_BACK or d.date() > today.date():  # 排除未来日期
                continue
            out.append({"id": t["id"], "title": t["trainingTitle"], "cat": cat,
                        "date": d.strftime("%Y-%m-%d")})
    out.sort(key=lambda x: x["date"])
    return out


def fetch_exam_id(tid):
    r = zz_post("/training/items", {"trainingId": tid, "type": 2})
    items = r.get("data") or []
    return items[0]["itemId"] if items else None


def fetch_examinees(tid):
    r = zz_post("/training/finishInfo",
                {"trainingId": tid, "learningStatus": 2, "pageNum": 1, "pageSize": 100})
    if r.get("code") != 0:
        return []
    return r["data"].get("list") or []


def clean_html(s):
    """剥 HTML 标签（<br> 换行）并反转义（部分题目的解析/选项是富文本）"""
    import html
    s = re.sub(r"<(br|/p|p)[^>]*>", "\n", s or "")
    s = re.sub(r"<[^>]+>", "", s)
    return html.unescape(s)


def fetch_wrong(tid, exam_id, uid, push_record_id, real_name):
    """单个考生的错题详情（跳过未作答）"""
    r = zz_get("/training/answerDetail",
               {"trainingId": tid, "examId": exam_id, "uid": uid,
                "pushRecordId": push_record_id})
    if r.get("code") != 0 or not isinstance(r.get("data"), dict):
        return []
    wrongs = []
    for sh in r["data"].get("answerSheets") or []:
        if sh.get("isCorrect") != 0:
            continue
        it = zz_post("/training/answerDetail/item", {"id": sh["id"]})
        d = it.get("data")
        if not isinstance(d, dict):
            continue
        if not (d.get("userChoices") or "").strip():  # 未作答不计入错题
            continue
        opts = d.get("options") or []
        n_ans = sum(1 for o in opts if o.get("isAnswer") == 1)
        qtype = "多选题" if n_ans > 1 else "单选题"
        lines = [f"【{qtype}】", clean_html(d.get("questionNameText")).strip()]
        lines += [f"{clean_html(o.get('prefix', ''))}.{clean_html(o.get('option', ''))}" for o in opts]
        kp = ""
        m = re.search(r"考点分类[:：]\s*([^\n]+?)(?=\s*答案讲解|\s*相关标准|$)",
                      clean_html(d.get("analysis")))
        if m:
            kp = m.group(1).strip()
        wrongs.append({"kp": kp, "content": "\n".join(lines),
                       "answer": f"答案：{(d.get('answer') or '').strip()}",
                       "person": real_name,
                       "media_url": (d.get("mediaUrls") or [None])[0],
                       "analysis_urls": d.get("analysisUrls") or [],
                       "snap_data": {
                           "qtype": qtype,
                           "qtext": clean_html(d.get("questionNameText")).strip(),
                           "options": [{"prefix": clean_html(o.get("prefix", "")),
                                        "text": clean_html(o.get("option", "")),
                                        "right": o.get("isAnswer") == 1} for o in opts],
                           "user": [x.strip() for x in re.split(r"[、，,]", d.get("userChoices") or "")
                                    if x.strip()],
                           "answer": (d.get("answer") or "").strip(),
                           "kp": kp, "analysis": clean_html(d.get("analysis")),
                           "media": d.get("mediaUrls") or [],
                           "analysisImgs": d.get("analysisUrls") or []}})
    return wrongs


def build_records(trainings):
    """汇总为留底表行：同日期+品类+知识点+题干合并，人员逗号连接"""
    agg = {}
    for t in trainings:
        exam_id = fetch_exam_id(t["id"])
        if not exam_id:
            print(f"  ⚠ {t['title']}: 未找到考试，跳过")
            continue
        people = fetch_examinees(t["id"])
        print(f"  {t['title']}（{t['date']}）: {len(people)} 人")
        for p in people:
            try:
                wrongs = fetch_wrong(t["id"], exam_id, p.get("uid"),
                                     p.get("pushRecordId"), p.get("realName"))
            except Exception as e:
                print(f"    ⚠ {p.get('realName')} 获取失败: {e}")
                continue
            for w in wrongs:
                key = (t["date"], t["cat"], w["kp"], w["content"])
                rec = agg.setdefault(key, {
                    "date": t["date"], "station": STATION, "examType": EXAM_TYPE,
                    "cat": t["cat"], "kp": w["kp"], "content": w["content"],
                    "persons": [], "answer": w["answer"],
                    "media_url": w["media_url"], "analysis_urls": w["analysis_urls"],
                    "snap_data": w["snap_data"]})
                if w["person"] not in rec["persons"]:
                    rec["persons"].append(w["person"])
    for r in agg.values():
        r["snap_data"].update({"date": r["date"], "cat": r["cat"]})
    return list(agg.values())


# ---------- 飞书留底表 ----------

def norm(s):
    """内容归一化（用于去重对比）：逐行去空白、去首尾标点，容忍手打的类型标签差异
    （漏打【、重复标签行、前导句点），合并连续重复行"""
    out = []
    for x in (s or "").splitlines():
        x = x.strip().strip("。，、；,;．.").lstrip("。，、；,;．.")
        if not x:
            continue
        m = re.fullmatch(r"【?[单多]选题】?", x)
        if m:
            x = "多选题】" if "多" in x else "单选题】"
        if out and out[-1] == x:
            continue
        out.append(x)
    return "\n".join(out)


def ident(r):
    return json.dumps([r["date"], r["examType"], r["cat"], norm(r["kp"]),
                       norm(r["content"]), sorted(r["persons"])], ensure_ascii=False)


def key_no_persons(r):
    """不含人员的去重键（用于同题人员补全合并）"""
    return json.dumps([r["date"], r["examType"], r["cat"], norm(r["kp"]),
                       norm(r["content"])], ensure_ascii=False)


def sheet_rows():
    """下载留底表并解析「错题集」现有行"""
    import openpyxl
    print("下载留底表 ...")
    data = ud.fetch_xlsx(ud.CUOTI_FILE_TOKEN)
    tmp = os.path.join(ZDIR, "_tmp_cuoti.xlsx")
    with open(tmp, "wb") as f:
        f.write(data)
    try:
        wb = openpyxl.load_workbook(tmp, data_only=True)
    finally:
        os.remove(tmp)
    ws = ud.find_sheet(wb, SHEET_TITLE)
    hdr = ud.find_header_row(ws, ["错题知识点", "错题内容"])
    c = {n: ud.col_of(ws, hdr, n) for n in
         ["日期", "站点", "考试类型", "品类", "错题知识点", "错题内容", "质检人员", "正确答案/解析"]}
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        dt = ws.cell(row=r, column=c["日期"]).value
        if not dt:
            continue
        rows.append({
            "date": ud.sval(dt),
            "station": ud.sval(ws.cell(row=r, column=c["站点"]).value),
            "examType": ud.sval(ws.cell(row=r, column=c["考试类型"]).value),
            "cat": ud.sval(ws.cell(row=r, column=c["品类"]).value),
            "kp": ud.sval(ws.cell(row=r, column=c["错题知识点"]).value),
            "content": ud.sval(ws.cell(row=r, column=c["错题内容"]).value),
            "persons": [p for p in re.split(r"[,，、;；]", ud.sval(ws.cell(row=r, column=c["质检人员"]).value)) if p],
            "answer": ud.sval(ws.cell(row=r, column=c["正确答案/解析"]).value),
            "_row": r,
        })
    # 单元格图片（人工贴图在 I 列起）：行号 → [(列0基, 图片对象)]
    img_map = {}
    for img in getattr(ws, "_images", []):
        a = img.anchor
        row = (a._from.row if hasattr(a, "_from") else a.from_.row) + 1
        col = a._from.col if hasattr(a, "_from") else a.from_.col
        if col >= 8:
            img_map.setdefault(row, []).append((col, img))
    # 最后一行真实数据的行号（追加从这里 +1 开始；末尾可能有空行/占位行）
    last_data = max((r["_row"] for r in rows), default=hdr)
    return rows, last_data, img_map


def get_sheet_id():
    """查询「错题集」工作表的 sheet_id"""
    app_id, app_secret = ud.get_app_credentials()
    t = ud._http_json(f"{ud.BASE}/open-apis/auth/v3/tenant_access_token/internal", "POST",
                      body={"app_id": app_id, "app_secret": app_secret})["tenant_access_token"]
    h = {"Authorization": f"Bearer {t}"}
    r = ud._http_json(f"{ud.BASE}/open-apis/sheets/v3/spreadsheets/{ud.CUOTI_FILE_TOKEN}/sheets/query",
                      headers=h)
    if r.get("code") != 0:
        raise SystemExit(f"查询工作表失败: {r.get('msg')}")
    for s in r["data"]["sheets"]:
        if s.get("title") == "错题集":
            return t, s["sheet_id"]
    raise SystemExit("未找到「错题集」工作表")


def append_rows(records, start_row):
    """追加写入飞书留底表「错题集」（v2 values_append）"""
    token, sid = get_sheet_id()
    values = [[r["date"], r["station"], r["examType"], r["cat"], r["kp"],
               r["content"], ",".join(r["persons"]), r["answer"]] for r in records]
    rng = f"{sid}!A{start_row}:H{start_row + len(values) - 1}"
    r = ud._http_json(f"{ud.BASE}/open-apis/sheets/v2/spreadsheets/{ud.CUOTI_FILE_TOKEN}/values_append",
                      "POST", headers={"Authorization": f"Bearer {token}",
                                       "Content-Type": "application/json"},
                      body={"valueRange": {"range": rng, "values": values}})
    if r.get("code") != 0:
        raise SystemExit(f"写入留底表失败: {r.get('msg')}")
    print(f"  已写入「错题集」{len(values)} 行（自第 {start_row} 行起）")


def _to_cell_jpeg(raw):
    """图片转适合单元格的 JPEG（宽超 1600 压缩），转换失败保留原图"""
    tmp = os.path.join(ZDIR, "_tmp_img_conv")
    with open(tmp, "wb") as f:
        f.write(raw)
    try:
        r = subprocess.run(["sips", "-s", "format", "jpeg", "-s", "formatOptions",
                            "80", "--resampleWidth", "1600", tmp, "--out", tmp + ".jpg"],
                           capture_output=True)
        data = open(tmp + ".jpg", "rb").read() if r.returncode == 0 else None
    except OSError:
        data = None
    if not data:
        try:
            from PIL import Image
            im = Image.open(tmp)
            if im.width > 1600:
                im = im.resize((1600, round(im.height * 1600 / im.width)))
            im.convert("RGB").save(tmp + ".jpg", "JPEG", quality=80)
            data = open(tmp + ".jpg", "rb").read()
        except Exception:
            data = raw
    for p in (tmp, tmp + ".jpg"):
        if os.path.isfile(p):
            os.remove(p)
    return data


def _img_same(a_bytes, b_bytes):
    """感知比较：缩到 64x64 后平均像素差 <8 视为同一张图（容忍飞书转码压缩）"""
    try:
        import io
        import numpy as np
        from PIL import Image

        def arr(b):
            im = Image.open(io.BytesIO(b)).convert("RGB").resize((64, 64))
            return np.asarray(im, dtype=np.int32)
        return np.abs(arr(a_bytes) - arr(b_bytes)).mean() < 8
    except Exception:
        return False


def embed_images(jobs):
    """把题目图片/全解释截图嵌入单元格（I 列起，与人工贴图同位置；解析文字在 H 列不受影响）
    jobs: [(行号, [(列0基, 类型'url'|'snap', 值), ...]), ...]"""
    import base64
    from openpyxl.utils import get_column_letter
    if not jobs:
        return
    token, sid = get_sheet_id()
    done = 0
    for row, items in jobs:
        for i, (col, kind, val) in enumerate(items):
            if kind == "url":
                try:
                    req = urllib.request.Request(val, headers={"User-Agent": UA})
                    raw = urllib.request.urlopen(req, timeout=60).read()
                except Exception as e:
                    print(f"  ⚠ 第{row}行 图片下载失败: {val} ({e})")
                    continue
            elif kind == "render":
                data_json = os.path.join(ZDIR, f"_tmp_snap_{row}.json")
                with open(data_json, "w", encoding="utf-8") as f:
                    json.dump(val, f, ensure_ascii=False)
                out = os.path.join(ZDIR, f"_tmp_snap_{row}")
                r = subprocess.run(["node", os.path.join(ZDIR, "render_snap.js"),
                                    data_json, out], cwd=ZDIR, capture_output=True)
                if os.path.isfile(data_json):
                    os.remove(data_json)
                if r.returncode != 0 or not os.path.isfile(out + ".png"):
                    print(f"  ⚠ 第{row}行 全解释图渲染失败（{(r.stdout or b'').decode().strip() or r.returncode}）")
                    if os.path.isfile(out + ".png"):
                        os.remove(out + ".png")
                    continue
                with open(out + ".png", "rb") as f:
                    raw = f.read()
                os.remove(out + ".png")
            else:
                continue
            data = _to_cell_jpeg(raw)
            # 过长截图拆分多张（单张不超 2600px，保证群聊里能看清字；尾部余片并入前一张）
            parts = []
            try:
                import io as _io
                from PIL import Image
                im = Image.open(_io.BytesIO(data))
                if im.height > 2900:
                    n = (im.height + 2599) // 2600
                    if im.height - (n - 1) * 2600 < 400:
                        n -= 1
                    for k in range(n):
                        buf = _io.BytesIO()
                        im.crop((0, k * 2600, im.width, min(im.height, (k + 1) * 2600))).save(buf, "JPEG", quality=80)
                        parts.append(buf.getvalue())
                else:
                    parts = [data]
            except Exception:
                parts = [data]
            for k, part in enumerate(parts):
                rng = f"{sid}!{get_column_letter(col + 1 + k)}{row}:{get_column_letter(col + 1 + k)}{row}"
                r = ud._http_json(f"{ud.BASE}/open-apis/sheets/v2/spreadsheets/"
                                  f"{ud.CUOTI_FILE_TOKEN}/values_image",
                                  "POST", headers={"Authorization": f"Bearer {token}",
                                                   "Content-Type": "application/json; charset=utf-8"},
                                  body={"range": rng, "image": base64.b64encode(part).decode(),
                                        "name": f"cuoti_{row}_{i + 1}_{k + 1}.jpg"})
                if r.get("code") != 0:
                    print(f"  ⚠ 第{row}行 图片嵌入失败: {r.get('msg')}")
                else:
                    done += 1
    if done:
        print(f"  已嵌入 {done} 张图片")


def merge_persons(existing_rows, records):
    """同题已存在但人员不全时，返回需更新的（行号, 合并后人员串）；不动答案与内容。
    所有同题行都更新到人员并集（表里同一道题可能被手动记了多行）"""
    by_key = {}
    for r in existing_rows:
        by_key.setdefault(key_no_persons(r), []).append(r)
    updates = []
    for rec in records:
        for e in by_key.get(key_no_persons(rec), []):
            missing = [p for p in rec["persons"] if p not in e["persons"]]
            if not missing:
                continue
            merged = e["persons"] + missing
            updates.append((e["_row"], ",".join(merged)))
    return updates


def update_person_cells(updates):
    """批量更新「质检人员」列（G 列）"""
    token, sid = get_sheet_id()
    body = {"valueRanges": [{"range": f"{sid}!G{row}:G{row}", "values": [[v]]}
                            for row, v in updates]}
    r = ud._http_json(f"{ud.BASE}/open-apis/sheets/v2/spreadsheets/{ud.CUOTI_FILE_TOKEN}/values_batch_update",
                      "POST", headers={"Authorization": f"Bearer {token}",
                                       "Content-Type": "application/json"},
                      body=body)
    if r.get("code") != 0:
        print(f"  警告: 人员合并更新失败: {r.get('msg')}")
        return
    print(f"  已合并更新 {len(updates)} 行质检人员")


# ---------- 主流程 ----------

def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--push", action="store_true", help="写完留底表后运行看板更新与群播报（默认关闭）")
    args = ap.parse_args()

    if not os.path.isfile(COOKIE_FILE) or not session_ok():
        qr_login()
    print("拉取培训列表 ...")
    trainings = fetch_trainings()
    if not trainings:
        print(f"近 {DAYS_BACK} 天没有武汉日考提升培训，本次无需处理")
        return
    print("抓取答题明细 ...")
    records = build_records(trainings)
    if not records:
        print("没有抓到错题数据（可能都未完成考试），本次结束")
        return
    rows, last_data, img_map = sheet_rows()
    existing = {ident(r) for r in rows}
    # 同题已存在（无论人员是否齐全）→ 走人员合并，不再新增行
    existing_no_p = {key_no_persons(r) for r in rows}
    new = [r for r in records if key_no_persons(r) not in existing_no_p]
    merges = merge_persons(rows, records)
    # 图片任务：新行贴全套（题干图@I + 全解释截图@J，无题干图则截图@I）；
    # 已有行缺图补齐；两张齐全但其中一张是平台解析配图 → 换成全解释截图
    def jobs_for(rec):
        items = []
        col = 8
        if rec.get("media_url"):
            items.append((col, "url", rec["media_url"]))
            col += 1
        if rec.get("snap_data"):
            items.append((col, "render", rec["snap_data"]))
        return items

    row_by_key = {}
    for r in rows:
        row_by_key.setdefault(key_no_persons(r), []).append(r["_row"])
    img_targets = []
    for i, r in enumerate(new):
        jobs = jobs_for(r)
        if jobs:
            img_targets.append((last_data + 1 + i, jobs))
    for rec in records:
        need = jobs_for(rec)
        if not need:
            continue
        for row in row_by_key.get(key_no_persons(rec), []):
            occ = sorted(col for col, _ in img_map.get(row, []))
            if len(occ) < len(need):
                jobs = [(col, kind, val) for col, kind, val in need if col not in occ]
                if jobs:
                    img_targets.append((row, jobs))
            elif len(occ) == len(need) and rec.get("snap_data") and rec.get("analysis_urls"):
                # 两张图齐全但其中一张是解析配图 → 换成全解释图
                bad = None
                for u in rec["analysis_urls"]:
                    try:
                        req = urllib.request.Request(u, headers={"User-Agent": UA})
                        raw = urllib.request.urlopen(req, timeout=30).read()
                    except Exception:
                        continue
                    for col, img in img_map.get(row, []):
                        if _img_same(img._data(), raw):
                            bad = col
                            break
                    if bad is not None:
                        break
                if bad is not None:
                    img_targets.append((row, [(bad, "render", rec["snap_data"])]))
    print(f"\n本次抓到错题 {len(records)} 条，去重后新增 {len(new)} 条，人员补全 {len(merges)} 行，"
          f"图片任务 {len(img_targets)} 行：")
    for r in new:
        print(f"  + {r['date']} {r['cat']} · {r['kp']}（{','.join(r['persons'])}）")
    for row, v in merges:
        print(f"  ~ 第{row}行 人员 → {v}")
    for row, jobs in img_targets:
        desc = "、".join("题干图" if kind == "url" else "全解释图" for _, kind, _ in jobs)
        print(f"  🖼 第{row}行 {desc}")
    if not new and not merges and not img_targets:
        print("留底表已是最新，无需写入")
        return
    if merges:
        update_person_cells(merges)
    if new:
        append_rows(new, last_data + 1)
    if img_targets:
        embed_images(img_targets)
    if args.push:
        print("运行看板更新与群播报 ...")
        subprocess.run([sys.executable, os.path.join(REPO, "update_dashboard.py")], cwd=REPO)
    else:
        print("已跳过看板更新与群播报（需要推送时加 --push，或自行运行 update_dashboard.py）")


if __name__ == "__main__":
    main()
