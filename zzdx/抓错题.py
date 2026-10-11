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
            if (today - d).days > DAYS_BACK:
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
                       "person": real_name})
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
                    "persons": [], "answer": w["answer"]})
                if w["person"] not in rec["persons"]:
                    rec["persons"].append(w["person"])
    return list(agg.values())


# ---------- 飞书留底表 ----------

def norm(s):
    """内容归一化：逐行去首尾空白，忽略空行，容忍手打的前导句点（用于去重对比）"""
    s = "\n".join(x.strip() for x in (s or "").splitlines() if x.strip())
    return s.lstrip("。.")


def ident(r):
    return json.dumps([r["date"], r["examType"], r["cat"], norm(r["kp"]),
                       norm(r["content"]), sorted(r["persons"])], ensure_ascii=False)


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
        })
    return rows, ws.max_row


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


def append_rows(records, max_row):
    """追加写入飞书留底表「错题集」（v2 values_append）"""
    token, sid = get_sheet_id()
    values = [[r["date"], r["station"], r["examType"], r["cat"], r["kp"],
               r["content"], ",".join(r["persons"]), r["answer"]] for r in records]
    # range 需覆盖全部待写行
    rng = f"{sid}!A{max_row + 1}:H{max_row + len(values)}"
    r = ud._http_json(f"{ud.BASE}/open-apis/sheets/v2/spreadsheets/{ud.CUOTI_FILE_TOKEN}/values_append",
                      "POST", headers={"Authorization": f"Bearer {token}",
                                       "Content-Type": "application/json"},
                      body={"valueRange": {"range": rng, "values": values}})
    if r.get("code") != 0:
        raise SystemExit(f"写入留底表失败: {r.get('msg')}")
    print(f"  已写入「错题集」{len(values)} 行（追加自第 {max_row + 1} 行）")


# ---------- 主流程 ----------

def main():
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
    rows, max_row = sheet_rows()
    existing = {ident(r) for r in rows}
    new = [r for r in records if ident(r) not in existing]
    print(f"\n本次抓到错题 {len(records)} 条，去重后新增 {len(new)} 条：")
    for r in new:
        print(f"  + {r['date']} {r['cat']} · {r['kp']}（{','.join(r['persons'])}）")
    if not new:
        print("留底表已是最新，无需写入")
        return
    append_rows(new, max_row)
    print("运行看板更新与群推送 ...")
    subprocess.run([sys.executable, os.path.join(REPO, "update_dashboard.py")], cwd=REPO)


if __name__ == "__main__":
    main()
