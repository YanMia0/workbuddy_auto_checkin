#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wb_multi.py —— WorkBuddy 多账号每日自动签到 + 猫猫旅行 + 积分台账

复用 wb_auto.py 的签到/旅行/解包逻辑，扩展为多账号批量执行：

  * 账号配置：accounts.json（本地）或环境变量 WB_ACCOUNTS_JSON（GitHub Actions）
  * 幂等：每个账号先查状态再决定是否领取（沿用 wb_auto 的接口幂等设计）
  * 隔离：单个账号网络/认证失败不影响其它账号继续执行
  * 台账：每次执行每个账号写一行到 ledger/points_ledger.csv（utf-8-sig，Excel 直接打开）
  * 汇总：自动重算 ledger/points_summary.csv 的每账号累计积分与签到天数

用法：
  python3 wb_multi.py                 # 全部账号执行，写台账
  python3 wb_multi.py --alias 主号    # 只执行指定别名账号
  python3 wb_multi.py --dry           # 只跑不写台账
  python3 wb_multi.py --list          # 列出已配置账号别名

环境变量：
  WB_ACCOUNTS_JSON    JSON 字符串，形如 [{"alias":"主号","token":"eyJ..."}]
  WB_RESULT_FILE      非空时把 Markdown 汇总表写入该文件（供 Actions step summary）
  WB_ENDPOINT         默认 https://copilot.tencent.com
"""

import argparse
import csv
import datetime
import json
import os
import re
import sys

# 复用 wb_auto.py 的模块级函数（该文件 import 无副作用）
from wb_auto import (DEFAULT_ENDPOINT, dig, log, masked, run_buddy, run_checkin)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ACCOUNTS_FILE = os.path.join(BASE_DIR, "accounts.json")
LEDGER_DIR = os.path.join(BASE_DIR, "ledger")
LEDGER_CSV = os.path.join(LEDGER_DIR, "points_ledger.csv")
SUMMARY_CSV = os.path.join(LEDGER_DIR, "points_summary.csv")

LEDGER_HEADER = [
    "date", "alias", "checkin_result", "checkin_credit", "streak_days",
    "total_credits", "buddy_result", "buddy_credit", "note", "ok",
]


# ---------------------------------------------------------------- 配置

def load_accounts():
    """账号来源优先级：环境变量 WB_ACCOUNTS_JSON > 本地 accounts.json。"""
    env = os.environ.get("WB_ACCOUNTS_JSON")
    if env and env.strip():
        try:
            data = json.loads(env)
        except ValueError as e:
            log("[配置] WB_ACCOUNTS_JSON 不是合法 JSON：%s" % e)
            return []
    else:
        if not os.path.isfile(ACCOUNTS_FILE):
            log("[配置] 未找到 %s，请先运行：python3 wb_tokens.py add <别名> <token>"
                % ACCOUNTS_FILE)
            return []
        try:
            with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (ValueError, OSError) as e:
            log("[配置] 读取 accounts.json 失败：%s" % e)
            return []
    accounts = [a for a in data if isinstance(a, dict) and a.get("token")]
    if len(accounts) != len(data):
        log("[配置] 已忽略 %d 条缺少 token 的账号" % (len(data) - len(accounts)))
    return accounts


# ---------------------------------------------------------------- 指标解析

def _num(text, pattern):
    m = re.search(pattern, text or "")
    try:
        return int(m.group(1)) if m else None
    except (TypeError, ValueError):
        return None


def extract_checkin_metrics(checkin):
    """从签到结论里抽出积分/连签/累计。返回 (credit_today, streak, total)。"""
    report = checkin.get("report") or ""
    if checkin.get("result") == "CLAIMED":
        credit = checkin.get("credit")
        if credit is None:
            credit = _num(report, r"成功领取 (\d+) 积分")
        return credit, _num(report, r"连续 (\d+) 天"), _num(report, r"累计 (\d+) 积分")
    if checkin.get("result") == "ALREADY":
        return _num(report, r"今日 \+(\d+)"), _num(report, r"连续 (\d+) 天"), \
            _num(report, r"累计 (\d+) 积分")
    return None, _num(report, r"连续 (\d+) 天"), _num(report, r"累计 (\d+) 积分")


def extract_buddy_credit(buddy):
    return _num(buddy.get("report") or "", r"领旅行礼物 \+(\d+) 积分")


# ---------------------------------------------------------------- 台账

def _today():
    return datetime.date.today().isoformat()


def write_ledger_row(row):
    os.makedirs(LEDGER_DIR, exist_ok=True)
    new_file = not os.path.isfile(LEDGER_CSV)
    with open(LEDGER_CSV, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=LEDGER_HEADER, extrasaction="ignore")
        if new_file:
            w.writeheader()
        w.writerow(row)


def rebuild_summary():
    """扫描台账，按别名重算累计积分 / 签到天数 / 首次与最近日期。"""
    if not os.path.isfile(LEDGER_CSV):
        return
    agg = {}
    with open(LEDGER_CSV, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            alias = (r.get("alias") or "").strip()
            if not alias:
                continue
            a = agg.setdefault(alias, {"days": 0, "credit": 0,
                                       "first": r["date"], "last": r["date"]})
            a["days"] += 1
            if r.get("ok") == "1" or r.get("ok") == "True":
                c = r.get("checkin_credit")
                if c not in (None, ""):
                    try:
                        a["credit"] += int(float(c))
                    except ValueError:
                        pass
            a["first"] = min(a["first"], r["date"])
            a["last"] = max(a["last"], r["date"])
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with open(SUMMARY_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["alias", "checkin_days", "total_credits_gained",
                    "first_date", "last_date"])
        for alias in sorted(agg):
            a = agg[alias]
            w.writerow([alias, a["days"], a["credit"], a["first"], a["last"]])


# ---------------------------------------------------------------- 单账号执行

def run_account(account, endpoint, dry):
    alias = account.get("alias") or "unnamed"
    token = account["token"]
    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer %s" % token,
        "Content-Type": "application/json",
        "User-Agent": "WorkBuddy",
    }
    log("\n======== 账号[%s] ========" % alias)
    checkin_ok, checkin = run_checkin(headers, endpoint)
    log("  [签到] %s" % checkin["report"])
    try:
        buddy = run_buddy(headers, endpoint)
    except Exception as e:
        buddy = {"result": "ERROR", "report": "猫猫模块异常（%s: %s）"
                 % (type(e).__name__, e)}
    log("  [猫猫] %s" % buddy.get("report", "无结论"))

    c_credit, streak, total = extract_checkin_metrics(checkin)
    row = {
        "date": _today(),
        "alias": alias,
        "checkin_result": checkin.get("result", "UNKNOWN"),
        "checkin_credit": c_credit if c_credit is not None else "",
        "streak_days": streak if streak is not None else "",
        "total_credits": total if total is not None else "",
        "buddy_result": buddy.get("result", "UNKNOWN"),
        "buddy_credit": extract_buddy_credit(buddy) or "",
        "note": "签到：%s；猫猫：%s" % (checkin.get("report", ""),
                                      buddy.get("report", "")),
        "ok": "1" if checkin_ok else "0",
    }
    if not dry:
        write_ledger_row(row)
    return {"alias": alias, "checkin_ok": checkin_ok, "row": row}


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser(description="WorkBuddy 多账号签到 + 台账")
    ap.add_argument("--alias", help="只执行指定别名账号")
    ap.add_argument("--dry", action="store_true", help="只跑不写台账")
    ap.add_argument("--list", action="store_true", help="列出已配置账号")
    args = ap.parse_args()

    endpoint = os.environ.get("WB_ENDPOINT") or DEFAULT_ENDPOINT

    if args.list:
        for a in load_accounts():
            print(a.get("alias") or "unnamed")
        return 0

    accounts = load_accounts()
    if not accounts:
        log("没有可用账号。本地配置：python3 wb_tokens.py add <别名> <token>；"
            "或设置环境变量 WB_ACCOUNTS_JSON")
        return 2
    if args.alias:
        accounts = [a for a in accounts if (a.get("alias") or "unnamed") == args.alias]
        if not accounts:
            log("未找到别名 [%s] 的账号" % args.alias)
            return 2

    log("=== 开始执行 %d 个账号 ===" % len(accounts))
    results = []
    for account in accounts:
        try:
            results.append(run_account(account, endpoint, args.dry))
        except Exception as e:
            log("账号[%s]整体异常：%s: %s"
                % (account.get("alias", "unnamed"), type(e).__name__, e))
            results.append({"alias": account.get("alias", "unnamed"),
                            "checkin_ok": False})

    if not args.dry:
        rebuild_summary()

    all_ok = all(r.get("checkin_ok") for r in results)
    log("\n=== 汇总 ===")
    for r in results:
        log("  %s：%s" % (r["alias"], "成功" if r.get("checkin_ok") else "未达成"))

    summary_md = ["| 账号 | 签到 | 猫猫 | 今日积分 | 连续天数 |",
                  "| --- | --- | --- | --- | --- |"]
    for r in results:
        row = r.get("row") or {}
        summary_md.append("| %s | %s | %s | %s | %s |" % (
            row.get("alias", r.get("alias", "?")),
            row.get("checkin_result", "ERROR"),
            row.get("buddy_result", "-"),
            row.get("checkin_credit", "-"),
            row.get("streak_days", "-")))

    result_file = os.environ.get("WB_RESULT_FILE")
    if result_file:
        try:
            with open(result_file, "w", encoding="utf-8") as f:
                f.write("### Buddy 加油站 · 多账号每日自动化\n\n")
                f.write("- 运行时间：%s\n" % datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"))
                f.write("\n".join(summary_md))
                f.write("\n\n整体：%s\n" % ("全部成功" if all_ok else "存在未完成项"))
        except OSError as e:
            log("结果文件写入失败：%s" % e)

    print("RESULT " + json.dumps(
        {"accounts": [r.get("alias") for r in results],
         "all_ok": all_ok}, ensure_ascii=False))
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
