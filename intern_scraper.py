#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
互联网大厂 2027届实习岗位信息爬虫
每天自动抓取字节、腾讯、阿里、百度、美团、快手、拼多多、小红书 的实习岗位信息
输出格式：Markdown 表格

使用方法：
    python intern_scraper.py                     # 默认输出到控制台
    python intern_scraper.py --output report.md  # 输出到文件
    python intern_scraper.py --send-email        # 输出并发送邮件
"""

import sys
import os
import re
import json
import argparse
from datetime import datetime
from dataclasses import dataclass, field, asdict
from typing import Optional

# ── 确保输出编码为 UTF-8 ──
if sys.stdout.encoding != 'utf-8':
    sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

# ─── 数据模型 ───────────────────────────────────────────────

@dataclass
class InternPosition:
    company: str          # 公司名称
    title: str            # 岗位名称
    location: str         = "未标注"  # 工作地点
    requirement: str      = "待查看"  # 岗位要求（精简版）
    salary: str           = "未公开"  # 薪资范围
    program: str          = ""        # 所属项目（如 Ace精英实习生计划）
    link: str             = ""        # 详情链接
    source_note: str      = ""        # 数据来源备注


# ─── 各公司爬虫 ─────────────────────────────────────────────

EDGE_PATH = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

def _init_browser():
    """初始化 Playwright 浏览器实例（跨平台：Windows 用 Edge，Linux 用 Chromium）"""
    from playwright.sync_api import sync_playwright
    p = sync_playwright().start()
    
    import sys
    if sys.platform == "win32":
        # Windows: 使用已安装的 Edge
        browser = p.chromium.launch(executable_path=EDGE_PATH, headless=True)
    else:
        # Linux (GitHub Actions): 使用 Playwright 自带的 Chromium
        browser = p.chromium.launch(headless=True)
    return p, browser


def _safe_text(page, selector, default=""):
    """安全获取元素文本"""
    try:
        el = page.query_selector(selector)
        return el.inner_text().strip() if el else default
    except Exception:
        return default


def scrape_xiaohongshu() -> list[InternPosition]:
    """小红书 - 实习岗位"""
    results = []
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://job.xiaohongshu.com/campus/position", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(4000)

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        # 解析岗位信息（小红书页面格式：岗位名 → 类别 → 地点 → 项目名 → 描述）
        i = 0
        seen_titles = set()
        while i < len(lines):
            line = lines[i]
            # 跳过导航和筛选栏
            if line in ["首页", "REDstar 顶尖校招", "Ace 精英实习生", "职位", "校招须知",
                         "社会招聘", "了解小红书", "登录", "注册", "筛选", "清空",
                         "招聘类别", "应届生", "实习生", "Ace精英实习生",
                         "招聘项目", "REDstar 顶尖人才计划", "2026 校园招聘",
                         "所有实习岗位", "职位类别", "技术类", "产品类", "运营类",
                         "设计类", "职能类", "市场类", "工作地点",
                         "北京市", "上海市", "广州市", "深圳市", "武汉市", "杭州市",
                         "南京市", "珠海市", "成都市", "其他",
                         "Ace 顶尖实习生计划", "Ace顶尖实习生计划"]:
                i += 1
                continue
            if line.startswith("全部职位") or line.startswith("共") or "个职位" in line:
                i += 1
                continue
            # 跳过非岗位行
            if line in ["Ace 顶尖实习生", "Ace顶尖实习生", "Ace顶尖实习生计划"]:
                i += 1
                continue
            # 检查是否是真实职位条目（以【开头或包含明确的技术方向）
            is_position = False
            if line.startswith("【") and ("实习生" in line or "实习" in line):
                is_position = True
            elif ("实习生" in line or "实习" in line) and len(line) > 10:
                is_position = True

            if is_position:
                title = line
                location = ""
                req = ""
                program = "Ace精英实习生计划"

                # 看后面几行获取更多信息
                j = i + 1
                while j < min(i + 6, len(lines)):
                    nxt = lines[j]
                    if "北京" in nxt or "上海" in nxt or "深圳" in nxt or "杭州" in nxt or \
                         "广州" in nxt or "武汉" in nxt or "南京" in nxt or "珠海" in nxt or \
                         "成都" in nxt:
                        location = nxt
                    elif len(nxt) > 15 and ("研究" in nxt or "开发" in nxt or "优化" in nxt or
                                           "系统" in nxt or "算法" in nxt or "模型" in nxt or
                                           "工程" in nxt or "平台" in nxt):
                        req = nxt[:100] + ("..." if len(nxt) > 100 else "")
                        break
                    j += 1

                # 去重
                clean_title = re.sub(r'[【】]', '', title)[:25]
                if clean_title not in seen_titles:
                    seen_titles.add(clean_title)
                    results.append(InternPosition(
                        company="小红书",
                        title=title,
                        location=location or "北京/上海/杭州",
                        requirement=req or "详见官网",
                        program=program,
                        link="https://job.xiaohongshu.com/campus/position",
                    ))
            i += 1

        # 如果解析结果太少，改用完整文本匹配
        if len(results) < 3:
            results.clear()
            # 用正则匹配岗位模式：中文+实习生/实习
            pattern = r'([\u4e00-\u9fff]{2,30}(?:实习生|实习)[\u4e00-\u9fff]{0,20})'
            matches = re.findall(pattern, text)
            seen = set()
            for m in matches:
                if m not in seen and len(m) > 4:
                    seen.add(m)
                    results.append(InternPosition(
                        company="小红书",
                        title=m,
                        location="北京/上海/杭州等",
                        link="https://job.xiaohongshu.com/campus/position",
                    ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="小红书", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_tencent() -> list[InternPosition]:
    """腾讯 - 实习项目"""
    results = []
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://join.qq.com/", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(3000)

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        # 提取招聘项目信息（过滤非岗位行）
        projects = []
        skip_tencent = ["毕业时间", "工作地点在中国", "首页", "登录", "了解腾讯"]
        for i, line in enumerate(lines):
            if any(s in line for s in skip_tencent):
                continue
            if "实习" in line and "招聘" in line:
                projects.append(line)

        for proj in projects:
            results.append(InternPosition(
                company="腾讯",
                title=proj,
                location="深圳/北京/上海/广州/成都",
                requirement="详见腾讯校招官网",
                link="https://join.qq.com/",
                source_note="招聘项目信息（具体岗位需在官网查看）"
            ))

        # 检查是否有"技术大咖-实习生"等项目
        for kw in ["技术大咖", "Pre留学生", "产品", "运营", "设计"]:
            for line in lines:
                if kw in line and ("实习" in line or "培训" in line or "计划" in line):
                    if not any(kw in r.title for r in results):
                        results.append(InternPosition(
                            company="腾讯",
                            title=f"{kw}相关实习项目",
                            location="深圳/北京/上海/广州/成都",
                            link="https://join.qq.com/",
                        ))
                    break

        if not results:
            results.append(InternPosition(
                company="腾讯",
                title="腾讯2026实习招聘（面向2027届）",
                location="深圳/北京/上海/广州/成都",
                link="https://join.qq.com/",
                source_note="招聘项目已开放，具体岗位请登录官网查看"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="腾讯", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_pinduoduo() -> list[InternPosition]:
    """拼多多 - 2027届校招/实习"""
    results = []
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://careers.pinduoduo.com/campus", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(5000)

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        # 提取招聘项目信息
        for line in lines:
            if "2027" in line and ("校招" in line or "实习" in line or "招聘" in line or "提前批" in line):
                results.append(InternPosition(
                    company="拼多多",
                    title=line,
                    location="上海",
                    link="https://careers.pinduoduo.com/campus",
                    source_note="2027届校招信息"
                ))

        # 找岗位卡片（员工分享卡片中的岗位信息）
        for line in lines:
            # 匹配 "岗位名 + 年份 + 学校" 模式
            m = re.match(r'^([\u4e00-\u9fff\u4e00-\u9fff·]+)\s+\d{4}年\s+.+', line)
            if m:
                pos_name = m.group(1).strip()
                if pos_name and len(pos_name) > 2:
                    results.append(InternPosition(
                        company="拼多多",
                        title=f"{pos_name}（往届参考）",
                        location="上海",
                        link="https://careers.pinduoduo.com/campus",
                        source_note="往届岗位参考"
                    ))

        if not results:
            results.append(InternPosition(
                company="拼多多",
                title="2027届校招提前批/实习项目已启动",
                location="上海",
                link="https://careers.pinduoduo.com/campus",
                source_note="具体岗位请在官网查看"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="拼多多", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_bytedance() -> list[InternPosition]:
    """字节跳动 - 实习岗位"""
    results = []
    seen_titles = set()
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://jobs.bytedance.com/campus/", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(3000)

        # 点击"职位"链接进入岗位列表页
        try:
            pos_link = page.query_selector('a[href="/campus/position"]')
            if pos_link:
                with page.expect_navigation(wait_until="domcontentloaded", timeout=15000):
                    pos_link.click()
                page.wait_for_timeout(5000)
        except Exception:
            # 如果点击失败，尝试直接导航
            try:
                page.goto("https://jobs.bytedance.com/campus/position", wait_until="domcontentloaded", timeout=15000)
                page.wait_for_timeout(5000)
            except Exception:
                pass

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        # 跳过过滤词
        skip_words = ["面向全体在校生", "为符合岗位要求的同学提供", "日常实习：", "职位 ID：",
                       "招聘项目", "项目说明", "技术人才项目", "一键投递",
                       "字节范", "字节跳动的使命", "我们因使命",
                       "1、结合实际", "2、配合产品", "3、协助管理", "4、对语音",
                       "团队介绍", "筛选", "清除", "职位类别", "工作地点", "更多",
                       "开启新的工作"]

        for line in lines:
            # 只保留包含"实习"且看起来像岗位名称的行
            if "实习" not in line:
                continue
            # 跳过说明文字
            if any(skip in line for skip in skip_words):
                continue
            # 跳过过短或过长的行
            if len(line) < 6 or len(line) > 80:
                continue
            # 跳过纯描述
            if line.startswith("日常实习") and "：" in line:
                continue

            # 提取城市信息
            location = "北京/上海/深圳/杭州/广州等"
            for city in ["北京", "上海", "深圳", "杭州", "广州", "成都"]:
                if f"{city}实习" in line or line.startswith(city):
                    location = city
                    break

            # 清理标题
            title = re.sub(r'\s*职位 ID：[A-Z0-9]+', '', line).strip()

            # 去重
            key = title[:20]
            if key not in seen_titles and len(title) > 4:
                seen_titles.add(key)
                results.append(InternPosition(
                    company="字节跳动",
                    title=title,
                    location=location,
                    link="https://jobs.bytedance.com/campus/",
                ))

        if not results:
            results.append(InternPosition(
                company="字节跳动",
                title="校园招聘/实习项目已开放",
                location="北京/上海/深圳/杭州/广州等",
                link="https://jobs.bytedance.com/campus/",
                source_note="请在官网注册后查看具体岗位"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="字节跳动", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载超时"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_meituan() -> list[InternPosition]:
    """美团 - 校园招聘/实习"""
    results = []
    seen_titles = set()
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://zhaopin.meituan.com/campus", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(3000)

        # 尝试点击"校园招聘"标签
        try:
            campus_tab = page.query_selector("text=校园招聘")
            if campus_tab:
                campus_tab.click()
                page.wait_for_timeout(4000)
        except Exception:
            pass

        # 尝试选择"实习"筛选
        try:
            intern_filter = page.query_selector("text=实习")
            if intern_filter:
                intern_filter.click()
                page.wait_for_timeout(3000)
        except Exception:
            pass

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        # 过滤词
        skip_words = [
            "首页", "LongCat人才招聘", "北斗计划", "社会招聘", "了解美团",
            "Global Careers", "登录", "筛选", "清除所有", "所在城市", "所在海外国家/地区",
            "搜索更多城市", "职位类别", "部门", "岗位职责", "美团有多少年",
            "更多疑问请查看", "应届校招", "全部校招职位", "日常实习",
            "1.负责", "2.负责", "3.负责", "4.负责",
        ]

        for i, line in enumerate(lines):
            # 只保留包含"实习"的行
            if "实习" not in line:
                continue
            # 跳过过滤词
            if any(skip in line for skip in skip_words):
                continue
            # 跳过过长或过短的行
            if len(line) < 6 or len(line) > 60:
                continue
            # 跳过纯数字职位ID
            if re.match(r'^\d+$', line):
                continue

            # 提取地点
            location = "北京/上海/深圳/成都等"
            # 看下一行是否是地点
            if i + 1 < len(lines):
                nxt = lines[i + 1]
                cities = []
                for city in ["北京", "上海", "深圳", "成都", "广州", "杭州", "武汉", "南京", "厦门", "西安", "重庆"]:
                    if city in nxt:
                        cities.append(city)
                if cities:
                    location = "/".join(cities[:3])

            # 去重
            key = line[:20]
            if key not in seen_titles:
                seen_titles.add(key)
                results.append(InternPosition(
                    company="美团",
                    title=line,
                    location=location,
                    link="https://zhaopin.meituan.com/campus",
                ))

        if not results:
            results.append(InternPosition(
                company="美团",
                title="2026春招/实习招聘进行中（转正实习、日常实习）",
                location="北京/上海/深圳/成都等",
                link="https://zhaopin.meituan.com/campus",
                source_note="请在官网筛选查看具体岗位"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="美团", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_alibaba() -> list[InternPosition]:
    """阿里巴巴 - 实习信息"""
    results = []
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://talent.alibaba.com/campus/home", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(5000)

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        for line in lines:
            if "实习" in line and len(line) > 2:
                results.append(InternPosition(
                    company="阿里巴巴",
                    title=line,
                    location="杭州/北京/上海/深圳/广州",
                    link="https://talent.alibaba.com/campus/home",
                    source_note="请在官网登录后查看"
                ))

        if not results:
            results.append(InternPosition(
                company="阿里巴巴",
                title="校园招聘进行中",
                location="杭州/北京/上海/深圳/广州",
                link="https://talent.alibaba.com/campus/home",
                source_note="具体岗位请登录官网查看"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="阿里巴巴", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_baidu() -> list[InternPosition]:
    """百度 - 实习信息"""
    results = []
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://talent.baidu.com/campus/", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(3000)

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        for line in lines:
            if "实习" in line and len(line) > 2:
                results.append(InternPosition(
                    company="百度",
                    title=line,
                    location="北京/上海/深圳",
                    link="https://talent.baidu.com/campus/",
                ))

        if not results:
            results.append(InternPosition(
                company="百度",
                title="校园招聘进行中（需登录查看）",
                location="北京/上海/深圳",
                link="https://talent.baidu.com/campus/",
                source_note="需登录百度账号查看具体岗位"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="百度", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


def scrape_kuaishou() -> list[InternPosition]:
    """快手 - 实习信息"""
    results = []
    p, browser = _init_browser()
    try:
        page = browser.new_page()
        page.goto("https://zhaopin.kuaishou.cn/#/campus", wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(5000)

        text = page.inner_text("body")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        for line in lines:
            if "实习" in line and len(line) > 2:
                results.append(InternPosition(
                    company="快手",
                    title=line,
                    location="北京/深圳",
                    link="https://zhaopin.kuaishou.cn/#/campus",
                ))

        if not results:
            results.append(InternPosition(
                company="快手",
                title="校园招聘/日常实习进行中",
                location="北京/深圳",
                link="https://zhaopin.kuaishou.cn/#/campus",
                source_note="具体岗位请在官网查看"
            ))

        page.close()
    except Exception as e:
        results.append(InternPosition(
            company="快手", title=f"抓取异常: {str(e)[:60]}",
            source_note="页面加载失败"
        ))
    finally:
        browser.close()
        p.stop()
    return results


# ─── 主流程 ───────────────────────────────────────────────

SCRAPERS = [
    ("xiaohongshu", "小红书", scrape_xiaohongshu),
    ("tencent", "腾讯", scrape_tencent),
    ("pinduoduo", "拼多多", scrape_pinduoduo),
    ("bytedance", "字节跳动", scrape_bytedance),
    ("meituan", "美团", scrape_meituan),
    ("alibaba", "阿里巴巴", scrape_alibaba),
    ("baidu", "百度", scrape_baidu),
    ("kuaishou", "快手", scrape_kuaishou),
]


def generate_report(all_positions: list[InternPosition]) -> str:
    """生成 Markdown 报告"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    date_cn = datetime.now().strftime("%Y年%m月%d日")

    lines = []
    lines.append(f"# 🎓 互联网大厂 2027届实习岗位信息汇总")
    lines.append(f"")
    lines.append(f"> **更新日期：{date_cn}** | 自动抓取于 {now}")
    lines.append(f"> 数据来源：各公司官方校园招聘网站")
    lines.append(f"> ⚠️ 部分公司需登录后查看完整岗位，以下信息仅供参考")
    lines.append(f"")
    lines.append(f"| 公司 | 岗位/项目名称 | 工作地点 | 岗位要求 | 薪资 | 备注 |")
    lines.append(f"|------|-------------|---------|---------|------|------|")

    company_order = ["字节跳动", "腾讯", "阿里巴巴", "百度", "美团", "快手", "拼多多", "小红书"]
    for company in company_order:
        company_positions = [p for p in all_positions if p.company == company]
        if not company_positions:
            lines.append(f"| **{company}** | 暂未获取到数据 | - | - | - | 请访问官网查看 |")
            continue

        for idx, pos in enumerate(company_positions):
            company_display = f"**{company}**" if idx == 0 else ""
            req_short = pos.requirement[:60] + "..." if len(pos.requirement) > 60 else pos.requirement
            note = pos.source_note or ""
            link_display = f"[官网]({pos.link})" if pos.link else ""
            lines.append(f"| {company_display} | {pos.title} | {pos.location} | {req_short} | {pos.salary} | {note} {link_display} |".replace("\n", " "))

    lines.append(f"")
    lines.append(f"---")
    lines.append(f"*报告由自动化脚本自动生成，具体信息以各公司官网为准*")

    return "\n".join(lines)


def send_email(report: str, smtp_config: dict = None):
    """通过邮件发送报告"""
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    cfg = smtp_config or {
        "host": os.environ.get("SMTP_HOST", "smtp.qq.com"),
        "port": int(os.environ.get("SMTP_PORT", 465)),
        "user": os.environ.get("SMTP_USER", ""),
        "password": os.environ.get("SMTP_PASSWORD", ""),
        "to": os.environ.get("SMTP_TO", ""),
    }

    if not cfg["user"] or not cfg["password"]:
        print("[邮件] 未配置 SMTP 账号密码，跳过发送")
        return

    msg = MIMEMultipart()
    msg["Subject"] = f"互联网大厂实习信息日报 - {datetime.now().strftime('%Y-%m-%d')}"
    msg["From"] = cfg["user"]
    msg["To"] = cfg["to"]

    html_content = report.replace("\n", "<br>\n")
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    try:
        # 根据端口选择 SSL 或 STARTTLS
        if cfg["port"] == 465:
            server = smtplib.SMTP_SSL(cfg["host"], cfg["port"])
        else:
            server = smtplib.SMTP(cfg["host"], cfg["port"])
            server.starttls()
        server.login(cfg["user"], cfg["password"])
        server.send_message(msg)
        server.quit()
        print(f"[邮件] 已发送至 {cfg['to']}")
    except Exception as e:
        print(f"[邮件] 发送失败: {e}")


def main():
    parser = argparse.ArgumentParser(description="互联网大厂实习信息爬虫")
    parser.add_argument("--output", "-o", type=str, help="输出文件路径")
    parser.add_argument("--send-email", action="store_true", help="发送邮件")
    parser.add_argument("--company", "-c", type=str, help="只抓取指定公司（逗号分隔）")
    args = parser.parse_args()

    print("=" * 60)
    print("  互联网大厂 2027届实习岗位信息爬虫")
    print(f"  启动时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    all_positions = []

    scrapers_to_run = SCRAPERS
    if args.company:
        companies = [c.strip() for c in args.company.split(",")]
        scrapers_to_run = [(k, n, f) for k, n, f in SCRAPERS if n in companies]

    for key, name, scraper_fn in scrapers_to_run:
        print(f"\n▶ 正在抓取 {name}...", end=" ", flush=True)
        try:
            positions = scraper_fn()
            all_positions.extend(positions)
            print(f"✓ 获取到 {len(positions)} 条信息")
            for p in positions[:3]:
                print(f"    - {p.title}")
        except Exception as e:
            print(f"✗ 失败: {e}")
            all_positions.append(InternPosition(
                company=name, title=f"抓取失败", source_note=str(e)[:60]
            ))

    print(f"\n{'=' * 60}")
    print(f"  共获取 {len(all_positions)} 条信息")

    # 生成报告
    report = generate_report(all_positions)
    print(f"\n{report}")

    # 输出到文件
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"\n📄 报告已保存至: {args.output}")

    # 发送邮件
    if args.send_email:
        send_email(report)

    print(f"\n✅ 完成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()
