from __future__ import annotations

import argparse
import csv
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


LEGACY_DEFAULT_CSV: Path | None = None
TIMEOUT = 8
MAX_WORKERS = 24
TEXT_LIMIT = 240
REQUIRED_COLUMNS = {"url", "name", "domain", "tagline"}


TIER1_WORKED = {
    "smartlead.ai": "cold email sending and deliverability",
    "instantly.ai": "cold email + inbox warming",
    "heyreach.io": "LinkedIn multi-account outreach automation",
    "aimfox.com": "LinkedIn lead generation automation",
    "uselinxa.com": "LinkedIn outreach automation",
    "reactin.io": "LinkedIn reaction-based engagement outreach",
    "trigify.io": "LinkedIn social intent signal triggers",
    "bettercontact.rocks": "B2B email and phone enrichment",
    "companyenrich.com": "company-level B2B data enrichment",
    "clay.com": "waterfall lead enrichment",
    "amplemarket.com": "AI SDR + full outreach automation",
    "lindy.ai": "AI agents for sales workflows",
}

TIER2_WORKED = {
    "apollo.io": "B2B database + outreach sequencing",
    "gojiberry.com": "B2B sales intelligence",
    "superagi.com": "AI agent platform (sales use cases)",
    "suna.so": "open-source AI agent framework",
}

WISHLIST = {
    "windsurf.com",
    "figma.com",
    "opus.pro",
    "replit.com",
    "gumloop.com",
    "n8n.io",
    "descript.com",
    "factors.ai",
    "writer.com",
    "relevanceai.com",
    "miro.com",
    "crewai.com",
    "elicit.com",
    "base44.com",
    "typeface.ai",
    "leonardo.ai",
    "regie.ai",
    "fireflies.ai",
    "otter.ai",
    "airtable.com",
    "lindy.ai",
    "mirage.app",
    "flowiseai.com",
    "lovable.dev",
    "jasper.ai",
    "dust.tt",
    "slidesai.io",
    "zapier.com",
}

CORE_GTM_PATTERNS = {
    "cold_email": (r"\bcold email\b|\bdeliverability\b|\binbox warming\b|\bemail warm", 22),
    "outreach": (r"\boutreach\b|\bsequencing\b|\bsequence(s)?\b|\bprospect(ing)?\b", 18),
    "linkedin": (r"\blinkedin\b|\bsocial selling\b|\bsales nav\b", 18),
    "leadgen": (r"\blead gen\b|\blead generation\b|\bbook meetings\b|\bmeeting(s)? booked\b", 18),
    "enrichment": (r"\benrichment\b|\bemail finder\b|\bphone finder\b|\bcontact data\b|\bcontact info\b", 20),
    "sales_data": (r"\bB2B\b|\bsales intelligence\b|\bbuyer intent\b|\bintent data\b|\baccount research\b", 16),
    "sales_workflow": (r"\bSDR\b|\bAI SDR\b|\bsales workflow\b|\brevenue\b|\bpipeline\b|\bCRM\b", 16),
}

ADJACENT_PATTERNS = {
    "automation": (r"\bautomation\b|\bworkflow(s)?\b|\borchestrat(e|ion)\b|\bintegration(s)?\b", 14),
    "agents": (r"\bagent(s)?\b|\bagentic\b|\bautonomous\b", 14),
    "ai_platform": (r"\bAI\b|\bLLM\b|\bmodel(s)?\b|\bcopilot\b", 10),
    "productivity": (r"\bproductivity\b|\bknowledge\b|\bmeeting(s)?\b|\bnotes?\b|\btranscri", 10),
    "dev_tool": (r"\bdeveloper\b|\bcoding\b|\bcode\b|\bIDE\b|\bdeploy\b|\bbuild apps\b", 10),
    "design_creative": (r"\bdesign\b|\bvideo\b|\bcreative\b|\bcontent\b|\bslides\b|\bpresentation\b", 8),
}

BUSINESS_PATTERNS = {
    "b2b_signal": (r"\bteam(s)?\b|\bbusiness(es)?\b|\bcompany\b|\benterprise\b|\bworkspace\b", 6),
    "saas_signal": (r"\bplatform\b|\bsoftware\b|\bSaaS\b|\bAPI\b|\bintegration\b|\bdashboard\b", 6),
}

NEGATIVE_PATTERNS = {
    "consumer_social": (r"\bdating\b|\bfriends?\b|\bfamily\b|\bkids?\b|\bparent", -18),
    "gaming_entertainment": (r"\bgame(s)?\b|\bgaming\b|\bstream\b|\bmovie\b|\bmusic\b", -16),
    "lifestyle": (r"\btravel\b|\bfood\b|\brecipe\b|\bfitness\b|\bworkout\b|\bhabit\b|\bsleep\b|\bwallpaper\b", -16),
    "commerce": (r"\bshop\b|\becommerce\b|\bmarketplace\b|\bbuy\b|\bsell products\b", -10),
    "crypto": (r"\bcrypto\b|\bon-chain\b|\btoken(s)?\b|\bNFT\b", -12),
}

CATEGORY_RULES = [
    ("core_gtm", r"\bcold email\b|\bdeliverability\b|\binbox warming\b|\boutreach\b|\blinkedin\b|\blead gen\b|\benrichment\b|\bB2B\b|\bSDR\b|\bsales intelligence\b"),
    ("ai_automation", r"\bagent(s)?\b|\bautomation\b|\bworkflow(s)?\b|\borchestrat(e|ion)\b|\bintegration(s)?\b|\bno-code\b"),
    ("creative_ai", r"\bdesign\b|\bvideo\b|\bimage\b|\bcreative\b|\bslides\b|\bpresentation\b|\bbrand\b"),
    ("dev_tools", r"\bcoding\b|\bcode\b|\bdeveloper\b|\bIDE\b|\bdeploy\b|\brepo\b"),
    ("knowledge_productivity", r"\bmeeting(s)?\b|\bnotes?\b|\btranscri\b|\bknowledge\b|\bresearch\b|\bwriting\b"),
]

PRIORITY_NOW = {
    "trueemailer.com": "direct outbound email platform",
    "seamless.ai": "direct B2B data + prospecting fit",
    "reachrobin.com": "direct LinkedIn + multichannel outreach fit",
    "beacon.citronai.one": "direct B2B lead gen + LinkedIn signals fit",
    "salesleadagent.com": "direct prospecting automation fit",
    "flyweel.co": "strong revops / CRM-attribution fit",
    "kipps.ai": "sales + lead qualification automation",
    "marketgenie.tech": "lead gen + outreach + CRM workflow fit",
    "leadagentx.com": "AI sales agent fit",
    "mailmeteor.com": "email outreach + reply workflow fit",
    "sendrise.io": "cold email platform fit",
    "conversifi.io": "LinkedIn prospecting automation fit",
    "smtpfox.ai": "cold outreach mail infra fit",
    "clientsreach.com": "B2B lead gen + contact data fit",
    "scippa.app": "B2B leads + market intelligence fit",
    "flowtask.work": "pipeline research + personalization fit",
    "komo.ai": "revenue workflow / CRM automation fit",
    "konvo.in": "sales conversation + AI calling fit",
    "msgin.io": "LinkedIn inbox workflow fit",
    "clientsoo.com": "AI outbound agent fit",
    "inboxparty.com": "B2B prospecting + campaigns fit",
    "findmyicp.com": "ICP / contact discovery fit",
}

GOOD_FIT = {
    "generatifyy.ai": "website / funnel builder with marketing adjacency",
    "owlbrowser.net": "useful browser infra for scraping / outreach workflows",
    "kensakuai.com": "growth + data enrichment adjacency",
    "stacksync.com": "workflow / CRM sync adjacency",
    "chatloom.app": "AI agent platform with sales/support use cases",
    "zoye.io": "CRM + business ops adjacency",
    "synk-ai.com": "workflow-heavy B2B legal ops product",
    "legittai.com": "AI contract workflow for business teams",
    "hirebetter.io": "workflow automation with sourcing/outreach motion",
    "contactformtoapi.com": "lead routing / automation adjacency",
    "yepcode.io": "automation / integrations platform adjacency",
    "wiiz.it": "AI orchestration platform adjacency",
    "swatle.ai": "team workflow / support automation adjacency",
    "instruct.ai": "general agent automation platform adjacency",
    "githire.ai": "automated sourcing/outreach adjacency",
    "syntermedia.ai": "ad ops / campaign automation adjacency",
    "geekflare.com": "AI/search/scraping infra adjacency",
    "abliteration.ai": "enterprise AI infrastructure adjacency",
    "sinaptic.ai": "enterprise AI governance / agent platform adjacency",
    "openwork.software": "agent workspace / platform adjacency",
    "verdent.ai": "coding agent platform similar wishlist neighborhood",
    "vellum.ai": "agent builder / automation adjacency",
    "trybrowzer.com": "workflow capture + automation adjacency",
    "skillrisk.org": "AI agent security adjacency",
    "preloop.ai": "agent control plane adjacency",
    "omoios.dev": "autonomous coding workflow adjacency",
    "o11.ai": "AI workflow product for business teams",
    "notifygate.io": "ops automation adjacency",
    "nizh.com": "agent workflow platform adjacency",
    "mnexium.com": "memory infra for agent workflows adjacency",
    "git.law": "B2B legal workflow adjacency",
    "getrunstack.com": "visual LLM workflow builder adjacency",
    "dopamine.chat": "agent builder adjacency",
    "cosmicjs.com": "AI-native app / CMS platform adjacency",
    "chatbix.ai": "AI support agent platform adjacency",
    "blinq.io": "AI test automation, strong dev-tools adjacency",
    "aismarttalk.tech": "customer support / ops automation adjacency",
    "usebonobo.com": "workflow automation adjacency",
    "notte.cc": "browser agent infrastructure adjacency",
    "makeplans.tech": "workflow builder adjacency",
    "integratestack.com": "SaaS integrations / revops adjacency",
    "bolka.ai": "enterprise voice agent adjacency",
    "blink.new": "AI app builder in wishlist neighborhood",
    "askyura.com": "ops / support workflow automation adjacency",
    "getcallum.com": "calendar / scheduling workflow adjacency",
    "langfinity.ai": "strong adjacent workflow / AI tool fit",
}

MAYBE = {
    "conthunt.app": "content intelligence, but less proven buyer overlap",
    "noteviral.com": "creator / content workflow adjacency",
    "oncue.so": "social scheduling adjacency, not core motion",
    "superdesign.dev": "design tooling adjacency, but indirect fit",
    "athenachat.bot": "broad AI workspace, but weak direct overlap",
    "templated.io": "creative automation adjacency",
    "witfy.social": "social media automation adjacency",
    "storychief.io": "content marketing workflow adjacency",
    "nebly.app": "content repurposing adjacency",
    "popjam.io": "ad creative tooling adjacency",
    "pixelora.art": "design agent adjacency",
    "pagepop.ai": "design workspace adjacency",
    "aidesigner.ai": "design generation adjacency",
    "obello.com": "creative design platform adjacency",
    "hyrstudio.com": "creator asset automation adjacency",
    "blog2visuals.dev": "content-to-visual workflow adjacency",
    "seo45.com": "SEO automation adjacency",
    "zunno.app": "social publishing adjacency",
    "chatfromai.com": "Instagram DM automation adjacency",
    "freqlab.app": "creative dev tooling adjacency",
}

SKIP = {
    "audioscribe.org": "transcription tool, too far from target motion",
    "tickerbell.com": "stock research, off ICP",
    "finalrun.app": "mobile test automation, weak buyer overlap",
    "aocr.in": "OCR infrastructure, too generic for target list",
    "reflectwith.facilitron.ai": "meeting facilitation, weak GTM relevance",
    "the-profanity-api.com": "content moderation API, off ICP",
    "vemory.ai": "meeting translation/transcription, weak fit",
    "vibecoders.bio": "developer portfolio / backlinks product, off ICP",
    "dimo.org": "vehicle platform, off ICP",
    "remote-tracking.goldenowl.asia": "workforce tracking, off ICP",
    "outtalent.com": "career mentorship, off ICP",
    "indie-forge.net": "community platform, off ICP",
    "videomp3word.com": "media conversion utility, off ICP",
    "lurna.co": "student study product, off ICP",
    "coverpilot.ai": "job application tool, off ICP",
    "tetramatrix.github.io": "broken GitHub Pages product page",
}


@dataclass
class SiteInfo:
    final_url: str
    title: str
    meta_description: str
    h1: str
    fetch_status: str


def squash(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def clip(text: str, limit: int = TEXT_LIMIT) -> str:
    text = squash(text)
    return text if len(text) <= limit else text[: limit - 3] + "..."


def fetch_site(domain: str) -> SiteInfo:
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = [f"https://{domain}", f"http://{domain}"]
    last_error = "fetch_failed"
    for url in urls:
        try:
            response = requests.get(url, timeout=TIMEOUT, headers=headers, allow_redirects=True)
            html = response.text[:300000]
            soup = BeautifulSoup(html, "html.parser")
            title = squash(soup.title.get_text(" ", strip=True) if soup.title else "")
            meta = ""
            meta_tag = soup.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
            if meta_tag and meta_tag.get("content"):
                meta = squash(meta_tag["content"])
            if not meta:
                og_tag = soup.find("meta", attrs={"property": re.compile(r"^og:description$", re.I)})
                if og_tag and og_tag.get("content"):
                    meta = squash(og_tag["content"])
            h1_tag = soup.find("h1")
            h1 = squash(h1_tag.get_text(" ", strip=True) if h1_tag else "")
            return SiteInfo(
                final_url=response.url,
                title=clip(title),
                meta_description=clip(meta),
                h1=clip(h1),
                fetch_status=f"http_{response.status_code}",
            )
        except Exception as exc:  # noqa: BLE001
            last_error = exc.__class__.__name__
    return SiteInfo("", "", "", "", last_error)


def add_pattern_scores(text: str, patterns: dict[str, tuple[str, int]], reasons: list[str], matched: set[str]) -> int:
    score = 0
    for label, (pattern, points) in patterns.items():
        if re.search(pattern, text, flags=re.I):
            score += points
            matched.add(label)
            reasons.append(label.replace("_", " "))
    return score


def choose_category(text: str) -> str:
    for category, pattern in CATEGORY_RULES:
        if re.search(pattern, text, flags=re.I):
            return category
    return "other"


def fit_label(score: int) -> str:
    if score >= 80:
        return "high"
    if score >= 60:
        return "medium"
    if score >= 40:
        return "low"
    return "reject"


def fit_bucket(score: int) -> str:
    if score >= 80:
        return "core_icp"
    if score >= 60:
        return "adjacent_icp"
    if score >= 40:
        return "strategic_bet"
    return "not_now"


def build_reason(domain: str, exact_match: str, reasons: list[str], site: SiteInfo) -> str:
    parts: list[str] = []
    if exact_match:
        parts.append(exact_match)
    unique_reasons: list[str] = []
    for reason in reasons:
        if reason not in unique_reasons:
            unique_reasons.append(reason)
    if unique_reasons:
        parts.append("signals: " + ", ".join(unique_reasons[:4]))
    if site.meta_description:
        parts.append("site: " + site.meta_description)
    elif site.title:
        parts.append("site: " + site.title)
    return " | ".join(parts) if parts else f"limited signals for {domain}"


def score_row(row: dict[str, str], site: SiteInfo) -> dict[str, str]:
    domain = (row.get("domain") or "").strip().lower()
    text = squash(" ".join([row.get("name") or "", row.get("tagline") or "", site.title, site.meta_description, site.h1, domain]))
    reasons: list[str] = []
    matched: set[str] = set()
    score = 0
    exact_match = ""
    core_hits = 0
    negative_hits = 0

    if domain in TIER1_WORKED:
        score = 98
        exact_match = f"exact tier 1 comp: {TIER1_WORKED[domain]}"
    elif domain in TIER2_WORKED:
        score = 88
        exact_match = f"exact tier 2 comp: {TIER2_WORKED[domain]}"
    elif domain in WISHLIST:
        score = 80
        exact_match = "exact wishlist company"

    if not exact_match:
        score += add_pattern_scores(text, CORE_GTM_PATTERNS, reasons, matched)
        core_hits = len(matched & set(CORE_GTM_PATTERNS))
        score += add_pattern_scores(text, ADJACENT_PATTERNS, reasons, matched)
        score += add_pattern_scores(text, BUSINESS_PATTERNS, reasons, matched)
        score += add_pattern_scores(text, NEGATIVE_PATTERNS, reasons, matched)
        negative_hits = len(matched & set(NEGATIVE_PATTERNS))

        if any(k in matched for k in {"outreach", "linkedin", "leadgen", "enrichment", "sales_data", "sales_workflow"}) and any(
            k in matched for k in {"automation", "agents", "ai_platform"}
        ):
            score += 10
            reasons.append("gtm + ai automation overlap")

        if re.search(r"\bdesign\b|\bdeveloper\b|\bworkflow\b|\bmeeting\b|\bwriting\b|\bresearch\b", text, flags=re.I):
            score += 6
            reasons.append("wishlist adjacency")

        if core_hits == 0:
            score = min(score, 54)
            reasons.append("no core gtm signal")
        elif core_hits == 1:
            score = min(score, 79)
            reasons.append("single core gtm signal")

        if negative_hits and core_hits == 0:
            score = min(score, 35)
            reasons.append("consumer/off-ICP tilt")

        if re.search(
            r"\bworkflow\b|\bautomation\b|\bagent(s)?\b|\bdeveloper\b|\bdesign\b|\bcreative\b|\bmeeting(s)?\b|\bwriting\b|\bresearch\b|\btranscri|\bcollaboration\b|\bno-code\b|\blow-code\b",
            text,
            flags=re.I,
        ):
            score += 8
            reasons.append("broad strategic adjacency")

        if re.search(
            r"\b(ai meeting|transcription|note taker|workflow automation|agent platform|developer tools|design platform|research assistant|writing assistant|collaboration|canvas|whiteboard)\b",
            text,
            flags=re.I,
        ):
            score += 10
            reasons.append("wishlist neighborhood")

        if re.search(r"\bfor teams\b|\bfor businesses\b|\bfor sales teams\b|\bfor marketers\b|\benterprise\b|\bworkspace\b|\bplatform for\b", text, flags=re.I):
            score += 6
            reasons.append("b2b commercial signal")

    score = max(0, min(100, score))
    return {
        **row,
        "website_final_url": site.final_url,
        "website_title": site.title,
        "website_meta_description": site.meta_description,
        "website_h1": site.h1,
        "website_fetch_status": site.fetch_status,
        "qualification_score": str(score),
        "qualification_label": fit_label(score),
        "qualification_bucket": fit_bucket(score),
        "qualification_category": choose_category(text),
        "qualification_reason": build_reason(domain, exact_match, reasons, site),
    }


def clean(text: str) -> str:
    return " ".join((text or "").replace("\n", " ").split()).lower()


def heuristic_label(text: str) -> tuple[str, str]:
    if any(
        key in text
        for key in ["cold email", "outreach", "prospecting", "lead generation", "lead gen", "sales agent", "b2b lead", "verified emails", "linkedin prospecting", "crm", "sales teams"]
    ):
        return "priority_now", "direct GTM / outbound motion"
    if any(
        key in text
        for key in ["workflow", "automation", "agent", "developer", "design", "support", "transcription", "meeting", "research", "writing", "integration", "chatbot", "voice ai", "no-code"]
    ):
        return "good_fit", "strong adjacent workflow / AI tool fit"
    if any(key in text for key in ["social media", "content", "creative", "seo", "ads", "marketing", "community", "portfolio"]):
        return "maybe", "adjacent but weaker overlap"
    return "skip", "weak overlap with target company set"


def curate_row(row: dict[str, str]) -> dict[str, str]:
    domain = row["domain"].strip().lower()
    text = clean(" ".join([row.get("name", ""), row.get("tagline", ""), row.get("website_meta_description", ""), row.get("website_title", "")]))
    if domain in PRIORITY_NOW:
        label, reason = "priority_now", PRIORITY_NOW[domain]
    elif domain in GOOD_FIT:
        label, reason = "good_fit", GOOD_FIT[domain]
    elif domain in MAYBE:
        label, reason = "maybe", MAYBE[domain]
    elif domain in SKIP:
        label, reason = "skip", SKIP[domain]
    else:
        label, reason = heuristic_label(text)
    return {**row, "curated_label": label, "curated_reason": reason}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        fieldnames = set(reader.fieldnames or [])
        missing = sorted(REQUIRED_COLUMNS - fieldnames)
        if missing:
            raise ValueError(f"Missing required columns: {', '.join(missing)}")
        return list(reader)


def resolve_input_csv(input_csv_arg: str | None) -> Path:
    if input_csv_arg:
        return Path(input_csv_arg)

    workspace_csvs = [
        path
        for path in Path.cwd().glob("*.csv")
        if not path.name.startswith("qualified-products")
    ]
    if len(workspace_csvs) == 1:
        return workspace_csvs[0]
    if LEGACY_DEFAULT_CSV is not None and LEGACY_DEFAULT_CSV.exists():
        return LEGACY_DEFAULT_CSV
    raise ValueError("No input CSV provided. Pass --input-csv or place one CSV in current workspace.")


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def write_shortlist_review(path: Path, rows: list[dict[str, str]]) -> None:
    parts: list[str] = []
    for idx, row in enumerate(rows, start=1):
        parts.extend(
            [
                f"{idx}. score={row['qualification_score']} bucket={row['qualification_bucket']} domain={row['domain']} name={row['name']}",
                f"tagline: {clip(row.get('tagline', '') or '', 220)}",
                f"site: {row.get('website_meta_description') or row.get('website_title') or ''}",
                f"reason: {row.get('qualification_reason', '')}",
                "",
            ]
        )
    path.write_text("\n".join(parts), encoding="utf-8")


def split_by_label(rows: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    out = {"priority_now": [], "good_fit": [], "maybe": [], "skip": []}
    for row in rows:
        out[row["curated_label"]].append(row)
    return out


def write_workbook(path: Path, curated_rows: list[dict[str, str]], split_rows: dict[str, list[dict[str, str]]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "curated_all"
    fieldnames = list(curated_rows[0].keys())
    sheets = [("curated_all", curated_rows)] + [(label, split_rows[label]) for label in ["priority_now", "good_fit", "maybe", "skip"]]
    for idx, (name, rows) in enumerate(sheets):
        sheet = ws if idx == 0 else wb.create_sheet(title=name)
        sheet.append(fieldnames)
        for row in rows:
            sheet.append([row.get(field, "") for field in fieldnames])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-csv")
    parser.add_argument("--output-dir", default=str(Path.cwd() / "outputs"))
    parser.add_argument("--shortlist-size", type=int, default=121)
    args = parser.parse_args()

    input_csv = resolve_input_csv(args.input_csv)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    broad_csv = output_dir / "qualified-products-broad.csv"
    shortlist_csv = output_dir / "qualified-products-broad-shortlist.csv"
    shortlist_review = output_dir / "shortlist-review.txt"
    curated_csv = output_dir / "qualified-products-curated.csv"
    workbook_path = output_dir / "qualified-products-google-upload.xlsx"

    rows = read_csv(input_csv)
    domains = [row["domain"] for row in rows if row.get("domain")]
    site_info_by_domain: dict[str, SiteInfo] = {}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(fetch_site, domain): domain for domain in domains}
        for future in as_completed(futures):
            domain = futures[future]
            try:
                site_info_by_domain[domain] = future.result()
            except Exception as exc:  # noqa: BLE001
                site_info_by_domain[domain] = SiteInfo("", "", "", "", exc.__class__.__name__)

    broad_rows = [score_row(row, site_info_by_domain.get(row["domain"], SiteInfo("", "", "", "", "missing"))) for row in rows]
    write_csv(broad_csv, broad_rows)

    ranked_rows = sorted(broad_rows, key=lambda r: (-int(r["qualification_score"]), r["domain"]))
    shortlist_rows = ranked_rows[: args.shortlist_size]
    write_csv(shortlist_csv, shortlist_rows)
    write_shortlist_review(shortlist_review, shortlist_rows)

    curated_rows = [curate_row(row) for row in shortlist_rows]
    label_rank = {"priority_now": 0, "good_fit": 1, "maybe": 2, "skip": 3}
    curated_rows.sort(key=lambda r: (label_rank[r["curated_label"]], -int(r["qualification_score"]), r["domain"]))
    write_csv(curated_csv, curated_rows)

    split_rows = split_by_label(curated_rows)
    for label, label_rows in split_rows.items():
        write_csv(output_dir / f"qualified-products-{label}.csv", label_rows)

    write_workbook(workbook_path, curated_rows, split_rows)

    print(f"INPUT={input_csv}")
    print(f"OUTPUT_DIR={output_dir}")
    print(f"BROAD_ROWS={len(broad_rows)}")
    print(f"SHORTLIST_ROWS={len(shortlist_rows)}")
    for label in ["priority_now", "good_fit", "maybe", "skip"]:
        print(f"{label.upper()}={len(split_rows[label])}")
    print(f"WORKBOOK={workbook_path}")


if __name__ == "__main__":
    main()
