"""Сбор открытых коммерческих процедур с Tender.Pro.

Адаптер не использует авторизацию и не пытается обходить ограничения сайта.
Недоступные публично цена, сумма и ОКПД2 остаются пустыми.
"""

from __future__ import annotations

import re
import time
from typing import Any, Callable
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup, Tag


SOURCE_NAME = "Tender.Pro"
BASE_URL = "https://www.tender.pro"
LIST_URL = f"{BASE_URL}/api/tenders/list"
DETAIL_URL = f"{BASE_URL}/api/tender/{{tender_id}}/view_public"

# Tender.Pro иногда указывает город вместо субъекта. Словарь нужен только для
# проверки региона; в Excel записывается нормальное наименование субъекта.
REGION_MARKERS = {
    "Республика Татарстан": (
        "республика татарстан", "татарстан", "казань", "набережные челны",
        "нижнекамск", "альметьевск", "елабуга", "бугульма", "зеленодольск",
    ),
    "Самарская область": (
        "самарская область", "самара", "тольятти", "сызрань", "новокуйбышевск",
        "чапаевск", "жигулевск", "жигулёвск", "отрадный",
    ),
}


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def _number(value: Any) -> float | None:
    text = re.sub(r"[^\d,.-]", "", _clean(value)).replace(",", ".")
    if not text:
        return None
    if text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(text)
    except ValueError:
        return None


def _get_html(
    session: requests.Session,
    url: str,
    timeout: int,
    params: dict[str, Any] | None = None,
) -> tuple[str | None, str]:
    try:
        response = session.get(url, params=params, timeout=timeout, verify=False)
        print(f"HTTP {response.status_code}: {response.url}")
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
        return response.text, response.url
    except requests.RequestException as exc:
        print(f"ОШИБКА Tender.Pro: {exc}")
        return None, url


def _extract_list_items(soup: BeautifulSoup) -> list[dict[str, str]]:
    """Извлекает ID и название, не завязываясь на CSS-классы сайта."""
    found: dict[str, dict[str, str]] = {}
    for link in soup.find_all("a", href=True):
        href = link.get("href", "")
        text = _clean(link.get_text(" ", strip=True))
        match = re.search(r"/tender/(\d+)/", href)
        if not match:
            # В списке ID часто находится прямо в тексте: (id1201736).
            match = re.search(r"\bid\s*(\d{5,})\b", f"{href} {text}", re.I)
        if not match:
            continue
        tender_id = match.group(1)
        if not text or text.isdigit() or len(text) < 8:
            continue
        current = found.get(tender_id)
        if current is None or len(text) > len(current["title"]):
            found[tender_id] = {"id": tender_id, "title": text}
    return list(found.values())


def _next_page(soup: BeautifulSoup, current_page: int) -> str | None:
    wanted = str(current_page + 1)
    for link in soup.find_all("a", href=True):
        if _clean(link.get_text(" ", strip=True)) != wanted:
            continue
        href = link.get("href", "")
        if "tenders/list" in href or "page" in href:
            return urljoin(BASE_URL, href)
    return None


def _matching_regions(text: str) -> list[str]:
    lowered = text.lower().replace("ё", "е")
    result = []
    for region, markers in REGION_MARKERS.items():
        if any(marker.replace("ё", "е") in lowered for marker in markers):
            result.append(region)
    return result


def _value_after_label(soup: BeautifulSoup, label: str) -> str:
    target = label.lower()
    for node in soup.find_all(["div", "span", "td", "th", "dt", "strong"]):
        text = _clean(node.get_text(" ", strip=True))
        if text.lower().rstrip(":") != target.rstrip(":"):
            continue
        sibling = node.find_next_sibling()
        if sibling:
            value = _clean(sibling.get_text(" ", strip=True))
            if value:
                return value
        nxt = node.find_next()
        if nxt and nxt is not node:
            value = _clean(nxt.get_text(" ", strip=True))
            if value and value.lower().rstrip(":") != target.rstrip(":"):
                return value
    return ""


def _find_goods_table(soup: BeautifulSoup) -> Tag | None:
    best: tuple[int, Tag | None] = (0, None)
    for table in soup.find_all("table"):
        text = _clean(table.get_text(" ", strip=True)).lower()
        score = sum(word in text for word in ("наименование", "кол-во", "количество", "ед.изм", "единица"))
        if score > best[0]:
            best = (score, table)
    return best[1] if best[0] >= 2 else None


def _parse_goods(table: Tag | None) -> list[dict[str, Any]]:
    if table is None:
        return []
    goods: list[dict[str, Any]] = []
    for row in table.find_all("tr"):
        values = [_clean(cell.get_text(" ", strip=True)) for cell in row.find_all(["td", "th"], recursive=False)]
        if len(values) < 2:
            continue
        row_text = " ".join(values).lower()
        if "наименование" in row_text and ("кол" in row_text or "ед." in row_text):
            continue
        # Первая ячейка нередко является порядковым номером.
        offset = 1 if values[0].isdigit() and len(values) >= 3 else 0
        name = values[offset]
        if not name or len(name) < 3:
            continue
        quantity = _number(values[offset + 1]) if len(values) > offset + 1 else None
        unit = values[offset + 2] if len(values) > offset + 2 else ""
        goods.append({"name": name, "quantity": quantity, "unit": unit})
    return goods


def _parse_detail(
    session: requests.Session,
    tender: dict[str, str],
    query: str,
    timeout: int,
) -> list[dict[str, Any]]:
    url = DETAIL_URL.format(tender_id=tender["id"])
    html, final_url = _get_html(session, url, timeout, {"mode": "buy", "page": "goods"})
    if not html:
        return []
    soup = BeautifulSoup(html, "html.parser")
    full_text = _clean(soup.get_text(" ", strip=True))
    regions = _matching_regions(full_text)
    if not regions:
        return []

    heading = soup.find(["h1", "h2"])
    title = _clean(heading.get_text(" ", strip=True)) if heading else tender["title"]
    customer = _value_after_label(soup, "Организатор конкурса")
    goods = _parse_goods(_find_goods_table(soup))
    if not goods:
        goods = [{"name": title, "quantity": None, "unit": ""}]

    result = []
    for good in goods:
        result.append({
            "source": SOURCE_NAME,
            "reg_number": tender["id"],
            "law": "Коммерческая",
            "customer": customer,
            "region": ", ".join(regions),
            "search_queries": query,
            "okpd": "",
            "ktru": "",
            "name": good["name"],
            "unit": good["unit"],
            "quantity": good["quantity"],
            "price": None,
            "amount": None,
            "url": final_url,
        })
    return result


def collect_tender_pro(
    session: requests.Session,
    search_queries: list[str],
    classify: Callable[[dict[str, Any]], tuple[str, str]],
    max_pages: int = 2,
    delay: float = 0.8,
    timeout: int = 40,
) -> list[dict[str, Any]]:
    """Возвращает позиции Tender.Pro в общем формате проекта."""
    candidates: dict[str, dict[str, Any]] = {}
    print("\n" + "=" * 80)
    print("КОММЕРЧЕСКИЕ ЗАКУПКИ: TENDER.PRO")
    print("=" * 80)

    for query_no, query in enumerate(search_queries, 1):
        print(f"\n[{query_no}/{len(search_queries)}] {query}")
        page_url: str | None = LIST_URL
        params: dict[str, Any] | None = {
            "tender_name": query,
            "goods": query,
            "tender_state": 100,
            "tender_type": 100,
        }
        for page in range(1, max_pages + 1):
            if not page_url:
                break
            html, _ = _get_html(session, page_url, timeout, params)
            params = None  # параметры уже находятся в ссылке следующей страницы
            if not html:
                break
            soup = BeautifulSoup(html, "html.parser")
            items = _extract_list_items(soup)
            before = len(candidates)
            for item in items:
                entry = candidates.setdefault(item["id"], {**item, "queries": set()})
                entry["queries"].add(query)
            print(f"  страница {page}: найдено {len(items)}, новых {len(candidates) - before}")
            page_url = _next_page(soup, page)
            time.sleep(delay)

    print(f"\nУНИКАЛЬНЫХ ПРОЦЕДУР TENDER.PRO: {len(candidates)}")
    products: list[dict[str, Any]] = []
    for index, tender in enumerate(candidates.values(), 1):
        query = ", ".join(sorted(tender["queries"]))
        print(f"[{index}/{len(candidates)}] Tender.Pro {tender['id']}")
        parsed = _parse_detail(session, tender, query, timeout)
        for product in parsed:
            product["category"], product["reason"] = classify(product)
        products.extend(parsed)
        time.sleep(delay)
    print(f"ПОЗИЦИЙ TENDER.PRO В ЦЕЛЕВЫХ РЕГИОНАХ: {len(products)}")
    return products
