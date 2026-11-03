#!/usr/bin/env python3
"""Scoped AccessBG -> supplied Temu template. Python 3.11+, lxml.

No store navigation, recommended products or global sitemap is followed.
The OOXML writer changes only Template data rows; all other ZIP members and
reserved template headers are copied byte-for-byte from the supplied workbook.
"""
import argparse
import concurrent.futures
import csv
import hashlib
import json
import logging
import re
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from lxml import etree, html

BASE = "https://accessbg.com"
SEEDS = {"phone_planet": BASE + "/search?searchTerm=phone%20planet",
         "bardahl": BASE + "/masla-i-dobavki-bardahl"}
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
S = "{" + NS["s"] + "}"
LOG = logging.getLogger("accessbg")


def text(node):
    if isinstance(node, str):
        return " ".join(node.split())
    return " ".join(node.itertext()).strip() if node is not None else ""


def column(ref):
    return re.sub(r"\d+$", "", ref)


def colnum(col):
    n = 0
    for c in col:
        n = n * 26 + ord(c) - 64
    return n


def money(value):
    if value is None or value == "":
        return None
    m = re.search(r"\d+(?:[.,]\d+)?", str(value))
    return float(Decimal(m[0].replace(",", ".")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)) if m else None


class Client:
    def __init__(self, cache):
        self.cache = Path(cache)
        self.cache.mkdir(parents=True, exist_ok=True)

    def get(self, url, fresh=False):
        if urllib.parse.urlparse(url).hostname != "accessbg.com":
            raise ValueError("Out-of-scope host: " + url)
        path = self.cache / (hashlib.sha256(url.encode()).hexdigest() + ".html")
        if not fresh and path.exists():
            return path.read_text(encoding="utf-8")
        for attempt in range(3):
            try:
                req = urllib.request.Request(url, headers={
                    "User-Agent": "AccessBG-CatalogExport/1.0", "Accept": "text/html,application/json",
                    "X-Requested-With": "XMLHttpRequest"})
                try:
                    response = urllib.request.urlopen(req, timeout=45)
                except urllib.error.HTTPError as exc:
                    # AccessBG's canonical redirect loops for encoded Unicode slugs
                    # such as the № sign. The same product route accepts an extra
                    # escaping layer; keep the original URL as the source/cache key.
                    parsed = urllib.parse.urlsplit(url)
                    if exc.code not in (301, 302, 303, 307, 308) or not re.search(r"%[89A-Fa-f][0-9A-Fa-f]", parsed.path):
                        raise
                    alternate = urllib.parse.urlunsplit(parsed._replace(path=parsed.path.replace("%", "%25")))
                    response = urllib.request.urlopen(urllib.request.Request(alternate, headers=dict(req.header_items())), timeout=45)
                with response:
                    value = response.read().decode("utf-8")
                path.write_text(value, encoding="utf-8")
                return value
            except (urllib.error.URLError, TimeoutError, OSError):
                if attempt == 2:
                    raise
                time.sleep(1 + attempt)


def cards(markup):
    root = html.fromstring(markup or "<div/>")
    return set(root.xpath('//div[contains(concat(" ",normalize-space(@class)," ")," product-card ")]//a[contains(@href,"/product/")]/@href'))


def discover(client):
    products, counts = defaultdict(set), {}
    for group, seed in SEEDS.items():
        root = html.fromstring(client.get(seed, fresh=True))
        cid = root.xpath('//*[@id="selected-category-id"]/@value')
        if group == "bardahl" and not cid:
            raise RuntimeError("BARDAHL category ID not found: site structure changed")
        seen, total, page, seen_pages = set(), None, 1, set()
        while True:
            params = {"page": page, "count": 100}
            params.update({"searchTerm": "phone planet"} if group == "phone_planet" else {"category": cid[0]})
            response = json.loads(client.get(BASE + "/getFilteredProducts?" + urllib.parse.urlencode(params), fresh=True))
            if not response.get("success"):
                raise RuntimeError("Listing request unsuccessful: " + group)
            data = response["data"]
            meta = data["paginationMetaData"]
            total = int(meta["totalProducts"])
            batch = cards(data["html"])
            signature = tuple(sorted(batch))
            if signature in seen_pages and page > 1:
                raise RuntimeError("Repeated listing page: " + group)
            seen_pages.add(signature)
            seen.update(batch)
            if page >= int(meta["lastPage"]):
                break
            if not batch or page > 500:
                raise RuntimeError("Pagination ended unexpectedly: " + group)
            page += 1
        if len(seen) != total:
            raise RuntimeError(f"Coverage mismatch {group}: found {len(seen)}, site reports {total}")
        counts[group] = {"reported": total, "discovered": len(seen)}
        for url in seen:
            if urllib.parse.urlparse(url).hostname != "accessbg.com" or not urllib.parse.urlparse(url).path.startswith("/product/"):
                raise ValueError("Unexpected product URL: " + url)
            products[url].add(group)
        LOG.info("%s: %s products", group, len(seen))
    return products, counts


def parse_product(markup, url, groups):
    root = html.fromstring(markup)
    records = []
    for raw in root.xpath('//script[@type="application/ld+json"]/text()'):
        try:
            obj = json.loads(raw)
            records.extend(obj if isinstance(obj, list) else obj.get("@graph", [obj]))
        except (ValueError, AttributeError):
            continue
    data = next((x for x in records if x.get("@type") == "Product"), None)
    if data is None:
        raise ValueError("Product JSON-LD missing")
    def first(xpath):
        values = root.xpath(xpath)
        return text(values[0]) if values else ""
    sku = str(data.get("sku") or first('//*[@id="product-details-sku"]')).strip()
    name = text(str(data.get("name") or first('//*[@id="product-details-title"]')))
    if not sku or not name:
        raise ValueError("Product name or original SKU missing")
    offer = data.get("offers", {})
    if isinstance(offer, list):
        offer = offer[0] if offer else {}
    nodes = root.xpath('//*[@id="product-descriptions"]//*[contains(@id,"-description-")]')
    description = "\n".join(text(n) for n in nodes)
    if not description:
        description = str(data.get("description", ""))
    # The site explicitly exposes original images in these anchors.
    images = list(dict.fromkeys(urllib.parse.urljoin(BASE, u) for u in root.xpath('//a[contains(@id,"image-anchor")]/@href') if "/products/main/" in u))
    if not images:
        images = list(dict.fromkeys(urllib.parse.urljoin(BASE, u) for u in root.xpath('//a[@data-lightbox]/@href') if "/products/" in u))
    issues = []
    if not images:
        issues.append("No original image URL exposed by product gallery")
    buttons = root.xpath('//button[contains(@id,"-buy-button")]')
    if any(b.get("data-is-configurable") not in (None, "", "0", "false") for b in buttons):
        issues.append("Configurable product: explicit variant export/overrides required; parent not importable")
    # Never enumerate arbitrary selects or synthesize a Cartesian product of variants.
    crumbs = [x for x in records if x.get("@type") == "BreadcrumbList"]
    breadcrumb = " / ".join(str(x.get("item", {}).get("name", "")) for b in crumbs for x in b.get("itemListElement", []))
    price_nodes = root.xpath('//*[@id="product-details-prices"]//span[contains(@id,"price-rance") and not(contains(@id,"discounted"))]')
    regular = money(text(price_nodes[0])) if price_nodes else money(offer.get("price"))
    avail = offer.get("availability", "")
    availability = "in_stock" if avail.endswith("/InStock") else "out_of_stock" if avail.endswith(("/OutOfStock", "/Discontinued", "/SoldOut")) else "unknown"
    return {"url": url, "groups": sorted(groups), "sku": sku, "name": name,
            "description": description, "breadcrumb": breadcrumb, "images": images,
            "price": money(offer.get("price")), "regular_price": regular,
            "currency": offer.get("priceCurrency"), "availability": availability,
            "issues": issues}


class Template:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(path) as z:
            strings = []
            if "xl/sharedStrings.xml" in z.namelist():
                strings = ["".join(e.itertext()) for e in etree.fromstring(z.read("xl/sharedStrings.xml"))]
            wb = etree.fromstring(z.read("xl/workbook.xml"))
            rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            targets = {e.get("Id"): e.get("Target") for e in rels}
            self.sheets, self.paths = {}, {}
            for e in wb.find(S + "sheets"):
                target = targets[e.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
                path = target.lstrip("/") if target.startswith("/") else "xl/" + target
                self.paths[e.get("name")] = path
                rows = []
                for row in etree.fromstring(z.read(path)).findall("s:sheetData/s:row", NS):
                    values = {}
                    for c in row:
                        v = c.find(S + "v")
                        val = v.text if v is not None else "".join(c.find(S + "is").itertext()) if c.find(S + "is") is not None else ""
                        if c.get("t") == "s":
                            val = strings[int(val)]
                        if val:
                            values[column(c.get("r"))] = val
                    if values:
                        rows.append((int(row.get("r")), values))
                self.sheets[e.get("name")] = rows
        template = dict(self.sheets["Template"])
        self.labels, self.ids = template[2], template[4]
        self.categories = {v["A"]: v["B"] for n,v in self.sheets["Category Name"]}
        self.modes = {v["A"]: v for n,v in self.sheets["GoodsLevelMode"]}
        self.conditions = {v["A"]: list(v.values())[2:] for n,v in self.sheets["Conditions List"]}
        self.valid = {v["B"]: list(v.values())[1:] for n,v in self.sheets["Valid Values"] if "B" in v}

    def allowed(self, col, cid, row):
        label = self.labels[col]
        candidates = [(k,v) for k,v in self.valid.items() if k == label or k == f"{label} - [{cid}]"]
        # Conditional dropdowns include the selected controlling value and category.
        if not candidates:
            candidates = [(k,v) for k,v in self.valid.items() if k.startswith(label + " - [") and k.endswith(f"- [{cid}]") and all(x in row.values() or x == cid for x in re.findall(r"\[([^]]+)\]", k))]
        return list(dict.fromkeys(x for k,values in candidates for x in values))

    def active(self, col, row):
        ident = self.ids[col]
        triggers = self.conditions.get(ident)
        if not triggers:
            return None  # unknown conditional rules block readiness conservatively
        dep = ident.split("|", 1)[1] if "|" in ident else ""
        dependencies = [c for c,i in self.ids.items() if any(i == "t_3_Property:" + d for d in dep.split(":"))]
        values = [row.get(c) for c in dependencies if row.get(c)]
        if not values:
            return None
        return any(str(v) in triggers for v in values)

    def validate(self, row):
        cid = str(row.get("E", ""))
        if cid not in self.categories:
            return ["Category cannot be mapped to supplied template"]
        required = set(self.modes[cid + "_require"]) - {"A"}
        conditional = set(self.modes[cid + "_condition_require"]) - {"A"}
        disabled = set(self.modes[cid + "_disabled"]) - {"A"}
        issues = []
        # Duplicate columns for a multiselect property require one value per property.
        done = set()
        for col in sorted(required | conditional, key=colnum):
            ident = self.ids[col]
            if ident in done:
                continue
            done.add(ident)
            siblings = [c for c in required | conditional if self.ids[c] == ident]
            if col in conditional:
                active = self.active(col, row)
                if active is False:
                    continue
                if active is None:
                    issues.append("Unresolved conditional rule: " + self.labels[col])
                    continue
            if not any(row.get(c) not in (None, "") for c in siblings):
                issues.append("Missing: " + self.labels[col] + " (" + col + ")")
        for c,value in row.items():
            if value in (None, "") or c not in self.labels:
                continue
            if c in disabled:
                issues.append("Disabled column populated: " + c)
            allowed = self.allowed(c, cid, row)
            if allowed and str(value) not in allowed:
                issues.append("Invalid dropdown value " + c + ": " + str(value))
        for c in ["OA", "OC", "OE", "OF", "OG", "OH"]:
            if row.get(c) is not None:
                try:
                    if float(row[c]) <= 0:
                        issues.append("Must be positive: " + c)
                except (ValueError, TypeError):
                    issues.append("Must be numeric: " + c)
        if row.get("OA") and row.get("OC") and float(row["OA"]) > float(row["OC"]):
            issues.append("Base price exceeds list price")
        return issues

    def workbook_bytes(self, rows):
        """Preserve namespaces, macros, validations, hidden sheets and reserved rows."""
        import io
        target = self.paths["Template"]
        stream = io.BytesIO()
        with zipfile.ZipFile(self.path) as source, zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as dest:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == target:
                    # Replace row bodies only. Do not reserialize the complete worksheet.
                    raw = data.decode("utf-8")
                    start = re.search(r"<sheetData(?:\s[^>]*)?>", raw)
                    end = raw.index("</sheetData>", start.end())
                    existing = raw[start.end():end]
                    first_data = next((m.start() for m in re.finditer(r'<row\b[^>]*\br="(\d+)"', existing) if int(m[1]) >= 5), len(existing))
                    headers = existing[:first_data]
                    if len(re.findall(r"<row\b", headers)) != 4:
                        raise ValueError("Expected exactly four reserved template header rows")
                    body = []
                    for n, values in enumerate(rows, 5):
                        r = etree.Element(S + "row", r=str(n), nsmap={None: NS["s"]})
                        for col,value in sorted(values.items(), key=lambda kv: colnum(kv[0])):
                            if value in (None, ""):
                                continue
                            c = etree.SubElement(r, S + "c", r=f"{col}{n}")
                            if isinstance(value, (int,float)):
                                etree.SubElement(c, S + "v").text = str(value)
                            else:
                                c.set("t", "inlineStr")
                                etree.SubElement(etree.SubElement(c, S + "is"), S + "t").text = str(value)
                        body.append(etree.tostring(r, encoding="unicode"))
                    raw = raw[:start.end()] + headers + "".join(body) + raw[end:]
                    data = raw.encode("utf-8")
                dest.writestr(info, data)
        return stream.getvalue()


def classify(product):
    n = product["name"].lower()
    if "bardahl" in product["groups"]:
        # Car-care cosmetics have no matching leaf in this particular template.
        if re.search(r"драскотин|полиращ|полираща|текстил|джанти|гризачи|дезинфекция на купе|шампоан", n):
            return None
        if re.search(r"\d+\s*w\s*\d+", n) and "добавка" not in n:
            return "20414" if re.search(r"(?:75|80|85)\s*w", n) else "20416"
        rules = [
            (r"антифриз", "20372"), (r"dot\s*[345]|спирачн.*течност", "20373"),
            (r"чистачки|стъкла", "20421"), (r"промив.*радиатор", "20394"),
            (r"промив.*двигател", "20390"), (r"хидравлични повдигачи", "20364"),
            (r"кормилна рейка", "20370"),
            (r"хидравлич", "20369"), (r"скорост|gear|atf|трансмиси", "20371"),
            (r"бенз.*диз|водата от горивото", "20423"),
            (r"добавк.*диз|injector.*диз|инжектор.*диз|common rail|dab|diesel", "20366"),
            (r"добавк.*бенз|injector.*бенз|инжектор.*бенз", "20367"),
            (r"октан", "20368"), (r"добавк.*масл|димене|компресия|stop.*oil", "20364"),
            (r"радиатор|охладителн", "20363"), (r"масло.*\d+w\d+|масло.*racing", "20416"),
            (r"масло", "20412"), (r"спирач.*почист|почист.*спирач", "20375"),
            (r"обезмаслител", "20380"), (r"egr|турбо", "20379"),
            (r"добавка", "20362")]
        return next((cid for pattern,cid in rules if re.search(pattern,n)), "20360")
    if "phone planet" not in (product["name"] + " " + product["description"]).lower():
        return None
    # Original gallery reviewed on 2026-11-03: HP05 is over-ear (not earbuds);
    # HD01/HD03 have automotive mounts. Preserve these SKU-specific findings.
    photo_categories={"PP-HP05 - GY":None,"HD01":"24269","HD03":"24269"}
    if product.get("sku") in photo_categories:
        return photo_categories[product["sku"]]
    if "стойка" in n and re.search(r"автомобил|шофьор|таблото|предното стъкло", product["description"].lower()):
        return "24269"
    # Mixed accessory sets and audio adapters need an appropriate template/category.
    if re.search(r"\bset\b|\baux\b|3[.,]5\s*мм", n):
        return None
    rules = [(r"видеорегистратор", "22906"), (r"ароматизатор|дифузер", "22059"),
             (r"часовник", "39162"), (r"външна батерия|power\s*bank", "24277"),
             (r"слушал", "24256"), (r"калъф|кейс", "24376"),
             (r"стойка.*автомоб|стойка.*кола|авто.*стойка", "24269"),
             (r"стойка", "24306"), (r"кабел.*lightning|кабел.*iphone|кабел.*лайтнинг", "24323"),
             (r"кабел", "24321"), (r"зарядн.*комплект|комплект.*зарядн", "24282"),
             (r"адаптер.*12\s*v|зарядн.*кола|авто.*зарядн", None),
             (r"адаптер.*220\s*v|зарядн", "24278")]
    for pattern,cid in rules:
        if re.search(pattern,n):
            return cid
    return None


def bardahl_package(p):
    """User-approved shipping heuristic, not measured mass or dimensions."""
    name = p["name"].lower()
    # Known typography errors m1/mI in AccessBG product titles mean ml.
    normalized = re.sub(r"(?<=\d)\s*m[1i]\b", "ml", name)
    matches = list(re.finditer(r"(\d+(?:[.,]\d+)?)\s*(ml|мл|литър|литра|л\b|l\b|gr\b|гр\b|g\b)", normalized))
    source = "title"
    if matches:
        # Prefer explicit ml to model text such as '01L BOOSTER 400ml'.
        match = next((m for m in reversed(matches) if m[2] in ("ml", "мл")), matches[-1])
        value,unit = float(match[1].replace(",", ".")),match[2]
        amount = value * 1000 if unit in ("литър","литра","л","l") else value
        net_unit = "g" if unit in ("gr","гр","g") else "ml"
    else:
        explicit = re.search(r"опаковка\s*[-–:]?\s*(\d+(?:[.,]\d+)?)\s*(мл|ml|л\b|l\b)",p["description"],re.I)
        if explicit:
            amount = float(explicit[1].replace(",", ".")) * (1000 if explicit[2].lower() in ("л","l") else 1)
            net_unit,source = "ml", "description: packaging size"
        else:
            # Visually similar container family; keep the unknown net content blank.
            amount = 1000 if p["sku"] == "BAR-5492" else 400 if p["sku"] == "BAR-3221" else 500
            net_unit,source = "ml", "estimated container class; net content not asserted"
    weight = int(Decimal(str(amount * 1.1)).quantize(Decimal("1"),rounding=ROUND_HALF_UP))
    tiers = [(150,(6,6,12)),(300,(7,7,18)),(600,(8,8,24)),(1000,(10,8,25)),(4000,(22,12,32)),(5000,(24,14,34)),(10000,(30,20,38))]
    dims = next((d for limit,d in tiers if amount<=limit),(35,25,45))
    values = {"OE":weight,"OF":dims[0],"OG":dims[1],"OH":dims[2]}
    if not source.startswith("estimated"):
        values.update({"OM":amount,"ON":amount,"OO":net_unit})
    return values, {"basis":source,"reference_amount":amount,"reference_unit":net_unit,
                    "method":"estimated shipping weight = stated ml or g × 1.1; dimensions estimated by container size"}


def phone_package(p):
    name = p["name"].lower()
    detail = p["description"].lower()
    if "външна батерия" in name:
        weight,dims = (550,(18,11,5)) if re.search(r"20\s*000",name) else (320,(16,10,4))
    elif "часовник" in name:
        weight,dims = 220,(16,11,5)
    elif "видеорегистратор" in name:
        weight,dims = 650,(22,16,9)
    elif "стойка" in name:
        weight,dims = 300,(18,12,8)
    elif "слушал" in name:
        weight,dims = (450,(22,20,10)) if "800mah" in detail or "pp-hp05" in p["sku"].lower() else (70,(12,9,3)) if "хендсфри" in name else (180,(14,11,6))
    elif "кабел" in name:
        weight,dims = (100,(16,11,3)) if "2m" in name else (80,(14,10,3))
    elif "адаптер" in name or "зарядно" in name:
        weight,dims = 180,(12,10,6)
    elif "ароматизатор" in name or "дифузер" in name:
        weight,dims = 180,(12,9,7)
    else:
        weight,dims = 250,(20,15,6)
    # Product body dimensions in a description are not package dimensions.
    return {"OE":weight,"OF":dims[0],"OG":dims[1],"OH":dims[2]}, {"basis":"estimated retail package by product type","method":"package weight and dimensions are estimates; not manufacturer measurements"}


def phone_technical(p,row,template):
    cid = row.get("E")
    if not cid:
        return
    detail = (p["name"]+" "+p["description"]).lower()
    disabled = set(template.modes[cid+"_disabled"])
    def put(col,value):
        if col in disabled:
            return
        allowed = template.allowed(col,cid,row)
        if not allowed or str(value) in allowed:
            row[col]=value
    cable=cid in {"24321","24323"}
    wired=bool(re.search(r"хендсфри|hf0",detail))
    if cable or wired:
        put("DS","USB");put("DZ","Without Battery")
    elif cid=="24278":
        if "220v" in detail or "220 v" in detail or "100–240v" in detail:
            put("DS","Power Supply");put("DZ","Without Battery")
        elif "usb" in detail:
            put("DS","USB");put("DZ","Without Battery")
    elif cid in {"24277","39162"} or (cid=="24256" and not wired and re.search(r"блутут|bluetooth|tws",detail)):
        put("DS","USB Charging");put("DZ","Rechargeable Battery")
    elif cid=="22906" and re.search(r"usb|type.c",detail):
        put("DS","USB")
    if re.search(r"wi.fi",detail):put("EF","With Wi-Fi function")
    elif re.search(r"блутут|bluetooth",detail):put("EF","wireless")
    elif wired:put("EF","none")
    if "безжично" in detail and "зарежд" in detail:put("EG","With wireless charging function")
    voltage=re.search(r"(?<!\d)(5|6|12|24)\s*[vв](?!\w)",detail)
    if row.get("DS")=="Power Supply":
        if re.search(r"100\s*[–-]\s*240\s*v",detail):put("DT","100V (inclusive)-240V (inclusive)")
        elif "220v" in detail or "220 v" in detail:put("DT","220V")
        if re.search(r"\beu\b",detail):put("DV","Type C Plug")
    elif voltage:
        put("DT",voltage[1]+"V")
    powers=re.findall(r"(?<![\w.,])(\d+(?:[.,]\d+)?)\s*(?:w|вт|вата?)\b",detail)
    if powers:put("EZ",max(float(v.replace(",",".")) for v in powers).__format__("g")+"W")
    # Literal source connector descriptions only; no generic connector is invented.
    if re.search(r"type[ -]?c|usb.c|usb-c",detail):put("II","Usbtypec")
    elif "lightning" in detail:put("II","Lightning")
    elif "usb" in detail:put("II","USB")
    for token,value in [(r"\babs\b","ABS"),(r"\bpc\b","PC"),(r"алуминиев.*сплав|алуминиева сплав","Aluminum Alloy"),(r"силикон","Silicone"),(r"алуминиев|алуминий","Aluminum"),(r"медни|медна","Copper")]:
        if re.search(token,detail):
            for c in ["EU","FJ"]:
                put(c,value)
            break
    if cid=="39162":
        if "цинкова сплав" in detail:put("MH","Alloy")
        if "силиконова" in detail:put("MI","Silicone")
        screen=re.search(r"(\d+(?:[.,]\d+)?)\s*(?:-инчов|инч|[\"])",detail)
        if screen:put("MQ",float(screen[1].replace(",",".")));put("MR","Inch")


def build_row(p, config, template, overrides):
    group = "bardahl" if "bardahl" in p["groups"] else "phone_planet"
    settings = {**config.get("defaults", {}), **config.get(group, {})}
    cid = classify(p)
    row = {"E": cid, "F": template.categories.get(cid, ""), "G": "Normal product",
           "L": p["name"][:500], "M": p["sku"], "N": p["sku"],
           "T": p["description"][:20000], "MY": "Model", "NM": p["sku"],
           "OA": p["price"] if settings.get("price_mode") == "current" else p["regular_price"],
           "OC": p["regular_price"], "QV": p["sku"]}
    for i,u in enumerate(p["images"][:11]):
        row[["NO","NP","NQ","NR","NS","NT","NU","NV","NW","NX","NY"][i]] = u
    keys = {"quantity":"NZ", "weight_g":"OE", "length_cm":"OF", "width_cm":"OG", "height_cm":"OH",
            "shipping_template":"OR", "country_of_origin":"OV", "manufacturer":"QW", "eu_responsible_person":"QX",
            "individually_packed":"OJ", "packaging_quantity":"OK", "packaging_unit":"OL"}
    for key,col in keys.items():
        if settings.get(key) is not None:
            row[col] = settings[key]
    if settings.get("estimate_packages"):
        package,method=bardahl_package(p) if group=="bardahl" else phone_package(p)
        row.update(package)
    if group=="phone_planet":
        phone_technical(p,row,template)
    if cid:
        disabled = set(template.modes[cid + "_disabled"])
        row = {c:v for c,v in row.items() if c not in disabled}
    for k,v in overrides.get(p["sku"], {}).items():
        if k not in template.labels:
            raise ValueError("Unknown override column " + k)
        row[k] = v
    if row.get("E"):
        row["F"] = template.categories.get(str(row["E"]), "")
    if group=="bardahl" and settings.get("omit_eu_responsible_person"):
        row.pop("QX",None)
    issues = list(p["issues"])
    if row.get("E") in {"20360", "20362"} and not overrides.get(p["sku"], {}).get("E"):
        issues.append("Review broad category mapping; confirm with per-SKU E override")
    if p["currency"] != "EUR":
        issues.append("Source price currency is not EUR")
    if p["availability"] != "in_stock":
        issues.append("Availability: " + p["availability"])
    # These confirmations are factual and cannot be inferred from a dropdown.
    if group=="phone_planet" and not settings.get("confirm_seller_values", False):
        issues.append("Confirm seller-specific manufacturer / EU RP / origin / shipping values in config.json")
    template_issues=template.validate(row)
    if group=="bardahl" and settings.get("omit_eu_responsible_person"):
        template_issues=["EU Responsible person deliberately blank per user instruction; original template still marks QX required" if issue=="Missing: EU Responsible person (QX)" else issue for issue in template_issues]
    issues.extend(template_issues)
    return row, list(dict.fromkeys(issues))


def run(args):
    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    # Never leave a stale previously validated workbook in a new failed run.
    for name in ["accessbg_VALIDATED.xlsx", "accessbg_REVIEW.xlsx", "validation.csv", "package_estimates.csv", "summary.json", "failures.json", "products.json", "rows.json"]:
        (out/name).unlink(missing_ok=True)
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    overrides = json.loads(Path(args.overrides).read_text(encoding="utf-8"))
    template = Template(args.template)
    client = Client(args.cache)
    counts = {}
    if args.offline_urls:
        scoped = json.loads(Path(args.offline_urls).read_text())
        urls = {u: {g} if isinstance(g,str) else set(g) for u,g in scoped.items()}
    else:
        urls, counts = discover(client)
    failures, products = [], []
    def fetch(item):
        url,groups = item
        try:
            return parse_product(client.get(url, fresh=not bool(args.offline_urls)), url, groups)
        except Exception as exc:
            return {"url":url, "error":str(exc)}
    with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
        for i,p in enumerate(pool.map(fetch,sorted(urls.items())),1):
            (failures if "error" in p else products).append(p)
            if i % 10 == 0:
                LOG.info("Processed %s/%s", i, len(urls))
    rows, ready, audits, sku_counts = [], [], [], defaultdict(int)
    for p in products:
        sku_counts[p["sku"]] += 1
    for p in products:
        row,issues = build_row(p, config, template, overrides)
        if sku_counts[p["sku"]] > 1:
            issues.append("Duplicate source SKU: resolve before import")
        rows.append(row)
        if not issues:
            ready.append(row)
        audits.append({"sku":p["sku"], "name":p["name"], "url":p["url"], "category":row.get("E"),
                       "status":"REVIEW" if issues else "VALIDATED", "issues":" | ".join(issues)})
    # A failed fetch or partial discovery never leads to a nominally complete import.
    if failures or len(products) != len(urls):
        ready = []
    (out/"products.json").write_text(json.dumps(products,ensure_ascii=False,indent=2),encoding="utf-8")
    (out/"rows.json").write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding="utf-8")
    (out/"failures.json").write_text(json.dumps(failures,ensure_ascii=False,indent=2),encoding="utf-8")
    with (out/"validation.csv").open("w",encoding="utf-8-sig",newline="") as f:
        writer=csv.DictWriter(f,fieldnames=["sku","name","url","category","status","issues"])
        writer.writeheader();writer.writerows(audits)
    estimate_rows=[]
    for p,row in zip(products,rows):
        values,method=bardahl_package(p) if "bardahl" in p["groups"] else phone_package(p)
        estimate_rows.append({"sku":p["sku"],"name":p["name"],"weight_g":row.get("OE"),"length_cm":row.get("OF"),"width_cm":row.get("OG"),"height_cm":row.get("OH"),
                              "basis":method["basis"],"method":method["method"],"stock_note":"20 units is an editable publishing assumption, not site stock count"})
    with (out/"package_estimates.csv").open("w",encoding="utf-8-sig",newline="") as f:
        writer=csv.DictWriter(f,fieldnames=["sku","name","weight_g","length_cm","width_cm","height_cm","basis","method","stock_note"])
        writer.writeheader();writer.writerows(estimate_rows)
    summary={"counts":counts,"discovered_urls":len(urls),"parsed_products":len(products),
             "failed_products":len(failures),"validated_rows":len(ready),"review_rows":len(rows)-len(ready),
             "in_stock":sum(p["availability"]=="in_stock" for p in products),
             "complete_crawl":not failures and len(products)==len(urls),
             "note":"VALIDATED checks template fields; it does not certify Temu approval or product compliance."}
    (out/"summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
    if not args.data_only:
        (out/"accessbg_REVIEW.xlsx").write_bytes(template.workbook_bytes(rows))
        if ready:
            (out/"accessbg_VALIDATED.xlsx").write_bytes(template.workbook_bytes(ready))
    LOG.info("Summary: %s",json.dumps(summary,ensure_ascii=False))
    if failures:
        raise SystemExit("Some products failed; see failures.json. No validated file created.")


if __name__ == "__main__":
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template",default="input/template.xlsx")
    parser.add_argument("--config",default="config.json")
    parser.add_argument("--overrides",default="overrides.json")
    parser.add_argument("--output",default="output")
    parser.add_argument("--cache",default=".cache")
    parser.add_argument("--workers",type=int,default=3)
    parser.add_argument("--offline-urls",help="Development only: use scoped URL JSON and cached product pages")
    parser.add_argument("--data-only",action="store_true",help="Export JSON/CSV without writing XLSX")
    logging.basicConfig(level=logging.INFO,format="%(levelname)s %(message)s")
    run(parser.parse_args())
