import hashlib
import os
import re
import time
from urllib.parse import urljoin, urlparse

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class PolicyHTMLReader:

    # Different state government WAFs disagree on what looks like a bot.
    # esd.ny.gov and mississippi.org 403 a full Chrome UA string but serve
    # real content to a bare "Mozilla/5.0"; business.delaware.gov does the
    # opposite (its F5 WAF returns a "Request Rejected" page, HTTP 200, for
    # the bare UA but serves real content for the Chrome string). There is
    # no single UA that satisfies both, so _fetch() tries each in turn and
    # treats a detected block/challenge page as a reason to try the next
    # one rather than accepting it as real content.
    USER_AGENT = (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )

    USER_AGENTS = (USER_AGENT, "Mozilla/5.0")

    # Substrings seen in real WAF/bot-challenge block pages during testing
    # (F5 ASM on business.delaware.gov, Cloudflare on esd.ny.gov, Radware
    # on mn.gov), plus common generic phrasing from other vendors.
    _BLOCK_PAGE_SIGNATURES = (
        "request rejected", "access denied", "just a moment",
        "attention required", "captcha", "are you a robot",
        "verify you are a human", "bot manager", "unusual traffic",
        "please enable javascript and cookies", "checking your browser",
    )

    TIMEOUT = 45

    # NOTE: "form" is deliberately NOT in this list. Some legacy ASP.NET
    # (.aspx) state government sites wrap their entire page body in a
    # single <form> for postback support (confirmed on Alaska's, Indiana's,
    # and New Hampshire's incentive pages) -- dropping <form> there deleted
    # all real content, leaving only nav/footer scraps.
    DROP_TAGS = ("script", "style", "noscript", "nav", "header",
                 "footer", "aside", "iframe")

    MIN_PARA_CHARS = 60

    WHITESPACE_RE = re.compile(r"\s+")

    # Anchor text or URL substrings that suggest an incentive detail link.
    INTERESTING_KEYWORDS = [
        "incentive", "incentives", "tax", "subsidy", "subsidies",
        "grant", "grants", "scheme", "schemes", "benefit", "benefits",
        "policy", "policies", "guide", "guides", "handbook",
        "sez", "free zone", "freezone", "special economic",
        "program", "programme", "framework",
        "investment", "invest ", "investor",
        "law", "act", "regulation",
        "rebate", "waiver", "exemption", "deduction",
        "package", "brochure", "download",
        "capital", "gcc", "bpo", "ites", "shared services",
        "beneficio", "incentivo", "exencion", "subvencion",
        "impuesto", "inversion", "guia",
    ]

    # Extensions we prefer to skip when scanning links
    SKIP_EXT = (".jpg", ".jpeg", ".png", ".gif", ".svg", ".mp4",
                ".mp3", ".zip", ".rar", ".mov", ".webp", ".css",
                ".js", ".ico", ".woff", ".woff2")

    # URL path substrings we refuse to follow even if the anchor text matches
    # an INTERESTING_KEYWORDS pattern. These trees on IPA sites are usually
    # consumer, individual, HR, or PR content that isn't a GCC incentive.
    URL_PATH_BLACKLIST = (
        "/consumer", "/citizen", "/individual", "/personal",
        "/passport", "/immigration/personal", "/visa/tourist",
        "/driving", "/driver-licen", "/driver-licence",
        "/birth", "/marriage", "/pension", "/retirement",
        "/unemployment", "/welfare", "/social-security",
        "/labour-rights", "/labor-rights", "/worker-rights",
        "/press-release", "/press_release", "/press-releases", "/press/",
        "/news/", "/news-", "/blog", "/media-centre", "/media-center",
        "/gallery", "/photo", "/video",
        "/success-stories", "/success-story", "/case-study", "/case-studies",
        "/events/", "/annual-report",
        "/careers", "/vacancy", "/vacancies", "/jobs/",
        "/about-us", "/about_us", "/aboutus",
        "/contact", "/faq",
        "/terms", "/privacy", "/disclaimer", "/sitemap",
        "/nomad", "/digital-nomad", "/tourist",
        "/consultation-response",
    )

    MAX_LINKS_PER_PAGE = 8
    MAX_PDF_SIZE_BYTES = 40 * 1024 * 1024  # 40 MB safety cap

    @classmethod
    def read(cls, url, country_dir=None, depth=1, visited=None,
             pdf_reader=None):
        """Fetch a URL and optionally follow one hop of same-domain links.

        Args:
            url: seed URL.
            country_dir: local folder where discovered PDFs get cached.
            depth: how many hops to follow (0 = just the seed).
            visited: set of already-fetched URLs for this crawl.
            pdf_reader: optional PolicyPDFReader class for PDF extraction.
        """

        if visited is None:
            visited = set()

        if url in visited:
            return []

        visited.add(url)

        response, ctype = cls._fetch(url)

        if response is None:
            return []

        if "pdf" in ctype.lower() or url.lower().split("?")[0].endswith(".pdf"):

            paragraphs = cls._handle_pdf_response(
                url, response, country_dir, pdf_reader
            )

            return paragraphs

        # HTML path
        html = response.text if hasattr(response, "text") else str(response)

        paragraphs = cls._extract_paragraphs(html)

        if depth <= 0:
            return paragraphs

        # html.parser (stdlib) rather than lxml: testing against a Drupal
        # page with slightly malformed markup found lxml mis-nesting the
        # real <div class="field--type-text-long"><p>...</p></div> content
        # inside an earlier, unrelated tag that then got dropped along with
        # nav/header/footer -- html.parser produced identical output to
        # lxml on every page that already worked, and recovered the content
        # lxml was losing on the malformed one.
        soup = BeautifulSoup(html, "html.parser")

        candidates = cls._candidate_links(soup, base_url=url)

        for link in candidates[: cls.MAX_LINKS_PER_PAGE]:

            if link in visited:
                continue

            child_paras = cls.read(
                link,
                country_dir=country_dir,
                depth=depth - 1,
                visited=visited,
                pdf_reader=pdf_reader,
            )

            paragraphs.extend(child_paras)

        return cls._dedupe(paragraphs)

    # Why the last fetch of each URL failed ("HTTP 403", "bot challenge
    # page", ...). read() returns [] for a failed fetch, so without this a
    # blocked, dead or timed out source is indistinguishable from an empty
    # one; callers use it to say WHY a source returned nothing.
    FAILURES = {}

    # HTTP 5xx is retried; 4xx, timeouts and connection errors are not
    # (they are not transient, and retrying a dead site would multiply the
    # wait). Paraguay's only source answered 502, 502, 200 in three
    # consecutive requests; with no retry a single bad gateway left the
    # country with zero rows for the whole run.
    SERVER_ERROR_RETRIES = 2

    @classmethod
    def _get(cls, url, headers, verify):

        for attempt in range(cls.SERVER_ERROR_RETRIES + 1):

            resp = requests.get(
                url, headers=headers,
                timeout=cls.TIMEOUT,
                allow_redirects=True,
                verify=verify,
                stream=True,
            )

            if resp.status_code < 500 or attempt == cls.SERVER_ERROR_RETRIES:
                return resp

            resp.close()
            time.sleep(2 * (attempt + 1))

    @classmethod
    def _fetch(cls, url):

        reason = "no response"

        for ua in cls.USER_AGENTS:

            headers = {
                "User-Agent": ua,
                "Accept": "text/html,application/pdf,application/xhtml+xml,*/*",
            }

            for verify in (True, False):

                try:

                    resp = cls._get(url, headers, verify)

                    if resp.status_code >= 400:
                        reason = f"HTTP {resp.status_code}"
                        break  # try the next UA

                    ctype = resp.headers.get("Content-Type", "").lower()

                    # Only sniff text-ish responses for a block/challenge
                    # page. A PDF must not be read via .text -- that would
                    # download and decode the whole binary for nothing.
                    if "pdf" not in ctype and cls._looks_like_block_page(resp):
                        reason = "bot challenge page"
                        break  # try the next UA

                    cls.FAILURES.pop(url, None)

                    return resp, ctype

                except requests.exceptions.SSLError:
                    reason = "SSL error"
                    if verify:
                        continue
                    break  # try the next UA

                except requests.RequestException as exc:
                    reason = type(exc).__name__
                    break  # try the next UA

        cls.FAILURES[url] = reason

        return None, ""

    @classmethod
    def _looks_like_block_page(cls, resp):

        try:
            sample = resp.text[:4000].lower()
        except Exception:
            return False

        return any(sig in sample for sig in cls._BLOCK_PAGE_SIGNATURES)

    @classmethod
    def _handle_pdf_response(cls, url, response, country_dir, pdf_reader):

        if pdf_reader is None or not country_dir:
            return []

        try:
            content_length = int(response.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0

        if content_length and content_length > cls.MAX_PDF_SIZE_BYTES:
            return []

        discovered_dir = os.path.join(country_dir, "discovered")

        os.makedirs(discovered_dir, exist_ok=True)

        slug = cls._slug_from_url(url)

        dest = os.path.join(discovered_dir, f"{slug}.pdf")

        if not os.path.exists(dest):

            try:

                written = 0

                with open(dest, "wb") as f:

                    for chunk in response.iter_content(chunk_size=64 * 1024):

                        if chunk:
                            f.write(chunk)
                            written += len(chunk)

                            if written > cls.MAX_PDF_SIZE_BYTES:
                                break

                # Validate PDF magic
                with open(dest, "rb") as f:
                    if f.read(4) != b"%PDF":
                        os.remove(dest)
                        return []

                print(f"    + discovered PDF: {slug}.pdf ({written // 1024} KB)")

            except Exception:
                if os.path.exists(dest):
                    try:
                        os.remove(dest)
                    except OSError:
                        pass
                return []

        try:
            return pdf_reader.read(dest)
        except Exception:
            return []

    @classmethod
    def _extract_paragraphs(cls, html):

        if not html:
            return []

        # See the parser choice note in read() above.
        soup = BeautifulSoup(html, "html.parser")

        for tag in soup(cls.DROP_TAGS):
            tag.decompose()

        paragraphs = []

        for el in soup.find_all(["p", "li", "td", "h2", "h3", "h4"]):

            text = el.get_text(separator=" ", strip=True)

            text = cls.WHITESPACE_RE.sub(" ", text).strip()

            if len(text) < cls.MIN_PARA_CHARS:
                continue

            paragraphs.append({
                "heading": "",
                "text": text,
            })

        return cls._dedupe(paragraphs)

    @classmethod
    def _candidate_links(cls, soup, base_url):

        base = urlparse(base_url)

        base_host = base.netloc.lower()

        base_reg = ".".join(base_host.split(".")[-2:]) if base_host else ""

        candidates = []

        seen = set()

        for a in soup.find_all("a", href=True):

            href = a.get("href", "").strip()

            if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
                continue

            absolute = urljoin(base_url, href)

            parsed = urlparse(absolute)

            if parsed.scheme not in ("http", "https"):
                continue

            path_lower = parsed.path.lower()

            if any(path_lower.endswith(ext) for ext in cls.SKIP_EXT):
                continue

            if any(bad in path_lower for bad in cls.URL_PATH_BLACKLIST):
                continue

            # Same registrable domain only
            host = parsed.netloc.lower()

            host_reg = ".".join(host.split(".")[-2:]) if host else ""

            if host_reg != base_reg:
                continue

            anchor_text = a.get_text(separator=" ", strip=True).lower()

            match_text = f"{path_lower} {anchor_text}"

            if not any(kw in match_text for kw in cls.INTERESTING_KEYWORDS):
                continue

            if absolute in seen:
                continue

            seen.add(absolute)

            is_pdf = path_lower.endswith(".pdf")

            # How many distinct keywords this link matches -- a page whose
            # URL/anchor text hits several ("competitive-incentives" +
            # "job-development-investment-grant") is much more likely to be
            # real program content than one that just barely qualifies.
            # This matters because on a navigation-heavy hub page there can
            # be far more candidates than MAX_LINKS_PER_PAGE; picking the
            # first few alphabetically (the old behavior) tended to grab
            # weakly-related pages while missing the actual program pages
            # further down the list.
            keyword_hits = sum(1 for kw in cls.INTERESTING_KEYWORDS if kw in match_text)

            candidates.append((is_pdf, keyword_hits, absolute))

        # PDFs first, then by keyword strength (descending), then URL for
        # a stable tie-break.
        candidates.sort(key=lambda t: (not t[0], -t[1], t[2]))

        return [c[2] for c in candidates]

    @staticmethod
    def _dedupe(paragraphs):

        seen = set()
        out = []

        for p in paragraphs:

            key = p["text"][:80]

            if key in seen:
                continue

            seen.add(key)
            out.append(p)

        return out

    @staticmethod
    def _slug_from_url(url):

        parsed = urlparse(url)

        base = os.path.basename(parsed.path) or "document"

        base = base.replace(".pdf", "").strip("/")

        base = re.sub(r"[^A-Za-z0-9._-]+", "_", base)[:60] or "document"

        digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:8]

        return f"{base}_{digest}"
