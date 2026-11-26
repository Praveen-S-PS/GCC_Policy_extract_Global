import hashlib
import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import requests


class ParaphraseDeadline(Exception):
    """Raised by paraphrase_many() when OllamaParaphraser.DEADLINE has passed.
    The cache has already been saved, so running the same command again
    continues where this one stopped."""


class OllamaParaphraser:

    # Absolute time (time.time()) after which paraphrase_many() stops at the
    # next chunk boundary, saves the cache and raises ParaphraseDeadline.
    # None = no limit. Set by country_extractor's --max-minutes.
    DEADLINE = None

    # The cache is written after every SAVE_EVERY items, not only at the end
    # of a call. A run stopped part way (time limit, crash, closed laptop)
    # used to lose every paraphrase done since its last completed call.
    SAVE_EVERY = 8

    HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")

    DEFAULT_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2:3b")

    CACHE_PATH = "output/.paraphrase_cache.json"

    # Only paraphrase substantial paragraphs; short bullets pass through as-is
    # to save 50-70% of LLM calls with minimal quality loss.
    MIN_LEN = 250

    REQUEST_TIMEOUT = 120

    # Bump PROMPT_VERSION whenever PROMPT_TEMPLATE changes. It gets folded
    # into the cache key so old cached responses (produced under a previous
    # prompt) are no longer used and are regenerated with the new prompt.
    #
    # v3 note: v2 opened with "...for a Global Capability Centre (GCC)
    # reference workbook" as the very first thing the model read. Small
    # models (llama3.2:3b) latched onto that framing and started inventing
    # "The Global Capability Centre requires/will focus on/..." as the
    # subject of the paraphrase, even when the source text never mentioned
    # a GCC at all (verified against raw source text for Colorado and
    # Alabama state incentive pages: zero occurrences of "global capability"
    # in either raw page, but the paraphrase fabricated it in several
    # bullets). A second round of testing found the model also conflates any
    # department name that merely contains the word "Global" (e.g. the real
    # Colorado "Global Business Development" division) with a GCC about 40%
    # of the time, even with an explicit anti-fabrication rule -- prompting
    # alone could not fully eliminate this on a 3B model, so paraphrase()
    # below also runs a deterministic post-check (_domain_terms_changed)
    # that discards any paraphrase whose domain-term set doesn't match the
    # input's, in either direction, falling back to the original text.
    PROMPT_VERSION = "v3"

    PROMPT_TEMPLATE = (
        "You are paraphrasing a single clause from a government policy or "
        "economic development document. The paraphrase will be stored in a "
        "research database that tracks incentives potentially relevant to "
        "corporate shared-services and technology operations centers, also "
        "known as Global Capability Centres (GCCs) -- but the input clause "
        "itself may be about a general business, tax, or investment program "
        "that never mentions a GCC directly.\n\n"
        "Rules:\n"
        "- Output a single short paragraph in clear English, 1 to 3 sentences.\n"
        "- Copy these EXACT phrases verbatim into your output whenever they "
        "appear in the input, character for character, never paraphrased "
        "into a synonym: \"Global Capability Centre\", \"Global Capability "
        "Center\", \"GCC\", \"capability centre\", \"capability center\", "
        "\"GBS\", \"BSS\", \"BPO\", \"KPO\", \"ITES\", \"IT/ITeS\", \"ITBPM\", "
        "\"shared services\", \"business services sector\", \"IT services\", "
        "\"R&D centre\", \"R&D center\", \"IP Box\", \"patent box\", "
        "\"innovation hub\", \"centre of excellence\", \"captive centre\", "
        "\"offshore\", \"nearshore\", \"SEZ\", \"EPZ\", \"free zone\", "
        "\"IT park\", \"technology park\", \"technology centre\".\n"
        "- Do NOT write \"Global Capability Centre\" or \"GCC\" unless one of "
        "those exact phrases appears in the input. A department or program "
        "whose name merely contains the word \"Global\" for an unrelated "
        "reason (for example \"Global Business Development\", \"Global Trade "
        "Office\") is NOT a Global Capability Centre -- keep its real name "
        "unchanged and do not relabel it as a GCC.\n"
        "- Never invent a sentence subject the input does not name -- if the "
        "input is about a general program, tax credit, agency, or "
        "requirement with no GCC-specific wording, keep describing it as "
        "that program/tax credit/agency/requirement.\n"
        "- Preserve all numbers, percentages, currency amounts, and time "
        "periods exactly as in the input.\n"
        "- Preserve city, state, region, and country names verbatim.\n"
        "- Drop legal boilerplate, navigation labels, table headers, and "
        "section codes.\n"
        "- If the input is in Spanish, French, Portuguese, Polish, or "
        "another language, translate to English as part of the paraphrase, "
        "and STILL preserve the domain terms above and place names.\n"
        "- If the input is already concise and clear, return it unchanged.\n"
        "- Output the paraphrase only, no preamble, no markdown, no quotes.\n\n"
        "Input:\n{text}\n\nParaphrase:"
    )

    # Deterministic backstop for the domain-term preservation rule above.
    # Prompting alone was not reliable enough on a 3B model in either
    # direction: testing found the model both invents GCC mentions that
    # weren't in the source (e.g. conflating "Global Business Development",
    # a real unrelated division name, with "GCC") AND, separately but just
    # as consistently (5/5 in testing), drops a genuine "(GCC)" / "Global
    # Capability Centre" mention that WAS in the source down to a generic
    # "eligible centres". Both are content-integrity failures, so every one
    # of these terms is checked for presence in the input and the output;
    # any mismatch in either direction discards the paraphrase in favor of
    # the original text rather than caching an unfaithful rewrite.
    _DOMAIN_TERM_RES = [
        re.compile(rf"\b{re.escape(t)}s?\b", re.IGNORECASE)
        for t in (
            "global capability centre", "global capability center", "gcc",
            "capability centre", "capability center",
            "gbs", "bss", "bpo", "kpo", "ites", "it/ites", "itbpm",
            "shared service", "business services sector",
            "business services industry", "it service",
            "r&d centre", "r&d center",
            "ip box", "patent box", "innovation hub",
            "centre of excellence", "center of excellence",
            "captive centre", "captive center",
            "offshore", "nearshore", "sez", "epz", "free zone",
            "it park", "technology park", "technology centre",
            "technology center",
        )
    ]

    # Concurrent requests to Ollama. Local Ollama typically handles 4 to 8
    # parallel generate calls without slowdown; going higher can bottleneck
    # the model. Overridable via OLLAMA_WORKERS env var.
    PARALLEL_WORKERS = int(os.environ.get("OLLAMA_WORKERS", "8"))

    def __init__(self, model=None, host=None, min_len=None):

        self.model = model or self.DEFAULT_MODEL
        self.host = host or self.HOST
        self._cache = self._load_cache()
        self._dirty = False
        self._available = None
        self._cache_lock = threading.Lock()

        # Per-instance override of the class-level MIN_LEN default, e.g. to
        # force every bullet through paraphrasing for a specific run
        # without changing the default for other callers.
        if min_len is not None:
            self.MIN_LEN = min_len

    def is_available(self):

        if self._available is not None:
            return self._available

        try:

            r = requests.get(f"{self.host}/api/tags", timeout=3)

            if r.status_code != 200:
                self._available = False
                return False

            tags = r.json().get("models", [])

            names = {t.get("name", "") for t in tags}

            if self.model not in names:

                print(
                    f"⚠ Ollama is running but model '{self.model}' is not "
                    f"installed. Run: ollama pull {self.model}"
                )

                self._available = False
                return False

            self._available = True
            return True

        except (requests.RequestException, ValueError):

            self._available = False
            return False

    def paraphrase(self, text):

        if not text or len(text) < self.MIN_LEN:
            return text

        text = text.strip()

        key = self._cache_key(text)

        # Read is safe without lock; dict get is atomic in CPython
        cached = self._cache.get(key)
        if cached is not None:
            return cached

        if not self.is_available():
            return text

        try:

            r = requests.post(
                f"{self.host}/api/generate",
                json={
                    "model": self.model,
                    "prompt": self.PROMPT_TEMPLATE.format(text=text),
                    "stream": False,
                    "options": {
                        "temperature": 0.2,
                        "num_predict": 350,
                    },
                },
                timeout=self.REQUEST_TIMEOUT,
            )

            r.raise_for_status()

            out = r.json().get("response", "").strip()

            if not out or len(out) < 20:

                return text

            if self._domain_terms_changed(text, out):

                # The model either invented domain wording that wasn't in
                # the source, or dropped domain wording that was (a known
                # llama3.2:3b failure mode in both directions -- see the
                # PROMPT_VERSION note above). The raw text is more
                # trustworthy than an unfaithful paraphrase, so use it
                # instead of caching the mismatch.
                out = text

            with self._cache_lock:
                self._cache[key] = out
                self._dirty = True

            return out

        except (requests.RequestException, ValueError):

            return text

    def paraphrase_many(self, items, label=""):

        if not items:
            return []

        if not self.is_available():
            return items

        total = len(items)

        if label:
            print(f"  paraphrase {label} [{total} items, "
                  f"{self.PARALLEL_WORKERS} workers]")

        out = []

        for i in range(0, total, self.SAVE_EVERY):

            batch = items[i:i + self.SAVE_EVERY]

            if self.DEADLINE is not None and time.time() > self.DEADLINE and any(
                self._needs_model(t) for t in batch
            ):
                self._save_cache()
                raise ParaphraseDeadline(
                    f"time limit reached in {label or 'paraphrasing'} "
                    f"after {i} of {total} items"
                )

            # ThreadPoolExecutor.map preserves input order.
            with ThreadPoolExecutor(max_workers=self.PARALLEL_WORKERS) as ex:
                out.extend(ex.map(self.paraphrase, batch))

            self._save_cache()

        return out

    def _needs_model(self, text):
        """True when paraphrase(text) would have to call the model."""

        if not text or len(text) < self.MIN_LEN:
            return False

        return self._cache.get(self._cache_key(text.strip())) is None

    def _domain_terms_changed(self, original, paraphrased):
        """True when the set of domain terms present differs between the
        original and the paraphrase, in either direction: a term appearing
        in the paraphrase that wasn't in the original (invented), or a term
        in the original that's missing from the paraphrase (dropped)."""

        for pat in self._DOMAIN_TERM_RES:

            if bool(pat.search(original)) != bool(pat.search(paraphrased)):
                return True

        return False

    def _cache_key(self, text):

        h = hashlib.sha1()
        h.update(self.model.encode("utf-8"))
        h.update(b"\n")
        h.update(self.PROMPT_VERSION.encode("utf-8"))
        h.update(b"\n")
        h.update(text.encode("utf-8"))

        return h.hexdigest()

    def _load_cache(self):

        if not os.path.exists(self.CACHE_PATH):
            return {}

        try:

            with open(self.CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)

        except (OSError, json.JSONDecodeError):

            return {}

    def _save_cache(self):

        with self._cache_lock:

            if not self._dirty:
                return

            os.makedirs(os.path.dirname(self.CACHE_PATH), exist_ok=True)

            snapshot = dict(self._cache)

            self._dirty = False

        # Write to a temp file and rename: a process killed mid write must
        # not leave a truncated cache, which loads as empty and would throw
        # away every saved paraphrase.
        tmp_path = self.CACHE_PATH + ".part"

        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, ensure_ascii=False, indent=2)

        os.replace(tmp_path, self.CACHE_PATH)
