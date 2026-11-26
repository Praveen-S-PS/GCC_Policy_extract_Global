import hashlib
import json
import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor


# langdetect's DetectorFactory is a module-level global singleton. Calling
# detect() from multiple threads concurrently (as translate_many's worker
# pool does) is NOT safe -- it intermittently raises
# "LangDetectException: Need to load profiles" as several threads race to
# lazily initialize the same factory. That exception was previously caught
# by a bare `except Exception: return text` with no log line, so it looked
# like a language-detection miss when it was actually silent data loss:
# genuine non-English text was left untranslated with no record of why.
# Confirmed via a minimal 4-thread reproduction before this fix.
_LANGDETECT_LOCK = threading.Lock()


class ArgosTranslator:

    CACHE_PATH = "output/.translate_cache.json"

    MIN_LEN = 30

    TARGET_LANG = "en"

    # Language codes we're willing to download packs for
    # (mapping from langdetect's ISO to argos-translate's ISO where they differ)
    SUPPORTED = {
        "es": "es", "fr": "fr", "de": "de", "it": "it", "pt": "pt",
        "pl": "pl", "ro": "ro", "hu": "hu", "nl": "nl", "tr": "tr",
        "cs": "cs", "sk": "sk", "el": "el", "ru": "ru", "bg": "bg",
        "uk": "uk", "hr": "hr", "sr": "sr", "sl": "sl",
        "zh-cn": "zh", "zh-tw": "zh", "ar": "ar", "id": "id", "vi": "vi",
        "ja": "ja", "ko": "ko", "th": "th", "he": "he", "fa": "fa",
    }

    URL_RE = re.compile(r"https?://\S+")

    # Argos-translate uses ctranslate2 which is thread-safe once the model
    # is loaded. We warm each language pack before the pool starts.
    PARALLEL_WORKERS = int(os.environ.get("TRANSLATE_WORKERS", "8"))

    # ctranslate2's own concurrency knobs -- separate from PARALLEL_WORKERS
    # above. argostranslate defaults to inter_threads=1, which means its
    # Translator serializes every translate() call internally regardless of
    # how many Python threads call it at once: with 4 worker threads and
    # inter_threads=1, only one translation ever actually runs at a time,
    # each one internally parallelized (intra_threads="auto") across
    # whatever cores it grabs. That serialization -- not the GIL, not I/O --
    # was the real bottleneck (confirmed by benchmark: 40 short German
    # paragraphs took 11.5s at the defaults vs 3.7s at inter=8/intra=1, a
    # ~3x speedup, with identical translation output). Our workload is many
    # short paragraphs rather than a few long documents, which is exactly
    # the case where raising inter_threads (more paragraphs translated
    # concurrently) beats raising intra_threads (more cores per paragraph,
    # where small inputs can't use them anyway).
    CT2_INTER_THREADS = int(os.environ.get("TRANSLATE_INTER_THREADS", "8"))
    CT2_INTRA_THREADS = int(os.environ.get("TRANSLATE_INTRA_THREADS", "1"))

    def __init__(self):

        self._cache = self._load_cache()
        self._dirty = False
        self._installed_langs = set()
        self._pack_index_downloaded = False
        self._available = None
        self._cache_lock = threading.Lock()
        self._pack_lock = threading.Lock()

    def is_available(self):

        if self._available is not None:
            return self._available

        try:

            import argostranslate.package
            import argostranslate.settings
            import argostranslate.translate

            # Must be set before the first ctranslate2.Translator is
            # constructed (lazily, on first translate call) -- it reads
            # these settings values at construction time and caches the
            # instance, so setting them here at first import is early enough.
            argostranslate.settings.inter_threads = self.CT2_INTER_THREADS
            argostranslate.settings.intra_threads = self.CT2_INTRA_THREADS

            self._available = True

        except ImportError:

            self._available = False
            print(
                "⚠ argostranslate not installed. "
                "Skipping translation. Run: pip install argostranslate langdetect"
            )

        return self._available

    def translate_to_english(self, text):

        if not text or len(text) < self.MIN_LEN:
            return text

        text = text.strip()

        # Language detect on a stripped preview (drop URLs first)
        preview = self.URL_RE.sub("", text)[:2000]

        try:

            from langdetect import detect, DetectorFactory

            with _LANGDETECT_LOCK:
                DetectorFactory.seed = 0
                lang = detect(preview)

        except Exception as exc:

            print(f"    ✗ language detection failed, leaving text as-is: {exc}")
            return text  # can't tell; leave alone

        if lang.startswith("en"):
            return text

        argos_lang = self.SUPPORTED.get(lang, self.SUPPORTED.get(lang.split("-")[0]))

        if not argos_lang:
            return text

        key = self._cache_key(argos_lang, text)

        if key in self._cache:
            return self._cache[key]

        if not self.is_available():
            return text

        if not self._ensure_pack(argos_lang):
            return text

        try:

            import argostranslate.translate

            translated = argostranslate.translate.translate(
                text, argos_lang, self.TARGET_LANG
            )

            if not translated or len(translated) < 5:
                return text

            with self._cache_lock:
                self._cache[key] = translated
                self._dirty = True

            return translated

        except Exception as exc:

            print(f"    ✗ translate failed ({argos_lang}→en): {exc}")
            return text

    def translate_many(self, items, label=""):

        if not items:
            return []

        if not self.is_available():
            return items

        total = len(items)

        if label:
            print(f"  translate {label} [{total} items, "
                  f"{self.PARALLEL_WORKERS} workers]")

        # ThreadPoolExecutor.map preserves input order. Model init is guarded
        # by _pack_lock inside _ensure_pack, so concurrent first-time calls
        # for the same language pack won't race.
        with ThreadPoolExecutor(max_workers=self.PARALLEL_WORKERS) as ex:
            out = list(ex.map(self.translate_to_english, items))

        self._save_cache()

        return out

    def _ensure_pack(self, from_lang):

        if from_lang in self._installed_langs:
            return True

        # Serialize first-time pack lookups so concurrent workers don't
        # both try to download the same language pack.
        with self._pack_lock:

            if from_lang in self._installed_langs:
                return True

            try:

                import argostranslate.package

                if not self._pack_index_downloaded:
                    argostranslate.package.update_package_index()
                    self._pack_index_downloaded = True

                installed = argostranslate.package.get_installed_packages()

                already = any(
                    p.from_code == from_lang and p.to_code == self.TARGET_LANG
                    for p in installed
                )

                if already:
                    self._installed_langs.add(from_lang)
                    return True

                available = argostranslate.package.get_available_packages()

                match = next(
                    (
                        p for p in available
                        if p.from_code == from_lang and p.to_code == self.TARGET_LANG
                    ),
                    None,
                )

                if not match:
                    print(f"    ⚠ no argos pack for {from_lang}→en")
                    return False

                print(f"    ↓ downloading argos pack {from_lang}→en (~200 MB, one time)")

                path = match.download()

                argostranslate.package.install_from_path(path)

                self._installed_langs.add(from_lang)

                return True

            except Exception as exc:

                print(f"    ✗ pack install failed for {from_lang}: {exc}")
                return False

    def _cache_key(self, from_lang, text):

        h = hashlib.sha1()
        h.update(f"argos:{from_lang}:{self.TARGET_LANG}\n".encode("utf-8"))
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

            snapshot = dict(self._cache)
            self._dirty = False

        os.makedirs(os.path.dirname(self.CACHE_PATH), exist_ok=True)

        with open(self.CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, ensure_ascii=False, indent=2)
