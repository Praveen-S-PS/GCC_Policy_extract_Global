import re


class PolicyBucketer:

    INCENTIVE_TYPES = [
        "Financial Incentives",
        "Infrastructure",
        "Talent",
        "Regulatory, Ease-of-Doing-Business & Non-financial Support",
        "Innovation / R&D / Patent",
        "Others",
    ]

    SCOPES = [
        "Across the state (Common)",
        "Specific to Tier 1 cities",
        "Specific to Tier 2 & 3 cities",
    ]

    KEYWORD_MAP = {
        "Financial Incentives": [
            # English core
            "capex", "opex", "capital subsidy", "capital support",
            "capital incentive", "stamp duty", "registration fee",
            "interest subsidy", "interest assistance", "tax exemption",
            "tax holiday", "tax break", "tax rebate", "tax refund",
            "tax deduction", "tax deferral", "tax credit", "tax relief",
            "sgst", "vat exemption", "vat refund", "income tax exemption",
            "income tax reduction", "corporate income tax", "cit reduction",
            "payroll subsidy", "salary subsidy",
            "financial assistance", "fiscal incentive", "reimbursement of",
            "concession", "subsidy", "grant", "cash grant", "cash back",
            "withholding tax exemption", "withholding tax reduction",
            "super deduction", "accelerated depreciation", "loss carry-forward",
            "customs duty exemption", "customs duty rebate",
            "waiver of", "duty exemption", "duty rebate",
            "employment incentive", "employment bonus", "employment grant",
            "training grant", "R&D grant", "innovation grant",
            "investment grant", "investment allowance", "capital allowance",
            "monetary incentive", "cash incentive", "financial support",
            "ip box", "patent box", "innovation box",
            # Headquarters family (2026-09 review: "headquarters" matched
            # nowhere in the lists, likely why the Saudi RHQ Programme,
            # Malaysia Principal Hub, Thailand IBC, Panama SEM, and several
            # US state HQ relocation credits were all missed on first pass)
            "headquarters", "regional headquarters", "rhq", "principal hub",
            "operational headquarters", "international business centre",
            "relocation",
            # Jobs family (most common US state-level incentive per the
            # review; "job" itself matched nowhere in the lists)
            "job creation", "jobs credit", "per new job", "payroll rebate",
            "withholding retention", "wage rebate", "cash rebate", "rebate",
            "refundable credit", "transferable credit", "sale and assignment",
            "abatement", "tax abatement", "payment in lieu of taxes",
            "tax increment financing", "tif", "industrial development bond",
            "industrial revenue bond", "forgivable loan", "repayable advance",
            "recoverable advance", "soft loan", "matching grant",
            "closing fund", "deal closing", "strategic fund",
            "discretionary fund", "enterprise zone", "opportunity zone",
            "freeport", "free zone", "free trade zone",
            "export processing zone", "special tax site", "investment zone",
            "first-year allowance", "full expensing", "bonus depreciation",
            "notional interest deduction",
            "social security contribution relief", "employer contribution relief",
            "national insurance relief",
            # Spanish
            "subsidio", "subvención", "exención fiscal", "deducción fiscal",
            "crédito fiscal", "reembolso", "bonificación", "exoneración",
            "impuesto sobre la renta", "isr", "iva", "itbis",
            "incentivo fiscal", "beneficio fiscal", "reducción impositiva",
            "estímulo fiscal", "exento", "tasa cero",
            "devolución", "aporte no reembolsable",
            # French (2026-09 review: France file came out thin because the
            # instruments are named only in French; these terms are meant to
            # match untranslated scheme names/vocabulary that survive
            # machine translation as proper nouns)
            "crédit d'impôt", "crédit d'impôt recherche",
            "crédit d'impôt collaboration de recherche", "régime des impatriés",
            "exonération", "allègement", "subvention", "aide", "prime",
            "avance récupérable", "agrément", "zone franche",
            "pôle de compétitivité", "suramortissement", "dégrèvement",
            # German (same reasoning as French, above)
            "forschungszulage", "investitionszulage", "grw zuschuss",
            "zuschuss", "förderung", "förderrichtlinie", "zulage", "beihilfe",
            "steuerbefreiung", "steuerermäßigung", "sonderabschreibung",
            "bürgschaft", "darlehen",
            # Japanese (same reasoning; Japan file came out thinnest of all)
            "研究開発税制", "地方拠点強化税制", "情報通信産業特別地区",
            "税制", "優遇措置", "補助金", "助成金", "交付金", "控除",
            "減税", "特区", "立地",
            # Polish
            "dotacja", "ulga podatkowa", "zwolnienie", "wsparcie",
            "pomoc publiczna",
            # Portuguese
            "incentivo", "benefício fiscal", "isenção", "crédito fiscal",
            "apoio", "majoração",
        ],
        "Infrastructure": [
            # English
            "electricity duty", "power tariff", "land subsidy", "land rebate",
            "land cost", "rental subsidy", "lease rental", "rental assistance",
            "bandwidth", "internet", "cloud rental", "gcc park", "it park",
            "gift city", "infrastructure", "right of way", "fibre", "fiber",
            "co-working", "office space", "industrial estate", "data centre",
            "data center", "plug and play", "broadband",
            # English (2026-09 review additions)
            "site readiness", "shovel ready", "certified site", "turnkey site",
            "brownfield", "greenfield", "land remediation",
            "utility connection", "grid connection",
            "structures and buildings allowance",
            # French
            "clés en main",
            # Spanish
            "zona económica", "zona franca", "parque industrial",
            "infraestructura", "electricidad", "energía", "banda ancha",
            "oficina", "alquiler", "arrendamiento", "suelo", "terreno",
            "telecomunicaciones",
        ],
        "Talent": [
            # English
            "skilling", "training", "internship", "academia", "upskill",
            "atmanirbhar", "rojgar", "employee provident fund",
            "employees' provident fund", "epf", "employment generation",
            "recruitment", "talent", "workforce", "human capital",
            # Immigration and mobility (2026-09 review: no visa/permit/
            # mobility vocabulary existed anywhere in the lists, which likely
            # explains why the whole immigration layer was missed -- UK
            # Global Business Mobility/Skilled Worker, French Passeport
            # Talent, Japanese Highly Skilled Professional system)
            "visa", "work permit", "residence permit", "immigration",
            "sponsorship", "certificate of sponsorship", "skilled worker",
            "blue card", "talent passport", "highly skilled", "expatriate",
            "impatriate", "inbound assignee", "intra-company transfer",
            "secondment", "global mobility",
            # Training/hiring (2026-09 review additions)
            "apprenticeship", "apprentice", "levy", "customised training",
            "pre-employment training", "recruitment screening",
            "curriculum development", "at no cost", "fully funded",
            "wage subsidy", "hiring credit", "doctorate", "researcher",
            # Spanish
            "capacitación", "formación", "empleo", "trabajador",
            "contratación", "recursos humanos", "fuerza laboral",
            "talento", "educación", "pasantía",
        ],
        "Regulatory, Ease-of-Doing-Business & Non-financial Support": [
            # English
            "single window", "single-window", "clearance", "regulatory",
            "compliance", "deemed approval", "ease of doing business",
            "sez", "fast track", "self-certification", "self certification",
            "labour law", "labor law",
            # English (2026-09 review additions)
            "investment promotion agency", "one-stop shop", "aftercare",
            "site selection assistance", "soft landing", "advance ruling",
            "tax ruling", "localisation quota", "nationalisation quota",
            "government procurement preference",
            # French
            "rescrit",
            # Spanish
            "ventanilla única", "aprobación", "licencia", "cumplimiento",
            "registro", "trámite", "facilitación", "permiso",
            "marco regulatorio", "simplificación",
        ],
        "Innovation / R&D / Patent": [
            # English
            "patent", "r&d", "research and development", "innovation lab",
            "innovation", "prototype", "prototyping", "intellectual property",
            "quality certification", "accelerator", "incubator", "centre of excellence",
            "center of excellence", "coe", "ip box",
            # English (2026-09 review additions)
            "knowledge development box", "ip regime", "nexus", "rdec",
            "research and experimentation", "collaborative research",
            "joint research", "approved research organisation",
            "contract research", "technology development zone",
            "science park", "technopole", "competitiveness cluster",
            "knowledge transfer partnership",
            # Spanish
            "innovación", "investigación y desarrollo", "i+d",
            "patente", "propiedad intelectual", "prototipo",
            "centro de excelencia", "incubadora",
        ],
        "Others": [
            # English
            "green", "sustainability", "women employee", "female employee",
            "csr", "customised package", "customized package",
            "mega project", "special package", "esg",
            # Spanish
            "sostenibilidad", "verde", "mujer trabajadora", "género",
            "responsabilidad social", "paquete especial", "régimen especial",
        ],
    }

    SCOPE_MARKERS = {
        "Specific to Tier 1 cities": [
            "tier 1", "tier-1", "tier i", "zone i ", "zone-i", "zone 1",
            "bengaluru", "bangalore", "mumbai", "pune", "chennai",
            "hyderabad", "ahmedabad", "vadodara", "indore", "bhopal",
            "lucknow", "noida", "ghaziabad", "gautambuddha nagar",
            "jaipur",
        ],
        "Specific to Tier 2 & 3 cities": [
            "tier 2", "tier-2", "tier ii", "tier 3", "tier-3", "tier iii",
            "zone ii", "zone-ii", "zone 2",
            "beyond bengaluru", "beyond bangalore", "hosur", "coimbatore",
            "mysuru", "mysore", "mangaluru", "gwalior", "jabalpur",
            "kanpur", "varanasi", "agra", "prayagraj", "udaipur",
            "kota", "jodhpur",
        ],
    }

    SKIP_PATTERNS = [
        # English boilerplate
        re.compile(r"^(to position|to establish|to attract|to promote|to foster|to enhance|to sustain|to align|to drive|to facilitate)", re.I),
        re.compile(r"^(establish|promote|foster|enhance|sustain|align|attract|facilitate|generate)\b", re.I),
        re.compile(r"\b(means|shall mean)\b", re.I),
        re.compile(r"^[A-Za-z]+\s+(means|shall mean)", re.I),
        re.compile(r"\b(committee|chairman|member secretary|sachivalaya|gandhinagar|notification|gazette|finance department|secretariat)\b", re.I),
        re.compile(r"\b(name a few|to name)\b", re.I),
        re.compile(r"\b(vision|mission|aim|objective|preamble)\b.{0,80}$", re.I),
        # Spanish boilerplate
        re.compile(r"\b(considerando|por tanto|decreta|publíquese|gaceta oficial)\b", re.I),
        re.compile(r"^(art[ií]culo|cap[ií]tulo|secci[oó]n)\s+\d", re.I),
        re.compile(r"\b(definici[oó]n|se entiende por|para los efectos)\b", re.I),
        # Personal / consumer / citizen topics (not investment incentives)
        re.compile(r"\b(personal|individual)\s+(income tax|tax return|declaration|filing)", re.I),
        re.compile(r"\b(consumer|citizen|resident)\s+(protection|service|complaint|helpline)", re.I),
        re.compile(r"\b(passport|driving licen[sc]e|driver.s licen[sc]e|birth certificate|marriage certificate)\b", re.I),
        re.compile(r"\b(pension|retirement|unemployment insurance|social welfare|old.age)\b.{0,80}$", re.I),
        re.compile(r"\b(nomad[ao]s? digital|digital nomad)\b", re.I),
        re.compile(r"\b(housing (allowance|benefit|subsidy) (for )?(employee|worker|staff|resident))\b", re.I),
        # Labor / worker rights (as opposed to investor-facing labor incentives)
        re.compile(r"\b(labor rights|labour rights|worker rights|employee rights|minimum wage|maternity leave|parental leave)\b", re.I),
        # News / gallery / careers on IPA sites
        re.compile(r"\b(press release|news update|photo gallery|read more|click here)\b", re.I),
        re.compile(r"\b(job vacancy|careers with us|about us|contact us|cookie policy|privacy policy|terms of use)\b", re.I),
        # Gulf Cooperation Council content (regional political bloc, not GCC as Global Capability Centre)
        re.compile(r"\b(gulf cooperation council|GCC country|GCC nationals?|GCC citizens?|GCC states?|GCC member)\b", re.I),
        # How-to / procedural / form boilerplate
        re.compile(r"\b(how to apply|how to register|how to enroll|to apply for|step\s?\d+\s*:)\b", re.I),
        re.compile(r"\b(EA Form|EC Form|Form No\.?|Form Number|ref no\.?|application form)\b", re.I),
        re.compile(r"\b(for further information|please contact|for more details|see appendix|kindly contact)\b", re.I),
        re.compile(r"\b(email:|phone:|telephone:|fax:|website:|address:)", re.I),
        re.compile(r"\bthis (service|form|application) (is |can |will |should )", re.I),
        # Environmental / compliance boilerplate
        re.compile(r"\b(environmental impact (report|assessment|study)|safety standards? check|machinery.*(life ?cycle|remaining life))\b", re.I),
        re.compile(r"\b(documentation such as|documents required|following documents|submit the following)\b", re.I),
        re.compile(r"\b(certification to iso|inspection report|energy consumption report|price estimation certificate)\b", re.I),
        # 2026-09 review additions: patterns observed in rows manually
        # removed from the first five country files.
        # Agency marketing copy ("We use tax credits and financial
        # incentives to help businesses invest in their workforce") rather
        # than a description of the policy itself.
        re.compile(r"^(we|our)\b", re.I),
        # Proposal/roadmap language, not an enacted instrument (one removed
        # Oregon row began with a bare list-item letter -- "d." -- and went
        # on to list recommendation items d, e and f).
        #
        # NOT skipping text that merely starts with a list marker ("A. ",
        # "a) "): a first version did, and on India's real policy PDFs it
        # deleted genuine clauses -- Gujarat's "A. Capital expenditure as per
        # GFCI..." and "A. Lease rental expenditure..." (its core CAPEX and
        # OPEX incentive definitions) and Uttar Pradesh's "a) Commercial
        # Spaces: The State Government shall facilitate...". Lettered clauses
        # are ordinary structure in policy documents.
        re.compile(r"\b(should|recommend|roadmap)\b", re.I),
        # Statute passages with no benefit content of their own (Delaware's
        # freedom of information act, a child-advocate confidentiality
        # clause, and a state-border definition all came through this way).
        re.compile(r"\bfreedom of information\b", re.I),
        re.compile(r"\bconfidential(ity)?\b.{0,60}\b(advocate|child|record)\b", re.I),
        re.compile(r"\b(state|county|city)\s+(border|boundary|boundaries)\b.{0,60}\b(defined|means|shall mean)\b", re.I),
    ]

    # Every bucketed paragraph must contain at least one of these business
    # context terms. Without this filter, generic paragraphs about consumer
    # benefits, personal tax, or unrelated policy leak into the workbook.
    # NOTE: standalone "gcc" is deliberately excluded because it matches both
    # "Gulf Cooperation Council" and "Global Capability Centre". The specific
    # phrases below cover the Capability Centre meaning.
    BUSINESS_CONTEXT = [
        # English
        "investor", "investors", "investment", "invest ", "investing",
        "company", "companies", "business", "businesses",
        "enterprise", "enterprises", "corporation", "corporate",
        "industry", "industries", "industrial", "sector", "sectors",
        "firm", "firms", "project", "projects",
        "fdi", "foreign direct investment", "foreign investor",
        "mnc", "multinational",
        # GCC unambiguated:
        "global capability centre", "global capability center",
        "capability centre", "capability center",
        # Related sector terms:
        "gbs", "bss", "bpo", "itbpm", "ites", "it/ites",
        "shared services", "innovation hub", "r&d center", "r&d centre",
        "manufacturer", "manufacturing", "startup", "start-up",
        "sme", "smes", "capital investment",
        # Spanish
        "empresa", "empresas", "inversión", "inversiones",
        "inversor", "inversores", "sector productivo", "compañía",
        "compañías", "industria", "corporación",
    ]

    # 2026-09 review: reviewing rows manually removed from the first five
    # country files, under ~150 characters was usually a menu heading and
    # over ~3,000 was usually a raw document dump rather than a policy
    # description -- the longest observed row in the United States file was
    # 32,012 characters.
    #
    # Oversized text is SPLIT into chunks of at most MAX_PARAGRAPH_LEN on
    # sentence boundaries rather than dropped (first version dropped it).
    # Dropping cost India ~15 real bucketed items: PolicyPDFReader merges
    # whole policy sections into single blocks (Karnataka has blocks up to
    # 15,822 characters, and Maharashtra's entire 2 page policy is two
    # blocks), so "over 3,000" in a genuine GCC policy PDF is content, not a
    # raw dump. Chunks still go through every other filter (skip patterns,
    # keyword classification, business context), so junk inside a long block
    # is filtered the same as any other paragraph.
    MIN_PARAGRAPH_LEN = 150
    MAX_PARAGRAPH_LEN = 3000

    # Short bullets that state a figure are kept even though they are under
    # MIN_PARAGRAPH_LEN. The 150 character floor exists to drop fragments
    # and labels, but it also dropped real one line incentive facts, e.g.
    # Fiji's "5-year tax holiday for investment between FJ$100,000 and
    # FJ$250,000 and minimum of 25 employees" (93 characters) and "Zero-rated
    # Corporate Income Tax for the first 5 years". A text of SHORT_MIN_LEN to
    # MIN_PARAGRAPH_LEN characters is therefore kept only if it contains a
    # percentage, currency amount, duration or thousands figure; it still
    # has to pass every other rule (skip patterns, incentive keyword,
    # business context). Measured on 20 countries' raw text: of 1,736 short
    # paragraphs only 10 qualified, all genuine incentive facts.
    SHORT_MIN_LEN = 50

    _FIGURE_RE = re.compile(
        r"\d+(?:[.,]\d+)?\s?%"
        r"|[$€£¥₹]\s?\d"
        r"|\b(?:USD|EUR|GBP|INR|JPY|CNY|FJD|AUD|CAD|CHF|SEK|PLN|CZK|HUF)\s?\d"
        r"|\d[\d,.]*\s?(?:million|billion|thousand)\b"
        r"|\b\d+[\s-]?(?:year|years|month|months|day|days)\b"
        r"|\b\d{1,3}(?:,\d{3})+\b",
        re.IGNORECASE,
    )

    _SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+")

    @classmethod
    def _split_oversized(cls, text):

        limit = cls.MAX_PARAGRAPH_LEN

        if len(text) <= limit:
            return [text]

        chunks = []
        current = ""

        for sentence in cls._SENTENCE_SPLIT_RE.split(text):

            # A single run-on "sentence" longer than the limit (common in
            # extracted PDF text with no punctuation): hard split on the
            # last space before the limit.
            while len(sentence) > limit:

                cut = sentence.rfind(" ", 0, limit)

                if cut <= 0:
                    cut = limit

                head, sentence = sentence[:cut], sentence[cut:].lstrip()

                if current:
                    chunks.append(current)
                    current = ""

                chunks.append(head)

            if current and len(current) + 1 + len(sentence) > limit:
                chunks.append(current)
                current = sentence
            else:
                current = f"{current} {sentence}".strip()

        if current:
            chunks.append(current)

        return chunks

    @classmethod
    def bucket(cls, paragraphs):

        buckets = {
            it: {scope: [] for scope in cls.SCOPES}
            for it in cls.INCENTIVE_TYPES
        }

        for para in paragraphs:

            for text in cls._split_oversized(para["text"]):

                if len(text) < cls.MIN_PARAGRAPH_LEN and (
                    len(text) < cls.SHORT_MIN_LEN or not cls._FIGURE_RE.search(text)
                ):
                    continue

                if cls._is_skip(text, para.get("heading", "")):
                    continue

                incentive_type = cls._classify_type(text)

                if not incentive_type:
                    continue

                if not cls._has_business_context(text):
                    continue

                scope = cls._classify_scope(text)

                buckets[incentive_type][scope].append(text)

        return buckets

    @classmethod
    def _has_business_context(cls, text):

        text_low = text.lower()

        for kw in cls.BUSINESS_CONTEXT:

            if kw.endswith(" "):
                if kw in text_low:
                    return True
            else:
                # Allow optional trailing "s" so "capability centre" matches
                # "capability centres" too.
                if re.search(rf"\b{re.escape(kw)}s?\b", text_low):
                    return True

        return False

    @classmethod
    def _is_skip(cls, text, heading):

        head_low = heading.lower()

        if any(
            kw in head_low
            for kw in ("preamble", "vision", "mission", "objectives",
                       "aims", "definitions", "definition", "interpretation",
                       "budgetary", "power to amend",
                       "policy implementation")
        ):
            return True

        for pat in cls.SKIP_PATTERNS:

            if pat.search(text[:120]):
                return True

        return False

    # CJK ideographs have no word-boundary/whitespace convention the way
    # Latin/Cyrillic scripts do -- Python's \b relies on \w transitions, and
    # two adjacent CJK characters (e.g. "の立地補助金") are both \w, so
    # \b立地\b would never match inside continuous Japanese text. Keywords
    # containing a CJK character fall back to plain substring counting.
    _CJK_RE = re.compile(r"[぀-ヿ㐀-鿿]")

    @classmethod
    def _keyword_count(cls, kw, text_low):

        if cls._CJK_RE.search(kw):
            return text_low.count(kw)

        return len(re.findall(rf"\b{re.escape(kw)}s?\b", text_low))

    @classmethod
    def _classify_type(cls, text):

        text_low = text.lower()

        best_type = None

        best_score = 0

        for incentive_type, keywords in cls.KEYWORD_MAP.items():

            # Allow optional trailing "s" so "tax holiday" matches
            # "tax holidays" too. Trailing \b keeps the tail bound.
            score = sum(cls._keyword_count(kw, text_low) for kw in keywords)

            if score > best_score:

                best_score = score
                best_type = incentive_type

        return best_type

    @classmethod
    def _classify_scope(cls, text):

        text_low = text.lower()

        for scope, markers in cls.SCOPE_MARKERS.items():

            if any(marker in text_low for marker in markers):

                return scope

        return "Across the state (Common)"
