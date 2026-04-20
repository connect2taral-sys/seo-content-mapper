import streamlit as st
import pandas as pd
import numpy as np
import re, io, json, time, urllib.request, urllib.error
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="SEO Content Mapper", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.main{padding-top:1rem;}
.stProgress>div>div>div{background-color:#1D9E75;}
.metric-card{background:#f0faf5;border:1px solid #1D9E75;border-radius:8px;padding:16px;text-align:center;}
.metric-num{font-size:2rem;font-weight:700;color:#1D9E75;}
.metric-label{font-size:0.85rem;color:#555;margin-top:4px;}
.info-box{background:#e8f5f0;border-left:4px solid #1D9E75;padding:12px 16px;border-radius:4px;margin:8px 0;}
.warn-box{background:#fffbea;border-left:4px solid #f59e0b;padding:12px 16px;border-radius:4px;margin:8px 0;}
.error-box{background:#fef2f2;border-left:4px solid #ef4444;padding:12px 16px;border-radius:4px;margin:8px 0;}
.sec-hdr{background:#1D9E75;color:white;padding:8px 16px;border-radius:4px;font-weight:600;margin:16px 0 8px 0;}
</style>""", unsafe_allow_html=True)

# ── TAXONOMY GROUPS ────────────────────────────────────────────────────────
HVAC_GRP = {'furnace','boiler','heat_pump','ac_cooling','ductless','hvac_general','geothermal'}
PLMB_GRP = {'water_pipes','drain_sewer','water_heater','toilet','faucet_sink','shower_tub',
             'bathroom','kitchen','plumbing_general','backflow','sump_pump','sprinkler'}
ELEC_GRP = {'electrical','generator','ev_charger'}
ALL_EX   = HVAC_GRP | PLMB_GRP | ELEC_GRP | {'gas_utility','air_quality'}

# Universal exclusion patterns — works for any website
EXCL = [
    # Non-service pages universal to all sites
    '/about','/about-us','/about-me','/our-story','/who-we-are',
    '/meet-the','/our-team','/team','/staff','/technicians','/employees',
    '/careers','/jobs','/join','/join-the-team','/work-for-us','/hiring',
    '/contact','/contact-us','/get-in-touch','/reach-us','/find-us',
    '/privacy','/privacy-policy','/terms','/terms-of-service','/disclaimer',
    '/sitemap','/site-map','/accessibility','/ada',
    '/cart','/checkout','/account','/login','/register','/my-account',
    '/tag/','/category/','/author/','/page/','/archive/',
    '/search','/404','/error','/thank-you','/confirmation','/success',
    '/podcast','/video','/videos','/webinar','/event','/events',
    '/press','/media','/news-room','/press-release',
    '/financing','/payment','/apply','/credit',
    '/warranty','/guarantee','/returns','/refund',
    '/affiliate','/partner','/referral','/ambassador',
    '/coupons','/specials','/deals','/offers','/promotions','/sale',
    '/community','/forum','/gallery','/portfolio',
    '/estore','/store','/shop','/product',
    '/solar','/case-stud','/case_study',
    # Truck wrap / contest pages (unusual but universal pattern)
    '/truck-wrap','/contest','/giveaway',
]

EXPAND = [(r'\bac\b','air conditioning'),(r'\ba/c\b','air conditioning'),
          (r'\bhvac\b','heating cooling air conditioning'),(r'\bfurnace\b','furnace heating'),
          (r'\bheat pump\b','heat pump heating cooling'),(r'\bboiler\b','boiler heating'),
          (r'\bgenerator\b','generator backup power'),(r'\bdrain\b','drain sewer'),
          (r'\bsump pump\b','sump pump basement'),(r'\bgfci\b','gfci outlet electrical'),
          (r'\bbackflow\b','backflow prevention'),(r'\buv\b','uv air sanitizer'),
          (r'\brepiping\b','repiping pipe replacement'),(r'\bwater heater\b','water heater hot water'),
          (r'\btankless\b','tankless water heater'),(r'\bmini.?split\b','mini split ductless heating cooling')]

# ── URL → (THEME, SUB-THEME) ordered patterns (most specific first) ────────
URL_PATTERNS = [
    ('tankless-water-heater-repair',      'Water Heater',       'Tankless Repair'),
    ('tankless-water-heater-install',     'Water Heater',       'Tankless Installation'),
    ('tankless-water-heater-maintenance', 'Water Heater',       'Tankless Maintenance'),
    ('tankless-water-heater-services',    'Water Heater',       'Tankless Service'),
    ('tankless-water-heater',             'Water Heater',       'Tankless Water Heater'),
    ('water-heater-repair',               'Water Heater',       'Water Heater Repair'),
    ('water-heater-maintenance',          'Water Heater',       'Water Heater Maintenance'),
    ('water-heater-install',              'Water Heater',       'Water Heater Installation'),
    ('heater-install',                    'Water Heater',       'Water Heater Installation'),
    ('water-heater',                      'Water Heater',       'Water Heater Service'),
    ('tankless',                          'Water Heater',       'Tankless Water Heater'),
    ('ductless-mini-split-repair',        'AC / Cooling',       'Ductless Repair'),
    ('ductless-mini-split-install',       'AC / Cooling',       'Ductless Installation'),
    ('ductless-mini-split-maintenance',   'AC / Cooling',       'Ductless Maintenance'),
    ('ductless-mini-split',               'AC / Cooling',       'Ductless Mini Split'),
    ('ductless',                          'AC / Cooling',       'Ductless Mini Split'),
    ('evaporator-coil',                   'AC / Cooling',       'Evaporator Coil'),
    ('ac-install',                        'AC / Cooling',       'AC Installation'),
    ('ac-tune',                           'AC / Cooling',       'AC Tune-up'),
    ('68-ac',                             'AC / Cooling',       'AC Tune-up'),
    ('ac-repair',                          'AC / Cooling',       'AC Repair'),
    ('ac-service',                         'AC / Cooling',       'AC Service'),
    ('air-conditioning',                  'AC / Cooling',       'Air Conditioning'),
    ('heat-pump-repair',                  'Heat Pump',          'Heat Pump Repair'),
    ('heat-pump-install',                 'Heat Pump',          'Heat Pump Installation'),
    ('heat-pump-maintenance',             'Heat Pump',          'Heat Pump Maintenance'),
    ('heat-pump-services',                'Heat Pump',          'Heat Pump Service'),
    ('heat-pump',                         'Heat Pump',          'Heat Pump Service'),
    ('mini-split',                         'Heat Pump',          'Mini Split Service'),
    ('furnace-repair',                    'Furnace / Heating',  'Furnace Repair'),
    ('furnace-install',                   'Furnace / Heating',  'Furnace Installation'),
    ('furnace-tune',                      'Furnace / Heating',  'Furnace Tune-up'),
    ('68-furnace',                        'Furnace / Heating',  'Furnace Tune-up'),
    ('furnace-service',                   'Furnace / Heating',  'Furnace Service'),
    ('furnace',                           'Furnace / Heating',  'Furnace Service'),
    ('steam-boiler',                      'Furnace / Heating',  'Steam Boiler'),
    ('boiler',                            'Furnace / Heating',  'Boiler Service'),
    ('/heating',                          'Furnace / Heating',  'Heating Service'),
    ('gas-leak',                          'Plumbing',           'Gas Leak Detection'),
    ('gas-line',                          'Plumbing',           'Gas Line Service'),
    ('water-leak',                        'Plumbing',           'Water Leak Detection'),
    ('main-water-line',                   'Plumbing',           'Water Line Service'),
    ('repiping',                          'Plumbing',           'Repiping'),
    ('frozen-pipe',                       'Plumbing',           'Frozen Pipe Repair'),
    ('burst-pipe',                        'Plumbing',           'Burst Pipe Repair'),
    ('flooded-basement',                  'Plumbing',           'Flooded Basement'),
    ('underground',                       'Plumbing',           'Underground Services'),
    ('excavation',                        'Plumbing',           'Excavation'),
    ('emergency-plumber',                 'Plumbing',           'Emergency Plumbing'),
    ('plumbing-inspection',               'Plumbing',           'Plumbing Inspection'),
    ('commercial-plumbing',               'Plumbing',           'Commercial Plumbing'),
    ('residential-plumbing',              'Plumbing',           'Residential Plumbing'),
    ('bathroom-drain-cleaning',           'Drain / Sewer',      'Bathroom Drain Cleaning'),
    ('clogged-kitchen-drain',             'Drain / Sewer',      'Kitchen Drain Cleaning'),
    ('interior-drain',                    'Drain / Sewer',      'Interior Drain Cleaning'),
    ('exterior-drain',                    'Drain / Sewer',      'Exterior Drain Cleaning'),
    ('hydrojetting',                      'Drain / Sewer',      'Hydrojetting'),
    ('drain-tile',                        'Drain / Sewer',      'Drain Tile'),
    ('bubbler',                           'Drain / Sewer',      'Storm Drain / Bubbler'),
    ('drain-cleaning',                    'Drain / Sewer',      'Drain Cleaning'),
    ('sewer-service',                     'Drain / Sewer',      'Sewer Repair'),
    ('sewer',                             'Drain / Sewer',      'Sewer Service'),
    ('drain',                             'Drain / Sewer',      'Drain Service'),
    ('sump-pump',                         'Sump Pump',          'Sump Pump Service'),
    ('backflow-testing',                  'Backflow',           'Backflow Testing'),
    ('backflow-repair',                   'Backflow',           'Backflow Repair'),
    ('backflow',                          'Backflow',           'Backflow Service'),
    ('water-treatment',                   'Water Treatment',    'Water Treatment'),
    ('garbage-disposal-install',          'Bathroom / Kitchen', 'Garbage Disposal Installation'),
    ('garbage-disposal-repair',           'Bathroom / Kitchen', 'Garbage Disposal Repair'),
    ('garbage-disposal',                  'Bathroom / Kitchen', 'Garbage Disposal'),
    ('clogged-toilet',                    'Bathroom / Kitchen', 'Toilet Repair'),
    ('toilet-install',                    'Bathroom / Kitchen', 'Toilet Installation'),
    ('toilet-repair',                     'Bathroom / Kitchen', 'Toilet Repair'),
    ('toilet',                            'Bathroom / Kitchen', 'Toilet Service'),
    ('bathroom-remodel',                  'Bathroom / Kitchen', 'Bathroom Remodel'),
    ('kitchen-remodel',                   'Bathroom / Kitchen', 'Kitchen Remodel'),
    ('circuit-breaker',                   'Electrical',         'Circuit Breaker'),
    ('electric-panels-rewiring',          'Electrical',         'Panel & Rewiring'),
    ('electrical-inspect',                'Electrical',         'Electrical Inspection'),
    ('electric-outdoor-lighting',         'Electrical',         'Outdoor Lighting'),
    ('electric-repair',                   'Electrical',         'Electrical Repair'),
    ('electric-install',                  'Electrical',         'Electrical Installation'),
    ('electrical-services',               'Electrical',         'Electrical Services'),
    ('car-home-charging',                 'EV Charger',         'EV Charger Installation'),
    ('generator-installs',                'Generator',          'Generator Installation'),
    ('standby-generator',                 'Generator',          'Standby Generator'),
    ('generators',                        'Generator',          'Generator Service'),
    ('geothermal',                        'HVAC General',       'Geothermal'),
    ('uv-air-sanitizer',                  'Air Quality',        'UV Air Sanitizer'),
    ('air-quality',                       'Air Quality',        'Air Quality'),
    ('duct-cleaning',                     'HVAC General',       'Duct Cleaning'),
    ('air-duct',                          'HVAC General',       'Duct Cleaning'),
    ('air-handler',                       'HVAC General',       'Air Handler'),
    ('hvac-service',                       'HVAC General',       'HVAC Service'),
    ('commercial-hvac',                   'Commercial',         'Commercial HVAC'),
    ('commercial-service',                 'Commercial',         'Commercial Service'),
    ('maintenance-plans',                 'HVAC General',       'Maintenance Plans'),
]

BLOG_PATTERNS = [
    ('furnace-blinking',     'Furnace / Heating', 'Furnace Troubleshooting'),
    ('furnace-fan',          'Furnace / Heating', 'Furnace Tips'),
    ('furnace-tune',         'Furnace / Heating', 'Furnace Tune-up Guide'),
    ('furnace-replace',      'Furnace / Heating', 'Furnace Replacement Guide'),
    ('furnace',              'Furnace / Heating', 'Furnace Tips & Guides'),
    ('boiler',               'Furnace / Heating', 'Boiler Tips & Guides'),
    ('heat-pump',            'Heat Pump',         'Heat Pump Tips & Guides'),
    ('tankless-water-heater','Water Heater',       'Tankless Water Heater Guide'),
    ('water-heater-not',     'Water Heater',       'Water Heater Troubleshooting'),
    ('water-heater-repair',  'Water Heater',       'Water Heater Repair Guide'),
    ('tank-vs-tankless',     'Water Heater',       'Water Heater Comparison'),
    ('cleaning-your-water',  'Water Heater',       'Water Heater Maintenance Guide'),
    ('water-heater',         'Water Heater',       'Water Heater Tips & Guides'),
    ('smelly-drain',         'Drain / Sewer',      'Drain Odor Guide'),
    ('unclog-kitchen',       'Drain / Sewer',      'Kitchen Drain Guide'),
    ('unclog-sewer',         'Drain / Sewer',      'Sewer Line Guide'),
    ('drain-cleaning',       'Drain / Sewer',      'Drain Cleaning Guide'),
    ('drain',                'Drain / Sewer',      'Drain Tips & Guides'),
    ('sewer',                'Drain / Sewer',      'Sewer Tips & Guides'),
    ('low-water-level',      'Bathroom / Kitchen', 'Toilet Water Level Guide'),
    ('toilet',               'Bathroom / Kitchen', 'Toilet Tips & Guides'),
    ('bathroom-smell',       'Bathroom / Kitchen', 'Bathroom Odor Guide'),
    ('garbage-disposal',     'Bathroom / Kitchen', 'Garbage Disposal Guide'),
    ('sump-pump',            'Sump Pump',          'Sump Pump Tips & Guides'),
    ('flickering-lights',    'Electrical',         'Electrical Troubleshooting'),
    ('back-flow',            'Backflow',           'Backflow Guide'),
    ('backflow',             'Backflow',           'Backflow Guide'),
    ('generator',            'Generator',          'Generator Tips & Guides'),
    ('frozen-pipe',          'Plumbing',           'Frozen Pipe Guide'),
    ('grease-trap',          'Drain / Sewer',      'Grease Trap Guide'),
    ('hvac-filter',          'HVAC General',       'HVAC Maintenance Guide'),
    ('hvac',                 'HVAC General',       'HVAC Tips & Guides'),
    ('indoor-air',           'Air Quality',        'Indoor Air Quality Guide'),
    ('air-quality',          'Air Quality',        'Air Quality Guide'),
    ('discolored-water',     'Water Heater',       'Water Heater Troubleshooting'),
    ('backup-power',         'Generator',          'Generator Guide'),
    ('plumbing',             'Plumbing',           'Plumbing Tips & Guides'),
]

def url_to_theme_subtheme(url):
    p = re.sub(r'https?://[^/]+', '', url.lower())
    if '/blog/' in p:
        slug = p.split('/blog/')[-1].strip('/')
        for pat, theme, sub in BLOG_PATTERNS:
            if pat in slug:
                return theme, sub
        words = slug.replace('-', ' ').replace('_', ' ')
        return 'General', words[:40].title()
    for pat, theme, sub in URL_PATTERNS:
        if pat in p:
            return theme, sub
    return '', ''

# ── INTENT CLASSIFICATION ─────────────────────────────────────────────────
# Strong informational signals — override Semrush
INFO_STRONG = [
    r'^what\b', r'^how\b', r'^why\b', r'^when\b', r'^does\b', r'^do\b',
    r'^is\b', r'^are\b', r'^can\b', r'^should\b', r'^which\b', r'^who\b',
    r'\bvs\b', r'\bversus\b', r'\bdifference between\b', r'\btips\b',
    r'\bbenefits of\b', r'\bsigns of\b', r'\bcauses of\b', r'\btypes of\b',
    r'\bhow to\b', r'\bwhat is\b', r'\bwhy is\b', r'\bhow does\b',
    r'\bwhat does\b', r'\badvantages\b', r'\bproblems with\b',
    r'\bcost of\b', r'\bcost to\b', r'\bhow much\b', r'\bprice of\b',
    r'\bpricing\b', r'\baverage cost\b', r'\bworth it\b',
    # Problem descriptions — person diagnosing, not hiring
    r'\bnot working\b', r"\bwon't\b", r"\bdoesn't\b", r'\bblowing\b',
    r'\bsmell\b', r'\bnoise\b', r'\bnoisy\b', r'\bleak\b', r'\bleaking\b',
    r'\bfailed\b', r'\bbroken\b', r'\bkeeps\b', r'\btripping\b',
    r'\bflickering\b', r'\bdripping\b', r'\bgurgling\b', r'\bhumming\b',
    r'\bclicking\b', r'\bbuzzing\b', r'\brunning constantly\b',
    r'\bnot heating\b', r'\bnot cooling\b', r'\bnot draining\b',
    r'\bwon\'t turn on\b', r'\bwon\'t start\b', r'\bno hot water\b',
    r'\bnot hot\b', r'\bwarm air\b', r'\bcold air\b', r'\blow water\b',
    r'\bhigh water\b', r'\bslow drain\b', r'\bclogged\b',
    # DIY signals
    r'\bmyself\b', r'\bmy own\b', r'\bat home\b', r'\bdiy\b',
    r'\bfix myself\b', r'\bdo it yourself\b',
    # Research/education
    r'\blifespan\b', r'\bhow long\b', r'\bwhat size\b', r'\bwhat type\b',
    r'\bhow often\b', r'\bhow many\b', r'\bguide\b', r'\bexplain\b',
    r'\bunderstand\b', r'\blearn\b', r'\bmeaning\b', r'\bdefinition\b',
    # How/what/why patterns not caught by ^ anchors
    r'\bhow it works\b', r'\bhow they work\b', r'\bwhat it does\b',
    r'\bhow it works\b', r'\bwhat are the\b', r'\bwhat is a\b',
    r'\bwhat is the\b', r'\bwhere does\b', r'\bwhere do\b', r'\bwhere is\b', r'\bwhere are\b',
    r'\bhow do you\b', r'\bhow do i\b', r'\bhow can i\b',
    r'\bexplained\b', r'\bexplanation\b', r'\boverview\b',
    r'\b\w+ uses\b', r'\b\w+ function\b', r'\bfunctions of\b',
    r'\bpurpose of\b', r'\bhow does a\b', r'\bhow does it\b',
    r'\bwhat does a\b', r'\bwhat does it\b',
    r'\bcome from\b', r'\bwork\?', r'\bworks\?',
    r'\btypes of\b', r'\bkinds of\b',
]

# Strong transactional signals — override Semrush
TRANS_STRONG = [
    r'\bnear me\b', r'\bin my area\b', r'\bclose to me\b',
    r'\bhire\b', r'\bcall\b', r'\bschedule\b', r'\bbook\b',
    r'\bget a quote\b', r'\bfree estimate\b', r'\bemergency\b',
    r'\bsame day\b', r'\b24 hour\b', r'\b24/7\b', r'\btoday\b',
    r'\bcompany near\b', r'\bcontractor near\b', r'\bservice near\b',
    r'\bplumber near\b', r'\belectrician near\b', r'\btechnician near\b',
    r'\bprofessional\b', r'\bcertified\b', r'\blicensed\b',
    r'\binstall.*near\b', r'\brepair.*near\b', r'\breplace.*near\b',
    r'\bservice.*near\b',
]

def classify_intent_rules(kw):
    """Returns 'Informational', 'Transactional', or 'Ambiguous'."""
    kl = kw.lower().strip()
    # Check strong informational
    for p in INFO_STRONG:
        if re.search(p, kl):
            return 'Informational'
    # Check strong transactional
    for p in TRANS_STRONG:
        if re.search(p, kl):
            return 'Transactional'
    # Location + service keyword = Transactional hiring intent
    # e.g. "sump pump rochester", "plumber in chicago", "drain cleaning buffalo ny"
    _SVC = {'plumber','plumbing','electrician','hvac','furnace','boiler','heater',
            'drain','sewer','pump','repair','install','service','contractor',
            'technician','company','fix','clean','replace','maintenance',
            'cooling','heating','electrical','generator','backflow','ac'}
    if any(w in kl for w in _SVC):
        words = kl.split()
        _US_ST = {'al','ak','az','ar','ca','co','ct','de','fl','ga','hi','id',
                  'il','in','ia','ks','ky','la','me','md','ma','mi','mn','ms',
                  'mo','mt','ne','nv','nh','nj','nm','ny','nc','nd','oh','ok',
                  'or','pa','ri','sc','sd','tn','tx','ut','vt','va','wa','wv',
                  'wi','wy','dc'}
        # Ends with state abbreviation → Transactional
        if len(words) >= 2 and words[-1] in _US_ST:
            return 'Transactional'
        # Contains explicit "in [city]" or "near [city]" where city ≠ generic word
        _GENERIC = {'me','area','home','town','city','state','local','zone',
                    'region','district','county','neighborhood','vicinity'}
        m = re.search(r'\b(?:in|near)\s+([a-z]{4,})$', kl)
        if m and m.group(1) not in _GENERIC:
            return 'Transactional'
        # Service + bare city name at end (no preposition, no state abbrev)
        # e.g. "sump pump rochester", "hvac repair greenlawn"
        # City = last word that is alphabetic, 4+ chars, not a common English word
        _NOT_CITY = {'repair','service','services','install','clean','replace',
                     'maintain','system','works','working','function','uses',
                     'explained','definition','meaning','overview','guide','tips',
                     'cost','price','water','pump','tank','pipe','line','unit',
                     'part','problems','issues','call','help','need','want',
                     'best','good','local','near','area','home','house','company',
                     'contractor','professional','licensed','certified','expert',
                     'emergency','residential','commercial','quality','test',
                     'inspection','cleaning','replacement','installation','pump',
                     'heater','furnace','boiler','drain','sewer','hvac','power',
                     'electric','electrical','generator','cooling','heating'}
        filtered = [w for w in words if w not in
                    {'a','an','the','in','on','at','for','to','of','and','or',
                     'my','your','our','near','best','good','local','great','free'}]
        if (len(filtered) >= 2 and
                filtered[-1] not in _NOT_CITY and
                filtered[-1] not in _GENERIC and
                len(filtered[-1]) >= 4 and
                filtered[-1].isalpha()):
            return 'Transactional'
    return 'Ambiguous'

# ── HELPERS ───────────────────────────────────────────────────────────────
def expand(t, extra_abbrevs=None):
    """Expand abbreviations. extra_abbrevs from Phase 0 Claude detection."""
    t = t.lower()
    for p, r in EXPAND: t = re.sub(p, r, t)
    if extra_abbrevs:
        for abbr, full in extra_abbrevs.items():
            t = re.sub(r'' + re.escape(abbr.lower()) + r'', full.lower(), t)
    return t

def extract_stop_words(sf_df):
    """
    Auto-detect brand name, city names and generic site-wide terms
    from Title + Meta + H1 frequency across all pages.
    Words appearing in 20%+ of pages = site-specific noise → stop words.
    Works for any domain, any city, any brand.
    """
    # Common English words to always keep (not stop words even if frequent)
    KEEP = {'the','and','in','of','for','to','a','an','is','are','we','our',
            'your','with','at','by','from','or','on','as','it','its','be',
            'was','has','have','will','can','all','new','best','top','get',
            'need','help','call','today','now','free','save','local','near',
            'professional','certified','licensed','expert','quality','service',
            'services','repair','repairs','installation','maintenance','company',
            'near','me','ny','llc','inc','co','ltd','corp'}
    word_page_count = {}
    total_pages = len(sf_df)
    for _, row in sf_df.iterrows():
        title = str(row.get('Title 1', row.get('Page Title','')) or '')
        meta  = str(row.get('Meta Description 1', row.get('Meta Description','')) or '')
        h1    = str(row.get('H1-1', row.get('H1','')) or '')
        combined = f"{title} {meta} {h1}".lower()
        words = set(re.findall(r'\b[a-z]{3,}\b', combined))
        for w in words:
            word_page_count[w] = word_page_count.get(w, 0) + 1
    threshold = total_pages * 0.20
    stop = set()
    for word, count in word_page_count.items():
        if count >= threshold and word not in KEEP:
            stop.add(word)
    # Also extract domain parts from first URL
    sample_urls = sf_df.apply(lambda r: get_url(r), axis=1).dropna()
    if len(sample_urls):
        m = re.search(r'https?://(?:www\.)?([^/]+)', str(sample_urls.iloc[0]))
        if m:
            domain_parts = re.split(r'[\.\-]', m.group(1).lower())
            for p in domain_parts:
                if len(p) > 2 and p not in KEEP:
                    stop.add(p)
    return stop

def is_excl(url, extra_excl=None):
    """Check exclusion against universal list + Title-based detection."""
    ul = str(url).lower()
    all_excl = EXCL + (extra_excl or [])
    return any(p in ul for p in all_excl)

def get_url(row):
    return str(row.get('Address', row.get('Landing Page', '')))

def slug_terms(url, stop=None):
    s = {'com','https','http','in','and','the','a','of','for','to','near','me','www'} | (stop or set())
    p = re.sub(r'https?://[^/]+', '', str(url))
    return ' '.join(x for x in re.split(r'[/\-_]', p) if x and x.lower() not in s)

def build_content(row, weights, stop=None):
    url   = get_url(row)
    slug  = expand(slug_terms(url, stop))
    title = expand(str(row.get('Title 1', row.get('Page Title', '')) or ''))
    meta  = expand(str(row.get('Meta Description 1', row.get('Meta Description', '')) or ''))
    h1r   = str(row.get('H1-1', row.get('H1', '')) or '')
    h1    = expand(h1r) if h1r.lower() not in ('nan', 'none', '') else ''
    sw, tw, hw, mw = weights
    return f"{slug} " * sw + f"{title} " * tw + f"{h1} " * hw + meta * mw

def classify_theme(kw):
    """
    Assign Theme from the keyword itself — not from any URL.
    Order matters: most specific checks first to prevent wrong matches.
    """
    kl = kw.lower().strip()
    ke = expand(kl)  # expanded version for abbreviations

    # ── Water Heater (check before general plumbing/heating) ─────────────
    if any(t in kl for t in ['water heater','hot water heater','tankless water heater',
                               'tankless heater','water heater','hot water tank',
                               'water heater repair','water heater install',
                               'water heater cost','water heater replace',
                               'water heater maintenance','water heater service',
                               'no hot water','hot water not','water not heating',
                               'leaking water heater','water heater leak']):
        return 'Water Heater'
    if 'tankless' in kl and any(t in kl for t in ['water','heat','hot']):
        return 'Water Heater'

    # ── Gas (check before plumbing) ───────────────────────────────────────
    if any(t in kl for t in ['gas leak','gas line','gas pipe','gas service',
                               'gas detector','natural gas leak','propane line',
                               'carbon monoxide','co detector','gas shut off']):
        return 'Gas Line'

    # ── Drain / Sewer (check before plumbing) ────────────────────────────
    # ── Drain / Sewer ─────────────────────────────────────────────────────
    if any(t in kl for t in ['drain clean','drain clog','drain repair','drain service',
                               'drain snake','drain block','clog drain','unclog',
                               'hydro jet','hydrojet','rooter service','sewage',
                               'sewer clean','sewer repair','sewer line','sewer service',
                               'drain smell','drain odor','sewer smell','sewer clog',
                               'clogged drain','slow drain','blocked drain','drain back',
                               'main line','storm drain','floor drain','exterior drain',
                               'interior drain','sewer camera','drain camera',
                               'sink clog','sink drain','sink keeps','toilet clog',
                               'tub drain','shower clog','basement drain',
                               'keeps clogging','keeps draining',
                               'drainage system','drain system','drain pipe',
                               'sewer pipe','sewer main']):
        return 'Drain / Sewer'
    if 'drain' in kl and not any(t in kl for t in ['drain field','drainage system','brain drain']):
        return 'Drain / Sewer'
    if 'sewer' in kl:
        return 'Drain / Sewer'

    # ── Water Treatment ───────────────────────────────────────────────────
    if any(t in kl for t in ['water filter','water filtration','water softener',
                               'water purif','water treatment','soft water','hard water',
                               'water quality','reverse osmosis','iron filter']):
        return 'Water Treatment'

    # ── Sump Pump (check before general plumbing) ────────────────────────
    if 'sump pump' in kl or 'sump-pump' in kl:
        return 'Sump Pump'

    # ── Backflow ─────────────────────────────────────────────────────────
    if 'backflow' in kl:
        return 'Backflow'

    # ── Sprinkler ────────────────────────────────────────────────────────
    if 'sprinkler' in kl:
        return 'Sprinkler'

    # ── Furnace / Heating ─────────────────────────────────────────────────
    if any(t in kl for t in ['furnace','boiler','steam boiler','steam heat',
                               'radiant heat','radiant heating','forced air',
                               'gas furnace','electric furnace','oil furnace',
                               'furnace repair','furnace install','furnace tune',
                               'furnace maintenance','furnace service','furnace clean',
                               'furnace filter','furnace not working','furnace cost',
                               'boiler repair','boiler service','boiler install',
                               'boiler maintenance','boiler tune','boiler cost',
                               'heat exchanger','flue','chimney liner']):
        return 'Furnace / Heating'
    # "heater" alone (not water heater) → furnace/heating
    if re.search(r'\bheater\b', kl) and 'water' not in kl and 'pool' not in kl:
        return 'Furnace / Heating'
    # "heating" alone with no other service context → furnace/heating
    if re.search(r'\bheating\b', kl) and not any(t in kl for t in [
            'water','pool','floor','radiant floor','geothermal',
            'heating and cooling','heating cooling','hvac']):
        return 'Furnace / Heating'

    # ── Heat Pump ────────────────────────────────────────────────────────
    if 'heat pump' in kl:
        return 'Heat Pump'

    # ── AC / Cooling (check after heat pump) ─────────────────────────────
    if any(t in kl for t in ['air condition','air conditioner','central air',
                               'ac repair','ac install','ac service','ac tune',
                               'ac unit','ac filter','ac replacement','ac cost',
                               'a/c repair','a/c install','a/c service',
                               'evaporator coil','refrigerant','freon',
                               'ac blowing','ac not cooling','ac not working',
                               'ac blows','air conditioner repair','air conditioner install',
                               'cooling system','central cooling','ac warm','warm air ac',
                               'air conditioner blowing','air conditioner not']):
        return 'AC / Cooling'
    if re.search(r'\bac\b', kl) and any(t in kl for t in ['repair','install','service','tune','cost','replace']):
        return 'AC / Cooling'
    if re.search(r'\bcooling\b', kl) and 'heating and cooling' not in kl:
        return 'AC / Cooling'

    # ── Ductless / Mini Split ─────────────────────────────────────────────
    if any(t in kl for t in ['mini split','ductless','mini-split']):
        return 'Ductless Mini Split'

    # ── HVAC General (heating AND cooling together, or generic hvac) ──────
    if any(t in kl for t in ['hvac','heating and cooling','heating & cooling',
                               'heat and cool','heating cooling',
                               'ventilation contractor','ventilation service',
                               'ventilation company','ventilation install',
                               'ventilation repair','ventilation near me']):
        return 'HVAC General'

    # ── Electrical ────────────────────────────────────────────────────────
    # ── Electrical (before ventilation to catch chandelier/lighting) ────
    if any(t in kl for t in ['electrician','electrical panel','electric panel',
                               'breaker','circuit breaker','breaker box',
                               'wiring','rewiring','electrical wiring',
                               'outlet','gfci','light fixture','lighting install',
                               'lighting repair','outdoor lighting','indoor lighting',
                               'surge protect','electrical inspect','electrical troubl',
                               'electrical repair','electrical install','electrical service',
                               'electrical work','electrical contractor','electric repair',
                               'electrical short','electrical main','electrical permit',
                               'local electrician','licensed electrician','home electrical',
                               'residential electrical','commercial electrical',
                               'chandelier','pendant light','ceiling fan install',
                               'smoke detector','carbon monoxide detector install',
                               'short circuit','electrical short','power outage',
                               'power surge','tripped breaker','tripping breaker']):
        return 'Electrical'
    # "electric" alone (not electric vehicle, not electric water heater etc.)
    if re.search(r'\belectric\b', kl) and not any(t in kl for t in [
            'water heater','vehicle','car','furnace','boiler','heat pump']):
        return 'Electrical'

    # ── Generator ────────────────────────────────────────────────────────
    if any(t in kl for t in ['generator','standby power','backup power','whole home generator',
                               'home generator','standby generator']):
        return 'Generator'

    # ── EV Charger ───────────────────────────────────────────────────────
    if any(t in kl for t in ['ev charger','ev charging','electric vehicle charg',
                               'electric car charg','charging station','home charging station',
                               'car charging']):
        return 'EV Charger'

    # ── Air Quality ───────────────────────────────────────────────────────
    if any(t in kl for t in ['air quality','air purif','air cleaner','air scrubber',
                               'dehumidif','humidif','indoor air','uv sanitizer',
                               'uv air','air filter service','whole home dehumid']):
        return 'Air Quality'

    # ── Geothermal ───────────────────────────────────────────────────────
    if 'geothermal' in kl:
        return 'Geothermal'

    # ── Ventilation / Duct (actual ventilation, not heating/cooling) ──────
    # Check BEFORE lighting to avoid air handler stealing chandelier
    if (any(t in kl for t in ['duct cleaning','air duct','ductwork','duct work',
                               'heat recovery','energy recovery',
                               'ventilation system','mechanical ventilation',
                               'air handler','air handling']) or
            re.search(r'\bhrv\b', kl) or re.search(r'\berv\b', kl)):
        return 'Ventilation / Duct'
    # drainage system → Drain / Sewer (not Ventilation)
    if 'drainage system' in kl or 'drain system' in kl:
        return 'Drain / Sewer'

    # ── Toilet ────────────────────────────────────────────────────────────
    if 'toilet' in kl:
        return 'Bathroom / Kitchen'

    # ── Bathroom / Kitchen fixtures ────────────────────────────────────────
    if any(t in kl for t in ['faucet','garbage disposal','kitchen sink','disposal unit',
                               'shower','bathtub','bath tub','shower drain',
                               'bathroom remodel','kitchen remodel','bathroom renov',
                               'kitchen renov','bathroom plumb','kitchen plumb',
                               'vanity','bathtub install','shower install']):
        return 'Bathroom / Kitchen'

    # ── General Plumbing (after all specific checks) ──────────────────────
    if any(t in kl for t in ['plumbing','plumber','pipe install','pipe repair',
                               'pipe replac','pipe clean','water line','water main',
                               'water pipe','burst pipe','pipe leak','slab leak',
                               'repiping','frozen pipe','main water','water pressure',
                               'water service','water shut off','plumbing inspect',
                               'emergency plumb','24 hour plumb','residential plumb',
                               'commercial plumb','plumbing company','plumbing service',
                               'local plumb','licensed plumb','plumbing contractor',
                               'plumbing repair','plumbing install']):
        return 'Plumbing'

    # ── Commercial (if still unclassified) ───────────────────────────────
    if any(t in kl for t in ['commercial service','commercial repair','commercial hvac',
                               'commercial plumbing','commercial electrical']):
        return 'Commercial'

    return 'Other'

def make_fill(c): return PatternFill(start_color=c, end_color=c, fill_type='solid')
def hdr(ws, row, cols, bg='1D9E75', fg='FFFFFF'):
    for i, v in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.fill = make_fill(bg); c.font = Font(bold=True, color=fg, size=11)
        c.alignment = Alignment(horizontal='left', vertical='center')
def cw(ws, widths):
    for col, w in zip('ABCDEFGHIJKLMNOPQRST', widths):
        ws.column_dimensions[col].width = w

# ── URL TAXONOMY ─────────────────────────────────────────────────────────
def cat_url(url):
    p = re.sub(r'https?://[^/]+', '', url.lower())
    c = set()
    if any(t in p for t in ['gas-leak','gas-line','gas-pipe','gas-service','gas-detection']): c.add('gas_utility')
    if any(t in p for t in ['water-leak','water-line','water-main','main-water','burst-pipe','slab-leak','repiping','frozen-pipe','water-shut']): c.add('water_pipes')
    if any(t in p for t in ['drain','sewer','hydro','jetting','rooter','bubbler','clogged-kitchen','flooded-basement']): c.add('drain_sewer')
    if any(t in p for t in ['water-heater','tankless']): c.add('water_heater')
    if any(t in p for t in ['furnace','heater-install']): c.add('furnace')
    if p.rstrip('/') == '/heating': c.update({'furnace','boiler','heat_pump','hvac_general'})
    if any(t in p for t in ['boiler','steam-boiler']): c.add('boiler')
    if 'heat-pump' in p: c.add('heat_pump')
    if any(t in p for t in ['ac-tune','ac-filter','ac-install','evaporator','ac-repair','68-ac',
                               'air-conditioning-repair','ac-service','ac-unit','ac-fix']): c.add('ac_cooling')
    if p.rstrip('/') == '/air-conditioning': c.update({'ac_cooling','hvac_general'})
    if any(t in p for t in ['ductless','mini-split']): c.update({'ductless','ac_cooling','heat_pump'})
    if any(t in p for t in ['commercial-hvac','duct-cleaning','air-duct','air-handler']): c.add('hvac_general')
    if 'commercial-hvac' in p: c.add('commercial')
    if 'geothermal' in p: c.update({'geothermal','hvac_general'})
    if any(t in p for t in ['electric','electrical','circuit-breaker','surge-protection','energy-efficient']): c.add('electrical')
    if 'generator' in p: c.update({'generator','electrical'})
    if 'car-home-charging' in p: c.update({'ev_charger','electrical'})
    if 'sump' in p: c.add('sump_pump')
    if any(t in p for t in ['air-quality','dehumidif','uv-air','iaq','indoor-air']): c.add('air_quality')
    if 'backflow' in p: c.add('backflow')
    if 'toilet' in p: c.update({'toilet','drain_sewer'})
    if any(t in p for t in ['garbage-disposal','faucet']): c.update({'faucet_sink','plumbing_general'})
    if any(t in p for t in ['bathroom-remodel','bathroom-renov']): c.update({'bathroom','plumbing_general'})
    if any(t in p for t in ['kitchen-remodel','kitchen-renov']): c.update({'kitchen','plumbing_general'})
    if 'sprinkler' in p: c.add('sprinkler')
    if any(t in p for t in ['residential-plumbing','emergency-plumber','commercial-plumbing','plumbing-inspection']): c.add('plumbing_general')
    if p.rstrip('/') in ['/maintenance-plans']: c.update({'plumbing_general','hvac_general','electrical'})
    # Location detection: match any URL segment that looks like a city/area
    # Uses a generic pattern rather than a hardcoded list
    path_parts = [x for x in re.split(r'[/\-]', p) if len(x) >= 4]
    # A URL segment is treated as a location if it appears at the start of the path
    # after the domain and is not a known service word
    SERVICE_SLUGS = {'repair','service','install','maintenance','emergency','commercial',
                     'residential','plumbing','heating','cooling','electrical',
                     'drain','sewer','water','heater','furnace','boiler','hvac','blog',
                     'about','contact','careers','privacy','terms','sitemap'}
    # Only treat as location page if the URL has a clear location+service structure:
    # e.g. /chicago/furnace-repair/ or /services/denver/ — location AND service present
    # Location detection: URL must have structure /city/service/ (separate path segments)
    # NOT /cityname-service/ (hyphenated slug like buffalo-repair)
    url_segments = [s for s in p.strip('/').split('/') if s]  # split on /
    non_svc_segs = [s for s in url_segments if s not in SERVICE_SLUGS
                    and s.isalpha() and len(s) >= 4 and '-' not in s]
    svc_segs     = [s for s in url_segments if any(sw in s for sw in SERVICE_SLUGS)]
    if non_svc_segs and svc_segs and len(url_segments) >= 2:
        # Has separate path segments for location and service
        for seg in non_svc_segs[:1]:
            c.add('location'); c.add('loc:'+seg)
    if '/blog/' in p: c.add('blog')
    # Brand detection: any word in URL that is not a common service word
    # This covers any brand (Trane, Mitsubishi, Lennox, Carrier, Goodman etc.)
    # without needing a hardcoded list
    for seg in re.split(r'[/\-]', p):
        if (len(seg) >= 4 and seg.isalpha() and seg not in
                {'repair','service','install','heating','cooling','electric','plumbing',
                 'drain','sewer','water','furnace','boiler','hvac','blog',
                 'about','contact','home','residential','commercial','emergency'}):
            c.add('brand:'+seg)
    if not (c - {'blog'}): c.add('plumbing_general')
    return frozenset(c)

def cat_kw(kw):
    kl = expand(kw.lower())
    c = set()
    GAS = ['gas leak','gas line','gas pipe','natural gas line','gas service line','gas detector',
           'gas fireplace','gas stove','gas dryer','gas connection','gas hookup','gas meter',
           'carbon monoxide detector','gas shut off','propane line']
    if any(t in kl for t in GAS):
        if not any(x in kl for x in ['furnace heating','boiler heating','water heater hot water']): c.add('gas_utility')
    WH = ['water heater hot water','water heater','hot water heater','tankless water heater',
          'hot water tank','water heater leak','water heater temp','water heater not','no hot water']
    if any(t in kl for t in WH): c.add('water_heater')
    if 'water_heater' not in c:
        WP = ['water leak','water line','water main','water pipe','burst pipe','pipe leak',
              'slab leak','repiping','main water','frozen pipe','water service line',
              'water shut off','water pressure']
        if any(t in kl for t in WP): c.add('water_pipes')
    DR = ['drain sewer','sewer','unclog','hydro jet','rooter','sewage','septic',
          'main line clog','backed up drain','slow drain','floor drain','bubbler',
          'storm drain','exterior drain','interior drain','basement drain','laundry drain',
          'bathroom drain','kitchen drain']
    if any(t in kl for t in DR): c.add('drain_sewer')
    FU = ['furnace heating','furnace tune','furnace service','furnace repair','furnace install',
          'furnace maint','furnace clean','furnace filter','furnace not','annual furnace',
          'gas furnace','electric furnace','new furnace','furnace replace','forced air']
    if any(t in kl for t in FU): c.add('furnace')
    BO = ['boiler heating','boiler repair','boiler service','boiler install','boiler maint',
          'boiler replace','boiler tune','annual boiler','gas boiler','steam heat','steam boiler','radiant heat']
    if any(t in kl for t in BO): c.add('boiler')
    if 'heat pump heating cooling' in kl: c.add('heat_pump')
    AC = ['air conditioning','air conditioner','central air','cooling system','evaporator coil','refrigerant','freon']
    if any(t in kl for t in AC): c.add('ac_cooling')
    if 'mini split ductless' in kl: c.update({'ductless','ac_cooling','heat_pump'})
    if 'heating cooling air conditioning' in kl: c.update({'hvac_general','furnace','ac_cooling'})
    elif 'heating cooling' in kl: c.update({'hvac_general','furnace','ac_cooling'})
    if 'geothermal' in kl: c.update({'geothermal','hvac_general'})
    if any(t in kl for t in ['commercial','industrial','business facility','office building']): c.add('commercial')
    is_combo = any(x in kl for x in ['water heater','furnace heating','boiler heating','heat pump heating'])
    EL = ['electrician','electric panel','breaker box','circuit breaker','electrical panel','wiring',
          'rewiring','outlet','gfci outlet electrical','light fixture','lighting','surge protect',
          'electrical inspect','electrical service','electrical repair','electrical install',
          'electrical work','electrical contractor','electric repair','residential electric',
          'home electrical','electrical company','local electrician']
    if any(t in kl for t in EL) and not is_combo: c.add('electrical')
    if 'generator backup power' in kl: c.update({'generator','electrical'})
    if any(t in kl for t in ['ev charger','electric vehicle charg','charging station','car charging','home charging']): c.update({'ev_charger','electrical'})
    if 'sump pump basement' in kl: c.add('sump_pump')
    AQ = ['air quality','dehumidif','humidif','air purif','uv air sanitizer','indoor air','air cleaner']
    if any(t in kl for t in AQ): c.add('air_quality')
    if 'backflow prevention' in kl: c.add('backflow')
    if 'toilet' in kl: c.update({'toilet','drain_sewer'})
    if any(t in kl for t in ['faucet','garbage disposal','kitchen sink','disposal unit']): c.update({'faucet_sink','plumbing_general'})
    if any(t in kl for t in ['shower','bathtub','bath tub']): c.update({'shower_tub','drain_sewer','plumbing_general'})
    if any(t in kl for t in ['bathroom renov','bathroom remodel']): c.update({'bathroom','plumbing_general'})
    if any(t in kl for t in ['kitchen renov','kitchen remodel']): c.update({'kitchen','plumbing_general'})
    if 'sprinkler' in kl: c.add('sprinkler')
    PL = ['plumbing','plumber','pipe install','pipe replac','pipe repair','pipe clean',
          'residential plumb','commercial plumb','plumbing inspect','emergency plumb',
          '24 hour plumb','plumbing company','plumbing service','local plumb','home plumb',
          'plumb contractor','licensed plumb','best plumb','professional plumb']
    if any(t in kl for t in PL): c.add('plumbing_general')
    # Brand detection in keywords: same approach as URL brand detection
    for word in re.findall(r'\b[a-z]{4,}\b', kl):
        if (word not in {'repair','service','install','heating','cooling','electrical',
                         'plumbing','drain','sewer','water','furnace','boiler','hvac',
                         'pump','tank','panel','filter','unit','system','pipe','line',
                         'near','local','best','cheap','cost','price','free','home',
                         'residential','commercial','emergency','professional','licensed',
                         'certified','company','contractor','plumber','electrician'} and
                len(word) >= 4 and
                re.search(r'\b' + word + r'\b', kl)):
            # Only flag as brand if it appears alongside a service term
            if any(st in kl for st in ['repair','install','service','dealer','authorized',
                                        'certified','technician','heat pump','furnace']):
                c.add('brand:'+word)
    if 'heater' in kl and 'water heater' not in kl and 'water_heater' not in c: c.update({'furnace','hvac_general'})
    if re.search(r'\bheating\b', kl) and not (c & {'furnace','boiler','heat_pump','hvac_general','water_heater'}): c.update({'furnace','hvac_general'})
    if re.search(r'\bcooling\b', kl) and not (c & {'ac_cooling','ductless','heat_pump','hvac_general'}): c.update({'ac_cooling','hvac_general'})
    return frozenset(c)

def is_compat(kc, uc, kw, url):
    if 'location' in uc:
        kl = expand(kw.lower())
        locs = [t for t in uc if t.startswith('loc:')]
        return bool(locs) and locs[0].replace('loc:', '') in kl
    url_brands = {t for t in uc if t.startswith('brand:')}
    kw_brands  = {t for t in kc if t.startswith('brand:')}
    if url_brands and not (url_brands & kw_brands): return False
    if 'commercial' in uc and 'commercial-hvac' in url.lower():
        if 'commercial' not in kc: return False
    if 'gas_utility' in uc and 'gas_utility' not in kc: return False
    if 'gas_utility' in kc and 'gas_utility' not in uc: return False
    ke = kc & ALL_EX; ue = uc & ALL_EX
    if not ke or not ue: return True
    if ke & ue: return True
    ksh = kc & (HVAC_GRP-{'hvac_general'}); ush = uc & (HVAC_GRP-{'hvac_general'})
    if 'hvac_general' in kc and not ksh and ue & HVAC_GRP: return True
    if 'hvac_general' in uc and not ush and ke & HVAC_GRP: return True
    ksp = kc & (PLMB_GRP-{'plumbing_general'}); usp = uc & (PLMB_GRP-{'plumbing_general'})
    if 'plumbing_general' in kc and not ksp and ue & PLMB_GRP: return True
    if 'plumbing_general' in uc and not usp and ke & PLMB_GRP: return True
    if 'electrical' in kc and ue & ELEC_GRP: return True
    if 'electrical' in uc and ke & ELEC_GRP: return True
    return False

def is_compat_universal(kw_theme, url_theme):
    """Universal compatibility: keyword theme must match URL theme.
    Works for any industry. Falls back to True if either theme unknown."""
    if not kw_theme or not url_theme: return True
    if kw_theme == 'Other' or url_theme == 'Other': return True
    return kw_theme == url_theme

def clean_kw(kw):
    kw = expand(kw.lower().strip())
    for p in [r'\bnear me\b', r'\bnear\b']: kw = re.sub(p, '', kw)
    return re.sub(r'\s+', ' ', kw).strip()

# ── CLAUDE API WITH KEEP-ALIVE ────────────────────────────────────────────
def call_claude(api_key, prompt, max_tokens=2000, timeout=50):
    """Hard timeout on every call — never hangs."""
    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": max_tokens,
        "temperature": 0,
        "messages": [{"role": "user", "content": prompt}]
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=payload,
        headers={"Content-Type": "application/json",
                 "x-api-key": api_key,
                 "anthropic-version": "2023-06-01"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.loads(r.read())
    txt = data['content'][0]['text'].strip()
    txt = re.sub(r'^```json?\s*', '', txt)
    txt = re.sub(r'\s*```$', '', txt)
    return json.loads(txt.strip())

def run_batches(api_key, batches, make_prompt, parse_result, fallback_fn,
                status_prefix, status_text, progress_bar, p0, p1,
                max_tokens=1500, timeout=50):
    """Generic batch runner with keep-alive UI updates on every batch."""
    out = {}
    total = len(batches)
    for bi, batch in enumerate(batches):
        # ── KEEP-ALIVE: update UI before every API call ───────────────────
        pct = min(p0 + int((bi / max(total,1)) * (p1 - p0)), p1 - 1)
        progress_bar.progress(pct)
        status_text.text(f"{status_prefix}: batch {bi+1}/{total} ({len(out):,} done)...")
        prompt = make_prompt(batch)
        for attempt in range(3):
            try:
                result = call_claude(api_key, prompt, max_tokens, timeout)
                parsed = parse_result(result, batch)
                out.update(parsed)
                break
            except Exception:
                if attempt < 2:
                    time.sleep(2)
                else:
                    out.update(fallback_fn(batch))
        time.sleep(0.1)
    return out

def claude_intent(api_key, keywords, biz_desc, status_text, progress_bar, p0, p1):
    """
    Claude classifies intent for every non-obvious keyword.
    No Ambiguous — must decide Informational or Transactional.
    Context-aware: understands that some keywords without question words
    are still informational (e.g. "ac unit common problems", "furnace lifespan").
    """
    batches = [keywords[i:i+120] for i in range(0, len(keywords), 120)]
    def make_prompt(batch):
        return f"""You are an expert SEO analyst. Classify each keyword as either
"Informational" or "Transactional" based on what a real searcher wants.

Business context: {biz_desc}

Informational = searcher wants to LEARN, RESEARCH, or DIAGNOSE
  - Questions and educational queries
  - Cost/pricing research ("how much does X cost", "X pricing")
  - Troubleshooting ("X not working", "X making noise")
  - Comparisons, guides, tips, explanations
  - Even without question words: "furnace lifespan", "ac common problems"

Transactional = searcher wants to HIRE, BUY, or GET a SERVICE
  - Service requests ("X repair", "X installation service")
  - Finding a provider ("X company", "X contractor", "X specialist")
  - Location-based service ("plumber in chicago", "ac repair rochester")
  - Product purchases

IMPORTANT: Do NOT assume "near me" is required for Transactional.
"furnace repair" alone is Transactional — searcher wants to hire.
"furnace repair cost" is Informational — searcher is researching.

Return ONLY valid JSON: {{"keyword": "Informational" or "Transactional", ...}}

Keywords: {json.dumps(batch)}"""

    def parse_result(r, batch):
        out = {}
        for kw in batch:
            v = r.get(kw, '')
            out[kw] = v if v in ('Informational','Transactional') else 'Transactional'
        return out

    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=parse_result,
        fallback_fn=lambda b: {k: 'Transactional' for k in b},
        status_prefix="Claude — intent classification",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=1500, timeout=45
    )


def claude_relevance(api_key, keywords, biz_desc, excl_str, status_text, progress_bar, p0, p1):
    batches = [keywords[i:i+100] for i in range(0, len(keywords), 100)]
    excl_note = f"\nServices NOT offered: {excl_str}" if excl_str.strip() else ""
    def make_prompt(batch):
        return f"""SEO analyst. Classify each keyword relevance for this business.
Business: {biz_desc}{excl_note}
Labels: RELEVANT | NOT_RELEVANT | BORDERLINE
Return ONLY JSON: {{"keyword":"LABEL",...}}
Keywords: {json.dumps(batch)}"""
    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=lambda r, b: ({k.lower(): v for k, v in r.items()} if isinstance(r, dict) else {k.lower(): 'BORDERLINE' for k in b}),
        fallback_fn=lambda b: {k: 'BORDERLINE' for k in b},
        status_prefix="Claude — business relevance",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=1200, timeout=45
    )

def claude_cluster(api_key, keywords_with_vol, content_type, status_text, progress_bar, p0, p1):
    batches = [keywords_with_vol[i:i+100] for i in range(0, len(keywords_with_vol), 100)]
    out = []
    total = len(batches)
    for bi, batch in enumerate(batches):
        pct = min(p0 + int((bi / max(total,1)) * (p1 - p0)), p1 - 1)
        progress_bar.progress(pct)
        status_text.text(f"Claude — clustering {content_type}: batch {bi+1}/{total}...")
        kw_list = [{"keyword": k, "volume": v} for k, v in batch]
        prompt = f"""SEO content strategist. Group these keywords into topic clusters for {content_type} creation.
Rules:
1. Each cluster = ONE piece of content
2. Group by same entity AND same searcher intent
3. Aim for 3-8 keywords per cluster
4. Primary keyword = highest volume with clearest intent
5. content_type: "Blog post" or "Service page"
6. intent: "Educational" | "Problem-aware" | "Comparison" | "Transactional" | "Cost/Pricing"
7. suggested_title = SEO-optimised title for the content piece
Return ONLY valid JSON array:
[{{"cluster_name":"short name","entity":"service entity","intent":"type","content_type":"Blog post or Service page","suggested_title":"title","primary_keyword":"kw","primary_volume":0,"secondary_keywords":["kw2"],"total_volume":0}}]
Keywords: {json.dumps(kw_list)}"""
        for attempt in range(3):
            try:
                result = call_claude(api_key, prompt, 3000, 55)
                if isinstance(result, list): out.extend(result)
                break
            except Exception:
                if attempt < 2: time.sleep(3)
        time.sleep(0.2)
    return out


# ── PHASE 0: GENERATE LOCKED TAXONOMY ────────────────────────────────────
def generate_taxonomy(api_key, biz_desc, sf_df, status_text, progress_bar):
    """
    Generate a locked taxonomy (Themes + Sub-themes) specific to this business
    using the business description and page titles/H1s from Screaming Frog.

    This runs ONCE before any keyword processing.
    All subsequent Claude calls use this taxonomy — guaranteeing consistent
    theme and sub-theme names across runs, accounts and websites.
    """
    status_text.text("Phase 0: Generating taxonomy for this business...")
    progress_bar.progress(1)

    # Collect page titles and H1s from Screaming Frog (up to 60 pages)
    page_signals = []
    for _, row in sf_df.head(120).iterrows():
        title = str(row.get('Title 1', row.get('Page Title','')) or '').strip()
        h1    = str(row.get('H1-1', row.get('H1','')) or '').strip()
        # Clean brand suffix from title
        title_clean = re.sub(r'\s*[\|\-–]\s*.{0,40}$', '', title).strip()
        if title_clean and title_clean.lower() not in ('nan','none',''):
            page_signals.append(title_clean)
        elif h1 and h1.lower() not in ('nan','none',''):
            page_signals.append(h1)

    # Deduplicate and limit
    seen = set()
    unique_signals = []
    for s in page_signals:
        if s.lower() not in seen:
            seen.add(s.lower())
            unique_signals.append(s)
        if len(unique_signals) >= 60:
            break

    prompt = f"""You are an expert SEO strategist building a content taxonomy.

Business description:
{biz_desc}

Existing page titles from the website:
{json.dumps(unique_signals)}

Generate a complete SEO content taxonomy for this business.

Rules:
1. Themes = top-level service categories (6-16 themes)
2. Sub-themes = specific content topics within each theme (3-8 per theme)
3. Each sub-theme represents exactly ONE page or blog post
4. Include BOTH transactional sub-themes (service pages) AND informational
   sub-themes (blog posts) for each theme
5. Sub-theme names: max 4 words, Title Case, specific not generic
6. Base themes on the BUSINESS DESCRIPTION — not just existing pages
   (include service areas even if no page exists yet)

Good sub-theme examples:
  "Furnace Repair", "Furnace Installation", "Furnace Tune-up",
  "Furnace Cost Guide", "Furnace Troubleshooting", "Furnace Lifespan"

Bad sub-theme examples (too generic):
  "Service", "Repair", "Guide", "Info"

Return ONLY valid JSON:
{{
  "taxonomy": {{
    "Theme Name": ["Sub-theme 1", "Sub-theme 2", "Sub-theme 3"],
    ...
  }}
}}"""

    try:
        result = call_claude(api_key, prompt, 3000, 60)
        taxonomy = result.get('taxonomy', {})
        if taxonomy:
            n_themes = len(taxonomy)
            n_subs   = sum(len(v) for v in taxonomy.values())
            status_text.text(f"Phase 0: Taxonomy ready — {n_themes} themes, {n_subs} sub-themes")
        else:
            taxonomy = {}
    except Exception:
        taxonomy = {}

    # Step 2: URL → theme map + industry abbreviations
    status_text.text("Phase 0: Categorising pages and detecting abbreviations...")
    url_meta = []
    for _, row in sf_df.iterrows():
        url   = get_url(row)
        if not url or is_excl(url): continue
        title = str(row.get('Title 1', row.get('Page Title','')) or '').strip()
        h1    = str(row.get('H1-1', row.get('H1','')) or '').strip()
        title_clean = re.sub(r'\s*[\|\-–]\s*.{0,40}$', '', title).strip()
        slug  = re.sub(r'https?://[^/]+', '', url).strip('/')[:60]
        if (title_clean or h1) and slug:
            url_meta.append({"slug": slug, "title": title_clean[:100], "h1": h1[:80]})
        if len(url_meta) >= 150: break

    url_theme_map = {}
    abbrev_map    = {}
    if url_meta:
        taxonomy_themes = list(taxonomy.keys()) if taxonomy else []
        theme_list = json.dumps(taxonomy_themes) if taxonomy_themes else "derive from business"
        url_prompt = f"""SEO expert. For each page assign its theme. Also return industry abbreviations.

Business: {biz_desc}
Available themes: {theme_list}

Return ONLY valid JSON:
{{
  "url_themes": {{"slug": "Theme Name", ...}},
  "abbreviations": {{"abbrev": "full form", ...}}
}}

Rules for url_themes: use title and H1 as primary signals.
Rules for abbreviations: only include abbreviations used in this industry (max 20).
Examples for HVAC: {{"ac": "air conditioning", "hvac": "heating ventilation air conditioning"}}
Examples for dental: {{"tmj": "temporomandibular joint", "cerec": "ceramic reconstruction"}}

Pages:
{json.dumps(url_meta[:80])}"""
        try:
            res2 = call_claude(api_key, url_prompt, 3000, 60)
            url_theme_map = res2.get('url_themes', {})
            abbrev_map    = res2.get('abbreviations', {})
            status_text.text(
                f"Phase 0: Done — {len(url_theme_map)} pages categorised, "
                f"{len(abbrev_map)} abbreviations")
        except Exception:
            pass

    progress_bar.progress(2)
    if not taxonomy and not url_theme_map:
        status_text.text("Phase 0: Using adaptive mode")
    return taxonomy, url_theme_map, abbrev_map

def taxonomy_to_prompt_block(taxonomy):
    """
    Format the locked taxonomy as a constraint block for Claude prompts.
    Injected into every theme/sub-theme assignment prompt.
    """
    if not taxonomy:
        return ""
    lines = ["\nYOU MUST use ONLY these themes and sub-themes (no new ones):"]
    for theme, subs in taxonomy.items():
        lines.append(f"  {theme}: {', '.join(subs)}")
    lines.append("If a keyword doesn't fit exactly, use the closest sub-theme.\n")
    return "\n".join(lines)

# ── PHASE 1: GSC VALIDATION ───────────────────────────────────────────────
def phase1_gsc(gsc_df, sf_df, weights, status_text, progress_bar):
    status_text.text("Phase 1: Extracting site vocabulary from Screaming Frog...")
    progress_bar.progress(3)

    # Auto-detect stop words from Title + Meta + H1 frequency
    stop = extract_stop_words(sf_df)

    # Build extra exclusions from page Title content (universal)
    TITLE_EXCL_SIGNALS = ['about us','meet the','our team','careers','contact us',
                          'privacy policy','terms of service','404','thank you',
                          'join our team','work for us','sign in','log in']
    extra_excl = []
    for _, row in sf_df.iterrows():
        title = str(row.get('Title 1', row.get('Page Title','')) or '').lower()
        if any(sig in title for sig in TITLE_EXCL_SIGNALS):
            url = get_url(row)
            if url:
                slug = re.sub(r'https?://[^/]+', '', url).strip('/')
                if slug and slug not in extra_excl:
                    extra_excl.append('/' + slug[:40])

    # Auto-detect homepage slugs from stop words (city names, brand names in URLs)
    # A page with slug that is entirely stop words = homepage or generic landing page
    hp_slugs = stop

    url_df = sf_df[~sf_df.apply(lambda r: is_excl(get_url(r), extra_excl), axis=1)].reset_index(drop=True)
    url_df['content'] = url_df.apply(lambda r: build_content(r, weights, stop), axis=1)
    url_df['cats']    = url_df.apply(lambda r: cat_url(get_url(r)), axis=1)
    url_df['ptype']   = url_df.apply(lambda r: 'blog' if '/blog/' in get_url(r).lower() else 'service', axis=1)
    # Store metadata for Claude Theme/Sub-theme lookup later
    url_df['title_raw'] = url_df.apply(lambda r: str(r.get('Title 1', r.get('Page Title','')) or ''), axis=1)
    url_df['meta_raw']  = url_df.apply(lambda r: str(r.get('Meta Description 1', r.get('Meta Description','')) or ''), axis=1)
    url_df['h1_raw']    = url_df.apply(lambda r: str(r.get('H1-1', r.get('H1','')) or ''), axis=1)

    content_map = {get_url(r).lower().strip(): url_df.at[i, 'content'] for i, r in url_df.iterrows()}
    gsc_df['kl'] = gsc_df['Query'].str.lower().str.strip()
    dedup = {}
    for kl, grp in gsc_df.groupby('kl'):
        best = grp.sort_values(['Clicks', 'Impressions'], ascending=False).iloc[0]
        dedup[kl] = {'url': best['Landing Page'], 'clicks': best['Clicks'],
                     'imps': best['Impressions'], 'pos': best.get('Position', np.nan), 'ctr': best.get('CTR', 0)}
    progress_bar.progress(8)
    status_text.text("Phase 1: Scoring GSC queries against page content...")
    queries = list(dedup.keys())
    cleaned_q = [clean_kw(q) for q in queries]
    all_c = list(content_map.values())
    url_keys = list(content_map.keys())
    vec = TfidfVectorizer(ngram_range=(1,3), min_df=1, sublinear_tf=True)
    vec.fit(all_c + cleaned_q)
    sims = cosine_similarity(vec.transform(cleaned_q), vec.transform(all_c))
    uk = {k: i for i, k in enumerate(url_keys)}
    validated = []
    for i, (q, cq) in enumerate(zip(queries, cleaned_q)):
        if i % 2000 == 0:
            progress_bar.progress(min(8 + int((i / len(queries)) * 14), 21))
            status_text.text(f"Phase 1: Validating {i:,}/{len(queries):,} GSC queries...")
        d = dedup[q]
        gu = str(d['url']).lower().strip()
        idx = uk.get(gu)
        score = round(float(sims[i, idx]), 4) if idx is not None else 0.0
        status = 'Confirmed' if score >= 0.30 else ('Plausible' if score >= 0.15 else 'Suspicious')
        try: pos_val = round(float(d['pos']), 1)
        except: pos_val = None
        validated.append({'Query': q, 'Mapped URL': d['url'], 'Clicks': d['clicks'],
                          'Impressions': d['imps'], 'Position': pos_val, 'CTR': d['ctr'],
                          'Content Score': score, 'Mapping Status': status})
    progress_bar.progress(22)
    return pd.DataFrame(validated), url_df, dedup, stop, hp_slugs

# ── PHASE 2: KEYWORD MAPPING ──────────────────────────────────────────────
def claude_validate_match(api_key, kw_candidates, biz_desc,
                          status_text, progress_bar, p0, p1):
    """
    Claude validates whether a candidate URL genuinely covers a keyword.
    Input: list of (keyword, [candidates]) where each candidate has url/title/h1/meta.
    Output: {keyword: {"url": best_url_or_null, "confidence": "High"/"Medium"/null}}

    High:   page directly covers this keyword topic, searcher would be satisfied
    Medium: page covers the parent topic, keyword is a natural subtopic
    null:   no genuine match — even if TF-IDF score was high
    """
    batches = [kw_candidates[i:i+50] for i in range(0, len(kw_candidates), 50)]

    def make_prompt(batch):
        items = []
        for kw, candidates in batch:
            items.append({
                "keyword": kw,
                "candidates": [
                    {"url": c["url"], "title": c["title"],
                     "h1": c["h1"], "meta": c["meta"][:120]}
                    for c in candidates
                ]
            })
        return f"""You are a senior SEO strategist. For each keyword, decide if any
candidate page genuinely covers that keyword's topic.

Business: {biz_desc}

For each keyword evaluate BOTH:
1. Does the page ALREADY cover this keyword's topic?
2. Would a searcher typing this keyword be satisfied landing on this page?

Use ONLY the title, H1 and meta description to judge — not the URL slug.

Confidence levels:
  "High":   Page directly and specifically covers this keyword.
             e.g. keyword "furnace repair near me" + title "Furnace Repair Services" → High
  "Medium": Page covers the parent topic and keyword is a natural subtopic.
             e.g. keyword "furnace tune-up cost" + title "Furnace Maintenance Services" → Medium
  null:     No genuine match. TF-IDF may have scored it but topic is different.
             e.g. keyword "ac troubleshooting" + title "Energy Efficient Lighting" → null

Return ONLY valid JSON:
{{
  "keyword": {{
    "url": "matched_url_or_null",
    "confidence": "High" or "Medium" or null
  }},
  ...
}}

Items to evaluate:
{json.dumps(items)}"""

    def parse_result(r, batch):
        out = {}
        for kw, _ in batch:
            val = r.get(kw, {})
            if isinstance(val, dict) and val.get('url') and val.get('confidence') in ('High','Medium'):
                out[kw] = {"url": val["url"], "confidence": val["confidence"]}
            else:
                out[kw] = {"url": None, "confidence": None}
        return out

    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=parse_result,
        fallback_fn=lambda b: {kw: {"url": None, "confidence": None} for kw, _ in b},
        status_prefix="Claude — URL match validation",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=2000, timeout=55
    )

def phase2_map(sem_df, url_df, gsc_dedup, gsc_val_df, weights, threshold,
               your_col, intent_map, hp_slugs, url_theme_map,
               status_text, progress_bar):
    """
    Phase 2: Two-step keyword-to-URL mapping.

    Step A — TF-IDF shortlist (fast, cheap):
      Finds top 3 candidate URLs per keyword above 0.10 threshold.
      Lower threshold than before — Claude decides quality, not score.
      GSC URL added as Candidate A (strongest prior signal).

    Step B — Claude validates (accurate, reliable):
      Claude reads keyword + each candidate Title + H1 + Meta.
      Returns best match URL + confidence (High/Medium) or null.
      Only High/Medium matches are kept — no weak matches ever shown.
    """
    status_text.text("Phase 2: Building TF-IDF shortlists...")
    progress_bar.progress(24)

    svc_df  = url_df[url_df['ptype'] == 'service'].reset_index(drop=True)
    blog_df = url_df[url_df['ptype'] == 'blog'].reset_index(drop=True)

    # Homepage detection
    hp_urls = set()
    for _, row in url_df.iterrows():
        u = get_url(row)
        slug = re.sub(r'https?://[^/]+', '', u).strip('/')
        slug_words = set(re.split(r'[/\-_]', slug.lower())) - {''}
        if not slug_words or slug_words.issubset(hp_slugs | {''}):
            hp_urls.add(u.lower().strip())

    def get_url_theme(url):
        if not url or not url_theme_map: return ''
        slug = re.sub(r'https?://[^/]+', '', url.lower()).strip('/')
        if slug in url_theme_map: return url_theme_map[slug]
        for k, v in url_theme_map.items():
            if slug.endswith(k) or k.endswith(slug): return v
        return ''

    # Build URL metadata lookup for Claude validation
    url_meta_lookup = {}
    for _, row in url_df.iterrows():
        u = get_url(row)
        if not u: continue
        title = str(row.get('title_raw', row.get('Title 1','')) or '').strip()
        h1    = str(row.get('h1_raw',    row.get('H1-1',''))    or '').strip()
        meta  = str(row.get('meta_raw',  row.get('Meta Description 1','')) or '').strip()
        title_clean = re.sub(r'\s*[\|\-–]\s*.{0,40}$', '', title).strip()
        url_meta_lookup[u.lower().strip()] = {
            'url':   u,
            'title': title_clean[:120],
            'h1':    h1[:80],
            'meta':  meta[:160],
        }

    gsc_status = dict(zip(gsc_val_df['Query'].str.lower().str.strip(),
                          gsc_val_df['Mapping Status']))
    keywords = sem_df['Keyword'].fillna('').tolist()
    volumes  = sem_df['Volume'].fillna(0).tolist()
    your_pos = sem_df[your_col].fillna('').tolist() if your_col in sem_df.columns else [''] * len(keywords)
    cleaned  = [clean_kw(k) for k in keywords]

    # TF-IDF models
    def build_sims(df):
        v = TfidfVectorizer(ngram_range=(1,3), min_df=1, sublinear_tf=True)
        v.fit(df['content'].tolist() + cleaned)
        return cosine_similarity(v.transform(cleaned), v.transform(df['content']))

    ss = build_sims(svc_df)
    bs = build_sims(blog_df)
    progress_bar.progress(38)

    # ── Step A: TF-IDF shortlist ─────────────────────────────────────────
    SHORTLIST_THRESHOLD = 0.10   # low — Claude decides quality not score
    MAX_CANDIDATES      = 3      # top N per keyword

    status_text.text("Phase 2: Building candidate shortlists...")
    kw_candidates = []   # list of (keyword, [candidates])
    kw_meta       = []   # parallel list of basic metadata

    for i, kw in enumerate(keywords):
        if i % 500 == 0:
            progress_bar.progress(min(38 + int((i / len(keywords)) * 20), 57))
            status_text.text(f"Phase 2: Shortlisting {i:,}/{len(keywords):,}...")

        kl     = kw.lower().strip()
        intent = intent_map.get(kl, intent_map.get(kw.lower(), 'Transactional'))

        candidates = []

        # Candidate A: GSC URL (strongest prior)
        gd  = gsc_dedup.get(kl, {})
        gu  = gd.get('url', '')
        gst = gsc_status.get(kl, '')
        if gu and not is_excl(gu) and gst in ('Confirmed','Plausible'):
            meta = url_meta_lookup.get(gu.lower().strip())
            if meta:
                candidates.append(meta)

        # Candidates from TF-IDF — use both service and blog pools
        pools = [(bs, blog_df), (ss, svc_df)]
        for sim_matrix, df in pools:
            for idx in np.argsort(sim_matrix[i])[::-1]:
                if len(candidates) >= MAX_CANDIDATES + 1: break
                s = float(sim_matrix[i][idx])
                if s < SHORTLIST_THRESHOLD: break
                url = get_url(df.iloc[idx])
                if url.lower().strip() in hp_urls: continue
                meta = url_meta_lookup.get(url.lower().strip())
                if meta and not any(c['url'] == url for c in candidates):
                    candidates.append(meta)

        # Determine ranking status from Semrush position
        rp = str(your_pos[i]).strip()
        try: pn = float(rp)
        except: pn = None
        if pn and pn > 0:
            if pn <= 10:    rs = 'Ranking p1-10'
            elif pn <= 20:  rs = 'Quick win p11-20'
            elif pn <= 50:  rs = 'Weak ranking p21-50'
            else:           rs = 'Very weak p51-100'
        else: rs = 'Not ranking'

        kw_candidates.append((kw, candidates))
        kw_meta.append({
            'Keyword': kw, 'Volume': int(volumes[i]) if volumes[i] else 0,
            'Your Position': rp or 'N/A', 'Intent': intent,
            'Ranking Status': rs, '_gsc_url': gu, '_gsc_status': gst
        })

    progress_bar.progress(58)

    # ── Step B: Claude validates candidates ─────────────────────────────
    # Only send keywords that have at least one candidate
    to_validate = [(kw, cands) for kw, cands in kw_candidates if cands]
    status_text.text(f"Phase 2: Claude validating {len(to_validate):,} keyword-URL matches...")

    validated = {}
    if to_validate:
        validated = claude_validate_match(
            api_key, to_validate, biz_desc,
            status_text, progress_bar, 58, 78
        )

    # ── Build final results ──────────────────────────────────────────────
    progress_bar.progress(78)
    results = []
    for i, km in enumerate(kw_meta):
        kw      = km['Keyword']
        match   = validated.get(kw, {"url": None, "confidence": None})
        chosen  = match.get('url') or ''
        conf    = match.get('confidence')

        # Score based on confidence
        if conf == 'High':
            cs = 0.85; source = 'Claude — High confidence'
        elif conf == 'Medium':
            cs = 0.55; source = 'Claude — Medium confidence'
        else:
            cs = 0.0;  source = ''
            chosen = ''

        # GSC score boost if GSC URL matches Claude match
        gs = 0.0
        gu = km.get('_gsc_url','')
        if chosen and gu and chosen.lower().strip() == gu.lower().strip():
            gs = 0.30; source += ' + GSC confirmed'
        elif not chosen and gu and km.get('_gsc_status') in ('Confirmed','Plausible'):
            # GSC has a URL but Claude rejected it — treat as no match
            pass

        fs = round(cs * 0.7 + gs * 0.3, 4)

        results.append({
            'Keyword':       kw,
            'Volume':        km['Volume'],
            'Your Position': km['Your Position'],
            'Landing Page':  chosen,
            'Intent':        km['Intent'],
            'Content Score': round(cs, 4),
            'GSC Score':     gs,
            'Final Score':   fs,
            'Match Source':  source,
            'Ranking Status':km['Ranking Status'],
            '_cats':         frozenset(),   # kept for compatibility, no longer used for filtering
        })

    progress_bar.progress(80)
    return results


# ── PHASE 3: INTENT CLASSIFICATION ───────────────────────────────────────
def phase3_intent(keywords, api_key, biz_desc, status_text, progress_bar):
    """
    Claude-first intent classification.
    Rules only catch 100% obvious signals — everything else goes to Claude.
    Claude re-examines even rule-classified Transactional keywords that may be
    informational in context. No "Ambiguous" category — Claude must decide.
    """
    status_text.text("Phase 3: Classifying keyword intent...")
    progress_bar.progress(66)

    intent_map  = {}
    for_claude  = []

    for kw in keywords:
        kl = kw.lower().strip()
        # 100% obvious informational: starts with question word
        if re.match(r'^(how|why|what|where|when|which|is|are|can|do|does|will|should|who)', kl):
            intent_map[kl] = 'Informational'
        # 100% obvious transactional: clear hire/buy signals
        elif any(t in kl for t in ['near me', 'near by', 'close to me', 'emergency ',
                                    '24 hour', '24/7', 'same day', 'hire a', 'hire an',
                                    'book a', 'schedule a', 'call a', 'find a']):
            intent_map[kl] = 'Transactional'
        else:
            # Everything else → Claude decides
            for_claude.append(kw)

    progress_bar.progress(68)
    status_text.text(f"Phase 3: {len(intent_map):,} rule-classified. "
                     f"Sending {len(for_claude):,} to Claude for intent judgment...")

    if for_claude:
        # Use existing claude_intent with updated prompt context
        claude_intents = claude_intent(api_key, for_claude, biz_desc,
                                       status_text, progress_bar, 68, 75)
        for kw, intent in claude_intents.items():
            val = intent if intent in ('Informational', 'Transactional') else 'Transactional'
            intent_map[kw.lower()] = val

    progress_bar.progress(75)
    return intent_map

# ── PHASE 4: BUSINESS RELEVANCE + BLOG FIXES ──────────────────────────────
def phase4_relevance_blogs(mapped, url_df, api_key, biz_desc, excl_str, status_text, progress_bar):
    """
    Phase 4: Business relevance check only.
    Blog validation and semantic blog matching removed — Claude already
    validated all keyword-URL matches with High/Medium confidence in Phase 2.
    Any match that survived Phase 2 is already trustworthy.
    """
    progress_bar.progress(82)
    status_text.text("Phase 4: Checking business relevance...")

    # Keywords with a Phase 2 Claude match are confirmed relevant
    confirmed = {r['Keyword']: 'RELEVANT'
                 for r in mapped if r['Landing Page'] and r['Final Score'] > 0}

    # Keywords without a match need relevance judgment
    unmatched = [r['Keyword'] for r in mapped if not r['Landing Page']]

    rel_map = {}
    if unmatched:
        rel_map = claude_relevance(api_key, unmatched, biz_desc, excl_str,
                                   status_text, progress_bar, 82, 92)

    all_rel = {**confirmed, **rel_map}
    progress_bar.progress(92)
    return all_rel, mapped

# ── HELPERS ──────────────────────────────────────────────────────────────
def subtheme_fallback(kw, theme):
    """Guaranteed sub-theme from keyword terms. Never returns blank."""
    kl = kw.lower()
    entity = theme.split('/')[0].strip() if '/' in theme else theme
    if entity in ('Other', ''): entity = 'General'
    # Job-to-be-done signals
    if any(t in kl for t in ['cost','price','how much','pricing','average cost','worth']): job = 'Cost Guide'
    elif any(t in kl for t in ['not working','not heating','not cooling','not draining',
                                 'won\'t','doesn\'t','blowing warm','blows warm','blowing hot',
                                 'blows hot','no heat','no hot water','no cool','no cold',
                                 'keeps','tripping','flickering','smell','noise','leak',
                                 'dripping','gurgling','humming','clicking','buzzing']): job = 'Troubleshooting'
    elif any(t in kl for t in ['how to','diy','myself','my own','at home']): job = 'DIY Guide'
    elif any(t in kl for t in ['what is','how does','how long','lifespan','last','life expectancy',
                                 'types of','difference','vs ','versus','benefits','signs','causes']): job = 'Educational Guide'
    elif any(t in kl for t in ['install','installation','replace','replacement','new ']): job = 'Installation'
    elif any(t in kl for t in ['maintenance','tune','tune-up','tune up','service','clean','flush','inspect']): job = 'Maintenance'
    elif any(t in kl for t in ['emergency','urgent','24 hour','24/7','same day']): job = 'Emergency Service'
    elif any(t in kl for t in ['repair','fix','fixing','broken']): job = 'Repair'
    else: job = 'Service'
    return f"{entity} {job}"

def build_content_groups(mapped):
    """
    Assign Content Group ID to every keyword row.
    Logic: Theme + Normalised Sub-theme + URL (or gap) = one content group.
    Same combination = same page. Primary = highest volume in group.
    Format: AC-001, FH-002, DS-003 etc.
    """
    # Build theme initials map
    INITIALS = {
        'AC / Cooling': 'AC', 'Furnace / Heating': 'FH', 'Heat Pump': 'HP',
        'Water Heater': 'WH', 'Drain / Sewer': 'DS', 'Plumbing': 'PL',
        'Electrical': 'EL', 'Generator': 'GN', 'Sump Pump': 'SP',
        'Air Quality': 'AQ', 'Water Treatment': 'WT', 'Backflow': 'BF',
        'Geothermal': 'GE', 'Bathroom / Kitchen': 'BK', 'Ventilation / Duct': 'VD',
        'EV Charger': 'EV', 'Commercial': 'CM', 'HVAC General': 'HV',
        'Gas Line': 'GL', 'Ductless Mini Split': 'DM', 'Sprinkler': 'SK',
        'Other': 'OT',
    }

    # Group keywords by (theme, subtheme, url_key)
    # url_key: actual URL for mapped kws, 'GAP-{intent}' for unmapped
    groups = {}
    for r in mapped:
        theme = r.get('Theme', 'Other') or 'Other'
        sub   = r.get('Sub-theme', '') or subtheme_fallback(r['Keyword'], theme)
        url   = r.get('Landing Page', '')
        if url:
            url_key = url.lower().strip()
        else:
            url_key = f"GAP-{r.get('Intent','Transactional')}"

        key = (theme, sub, url_key)
        if key not in groups:
            groups[key] = []
        groups[key].append(r)

    # Sort groups by total volume descending, assign IDs
    sorted_groups = sorted(groups.items(),
                           key=lambda x: sum(r['Volume'] for r in x[1]),
                           reverse=True)

    # Track ID counter per theme initial
    id_counter = {}
    cg_map = {}  # keyword → (cg_id, is_primary)

    for (theme, sub, url_key), rows in sorted_groups:
        initial = INITIALS.get(theme, 'XX')
        id_counter[initial] = id_counter.get(initial, 0) + 1
        cg_id = f"{initial}-{id_counter[initial]:03d}"

        # Primary = highest volume in group
        rows_sorted = sorted(rows, key=lambda r: -r.get('Volume', 0))
        for i, r in enumerate(rows_sorted):
            cg_map[r['Keyword']] = (cg_id, 'Primary' if i == 0 else 'Secondary')

    # Apply back to mapped
    for r in mapped:
        cg_id, role = cg_map.get(r['Keyword'], ('', ''))
        r['Content Group'] = cg_id
        r['KW Role']       = role

    return mapped

# ── PHASE 5: THEME + SUB-THEME — FULLY AI DRIVEN ─────────────────────────
def cluster_for_batching(keywords, n_clusters=None):
    """TF-IDF KMeans clustering — universal pre-sorter for Phase 5.
    Works for any industry. Returns {keyword: cluster_id}."""
    from sklearn.cluster import KMeans as _KMeans
    if len(keywords) < 5:
        return {kw: 0 for kw in keywords}
    n = n_clusters or max(5, min(40, len(keywords) // 15))
    n = min(n, len(keywords) - 1)
    try:
        vec = TfidfVectorizer(ngram_range=(1,2), min_df=1,
                              max_features=8000, sublinear_tf=True)
        X   = vec.fit_transform(keywords)
        km  = _KMeans(n_clusters=n, random_state=42, n_init=5, max_iter=100)
        labels = km.fit_predict(X)
        return {kw: int(lbl) for kw, lbl in zip(keywords, labels)}
    except Exception:
        # Fallback: group by first word
        groups = {}
        for kw in keywords:
            key = kw.lower().split()[0] if kw.strip() else '0'
            if key not in groups: groups[key] = len(groups)
        return {kw: groups.get(kw.lower().split()[0] if kw.strip() else '0', 0)
                for kw in keywords}

def phase5_themes(mapped, url_df, api_key, biz_desc, taxonomy, status_text, progress_bar):
    """
    AI-powered keyword grouping. Claude groups keywords into content clusters.
    Each cluster = one piece of content = one sub-theme + one content type.
    Pre-sort: TF-IDF KMeans (universal). Fallback: per-keyword assignment.
    """
    from collections import defaultdict
    status_text.text("Phase 5: AI-powered content grouping...")
    progress_bar.progress(93)

    all_kws = [r['Keyword'] for r in mapped]

    # KMeans pre-sort — groups similar keywords together before sending to Claude
    status_text.text("Phase 5: Clustering keywords by similarity...")
    kw_cluster = cluster_for_batching(all_kws)

    buckets = defaultdict(list)
    for r in mapped:
        buckets[kw_cluster.get(r['Keyword'], 0)].append({
            'keyword': r['Keyword'],
            'intent':  r['Intent'],
            'volume':  r['Volume'],
        })

    group_map  = {}   # keyword → (theme, subtheme, group_type)
    total_b    = len(buckets)
    success_ct = 0
    taxonomy_block = taxonomy_to_prompt_block(taxonomy)

    for bi, (bucket_id, items) in enumerate(buckets.items()):
        pct = min(93 + int((bi / max(total_b, 1)) * 3), 95)
        progress_bar.progress(pct)
        status_text.text(f"Phase 5: Grouping keywords ({bi+1}/{total_b})...")

        for batch_start in range(0, len(items), 80):
            batch   = items[batch_start:batch_start+80]
            kw_list = [{"keyword": x['keyword'], "intent": x['intent'],
                        "volume": x['volume']} for x in batch]

            prompt = f"""SEO strategist. Group these keywords into content clusters.
Each cluster = ONE piece of content (one service page OR one blog post).

Business: {biz_desc}
{taxonomy_block}

RULES:
1. Keywords with the same searcher job → same cluster
2. Transactional (hire/buy) → content_type: "Service page"
3. Informational (learn/research/diagnose) → content_type: "Blog post"
4. NEVER mix content types in one cluster
5. cluster_name: specific, max 4 words, title case
   Good: "Furnace Repair", "AC Troubleshooting", "AC Cost Guide"
   Bad:  "AC Service", "HVAC", "Repair" (too generic)
6. theme: top-level service category from taxonomy

Return ONLY valid JSON array:
[{{"cluster_name":"X","theme":"Y","content_type":"Service page or Blog post",
  "keywords":["kw1","kw2"]}}]

Keywords:
{json.dumps(kw_list)}"""

            for attempt in range(3):
                try:
                    result = call_claude(api_key, prompt, 3000, 55)
                    if isinstance(result, list) and result:
                        for cluster in result:
                            cname  = cluster.get('cluster_name', '')
                            ctheme = cluster.get('theme', '')
                            ctype  = cluster.get('content_type', 'Service page')
                            for kw in cluster.get('keywords', []):
                                group_map[kw] = (ctheme, cname, ctype)
                        success_ct += 1
                    break
                except Exception:
                    if attempt < 2:
                        time.sleep(2)
                    else:
                        for item in batch:
                            kw = item['keyword']
                            if kw not in group_map:
                                group_map[kw] = (
                                    classify_theme(kw), '',
                                    'Blog post' if item['intent'] == 'Informational'
                                    else 'Service page')
            time.sleep(0.1)

    phase5_succeeded = success_ct > (total_b * 0.5)

    for r in mapped:
        kw = r['Keyword']
        if kw in group_map:
            theme, subtheme, gtype = group_map[kw]
            r['Theme']       = theme or classify_theme(kw)
            r['Sub-theme']   = subtheme
            r['_group_type'] = gtype
        else:
            r['Theme']       = classify_theme(kw)
            r['Sub-theme']   = ''
            r['_group_type'] = ('Blog post' if r['Intent'] == 'Informational'
                                else 'Service page')
        r['_phase5_grouped'] = phase5_succeeded

    progress_bar.progress(96)
    return mapped


# ── PHASE 6: SEMANTIC CLUSTERING ──────────────────────────────────────────
def phase6_cluster(mapped, rel_map, api_key, status_text, progress_bar):
    status_text.text("Phase 6: Semantic clustering...")
    progress_bar.progress(97)
    unmapped_info  = sorted([(r['Keyword'], r['Volume']) for r in mapped
                              if not r['Landing Page'] and rel_map.get(r['Keyword'].lower(), '') in ('RELEVANT','BORDERLINE')
                              and r['Intent'] == 'Informational'], key=lambda x: -x[1])[:800]
    unmapped_trans = sorted([(r['Keyword'], r['Volume']) for r in mapped
                              if not r['Landing Page'] and rel_map.get(r['Keyword'].lower(), '') in ('RELEVANT','BORDERLINE')
                              and r['Intent'] == 'Transactional'], key=lambda x: -x[1])[:1200]
    clusters = []
    if unmapped_info:
        clusters += claude_cluster(api_key, unmapped_info, "blog post", status_text, progress_bar, 97, 98)
    if unmapped_trans:
        clusters += claude_cluster(api_key, unmapped_trans, "service page", status_text, progress_bar, 98, 99)

    # ── Write cluster names back to every keyword row as Sub-theme ────────
    # Build lookup: keyword → (cluster_name, entity)
    cluster_kw_lookup = {}
    for cl in clusters:
        cname  = cl.get('cluster_name', '')
        entity = cl.get('entity', '')
        theme  = entity if entity else ''
        for kw2 in [cl.get('primary_keyword','')] + cl.get('secondary_keywords',[]):
            if kw2:
                cluster_kw_lookup[kw2.lower()] = (theme, cname)

    # Fallback sub-theme from keyword terms for unclustered unmapped keywords
    def keyword_fallback_subtheme(kw):
        """Generate a readable sub-theme from the keyword itself."""
        kl = kw.lower()
        # Service type signals
        if any(t in kl for t in ['repair','fix','fixing','broken','not working']): svc = 'Repair'
        elif any(t in kl for t in ['install','installation','replace','replacement','new']): svc = 'Installation'
        elif any(t in kl for t in ['maintenance','tune','service','clean','flush','inspect']): svc = 'Maintenance'
        elif any(t in kl for t in ['cost','price','how much','pricing','average']): svc = 'Cost Guide'
        elif any(t in kl for t in ['smell','odor','stink']): svc = 'Odor Solutions'
        elif any(t in kl for t in ['clog','clogged','unclog','blockage']): svc = 'Clog Solutions'
        elif any(t in kl for t in ['emergency','urgent','24 hour','same day']): svc = 'Emergency Service'
        elif any(t in kl for t in ['near me','local','in my area','close to']): svc = 'Local Service'
        elif any(t in kl for t in ['what is','how does','how to','why','what are']): svc = 'Guide'
        else: svc = 'Service'
        # Entity
        topic = classify_theme(kw)
        short = topic.split('/')[0].strip() if '/' in topic else topic
        return f"{short} {svc}" if short != 'Other' else svc

    # Apply cluster sub-theme (or fallback) to all unmapped keyword rows
    for r in mapped:
        if not r['Landing Page']:
            kl = r['Keyword'].lower()
            if kl in cluster_kw_lookup:
                theme_cl, sub_cl = cluster_kw_lookup[kl]
                if not r.get('Theme') or r.get('Theme') == 'Other':
                    r['Theme'] = theme_cl or classify_theme(r['Keyword'])
                r['Sub-theme'] = sub_cl
            else:
                # Fallback for keywords outside the volume cap
                if not r.get('Sub-theme'):
                    r['Sub-theme'] = keyword_fallback_subtheme(r['Keyword'])

    # Build URL-based clusters (mapped keywords grouped by URL + intent)
    url_groups = {}
    for r in mapped:
        if r['Landing Page']:
            key = (r['Landing Page'], r['Intent'])
            url_groups.setdefault(key, []).append(r)
    url_clusters = []
    for (url, intent), rows in url_groups.items():
        rows_s = sorted(rows, key=lambda x: (-x.get('Final Score', 0), -x.get('Volume', 0)))
        theme, sub = url_to_theme_subtheme(url)
        if not theme: theme = rows_s[0].get('Theme', '')
        if not sub:   sub   = rows_s[0].get('Sub-theme', '')
        url_clusters.append({
            'URL': url, 'Theme': theme, 'Sub-theme': sub, 'Intent': intent,
            'Primary Keyword': rows_s[0]['Keyword'],
            'Primary Volume':  rows_s[0]['Volume'],
            'Secondary Keywords': ' | '.join(r['Keyword'] for r in rows_s[1:10]),
            'Total Volume': sum(r['Volume'] for r in rows_s),
            'Keyword Count': len(rows_s),
            'Best Score':    rows_s[0].get('Final Score', 0),
            'Content Type':  'Blog post' if '/blog/' in url else 'Service page',
        })
    progress_bar.progress(99)
    return clusters, url_clusters


# ── CONTENT GROUP ASSIGNMENT ──────────────────────────────────────────────
THEME_CODES = {
    'Furnace / Heating':'FH','AC / Cooling':'AC','Heat Pump':'HP',
    'Water Heater':'WH','Drain / Sewer':'DS','Plumbing':'PL',
    'Electrical':'EL','Generator':'GN','Sump Pump':'SP',
    'Air Quality':'AQ','Water Treatment':'WT','Backflow':'BF',
    'Geothermal':'GT','Bathroom / Kitchen':'BK','Ventilation / Duct':'VD',
    'EV Charger':'EV','Commercial':'CM','HVAC General':'HG',
    'Gas Line':'GL','Ductless Mini Split':'DM','Other':'OT',
}

# Words that are NEVER location names
US_STATES = {
    'al','ak','az','ar','ca','co','ct','de','fl','ga',
    'hi','id','il','in','ia','ks','ky','la','me','md',
    'ma','mi','mn','ms','mo','mt','ne','nv','nh','nj',
    'nm','ny','nc','nd','oh','ok','or','pa','ri','sc',
    'sd','tn','tx','ut','vt','va','wa','wv','wi','wy','dc',
}
def extract_location(kw):
    """
    Detect specific geographic location in a keyword.
    Handles:
      1. Explicit: "furnace repair in chicago" / "plumber near dallas"
      2. Service + city + state: "plumber clifton park ny"
      3. Service + multi-word city: "hvac repair west palm beach"
      4. City + state abbreviation at end: "drain cleaning buffalo ny"
    Returns location string or empty string.
    """
    kl = kw.lower().strip()

    # Common words that are NEVER location names
    NOT_LOC = {
        'me','area','my','the','your','our','home','house','local','state',
        'repair','repairs','service','services','install','installation',
        'cleaning','clean','company','companies','contractor','contractors',
        'plumber','electrician','technician','professional','emergency',
        'maintenance','replacement','near','local','best','top','cheap',
        'affordable','licensed','certified','residential','commercial',
    }

    SERVICE_WORDS = {
        'repair','repairs','service','services','install','installation',
        'cleaning','clean','company','companies','contractor','contractors',
        'plumber','plumbers','electrician','electricians','technician',
        'technicians','maintenance','replacement','emergency','near','local',
        'hvac','drain','sewer','furnace','boiler','heater','heaters','plumbing',
        'electrical','generator','pump','heating','cooling','water','tankless',
        'ac','mini','split','ductless','backflow','sump','sprinkler',
    }

    words = kl.split()

    # ── Pattern 1: "... in/near/for [city] [optional state]" ─────────────
    # STRICT: only match when the candidate looks like a proper place name
    # Reject: descriptive phrases like "cold winters", "cold weather", "leaking water heater"
    m = re.search(r'\b(?:in|near|for)\s+([a-z][a-z\s]{2,30})$', kl)
    if m:
        candidate = m.group(1).strip()
        cwords = candidate.split()
        # Remove trailing state abbreviation for rejection check
        core_words = cwords[:-1] if (len(cwords) > 1 and cwords[-1] in US_STATES) else cwords
        # Big reject list — any word here means it's NOT a city name
        _REJECT_P1 = NOT_LOC | {
            'come','from','does','pump','tank','pipe','line','unit','also',
            'this','that','what','when','will','with','them','they','then',
            'more','most','some','here','have','been','both','only','sump',
            'cold','warm','hot','cool','heat','weather','winter','winters',
            'summer','spring','springs','fall','season','temperature','temp',
            'leaking','dripping','noise','banging','humming','clicking','signs',
            'settings','setting','issues','problems','going','working','broken',
            'thermostat','pressure','flow','power','energy','money','savings',
            'cost','costs','price','prices','years','days','time','times',
            'small','large','high','low','best','good','bad','common','normal',
        }
        if not any(w in _REJECT_P1 for w in core_words):
            return candidate

    # ── Pattern 2: "[service] [city words] [state abbrev]" ───────────────
    # e.g. "plumber clifton park ny", "ac repair dallas tx"
    # Last word is a 2-letter state abbreviation
    if len(words) >= 3 and words[-1] in US_STATES:
        state = words[-1]
        # Find where service words end
        first_non_service = len(words)
        for i, w in enumerate(words):
            if w not in SERVICE_WORDS and len(w) >= 3 and w not in NOT_LOC:
                first_non_service = i
                break
        # City = words between first non-service word and state abbreviation
        city_words = words[first_non_service:-1]
        if city_words:
            candidate = ' '.join(city_words) + ' ' + state
            if not any(w in NOT_LOC for w in city_words):
                return candidate

    # ── Pattern 3: "[service] [city name]" — city at tail, no state abbrev ──
    # e.g. "hvac repair rochester", "sewer cleaning greenlawn"
    # STRICT: only fire when the tail word is very unlikely to be a common word
    # Most false positives (system, works, function, uses, explained, water)
    # are common English dictionary words — reject them with comprehensive blocklist
    has_service_start = any(w in SERVICE_WORDS for w in words[:2])
    if has_service_start and len(words) >= 3:
        # Large blocklist of common English words that are NOT place names
        COMMON_ENGLISH = {
            # Descriptive/functional words
            'system','systems','works','working','function','functions',
            'uses','used','using','explained','explanation','definition',
            'meaning','purpose','overview','basics','basics','guide',
            'tips','advice','information','info','details','facts',
            'types','kind','kinds','form','forms','style','styles',
            'cost','costs','price','prices','pricing','rates','rate',
            'size','sizes','capacity','power','pressure','flow',
            'water','heat','cool','warm','cold','hot','fire',
            'issues','issue','problems','problem','call','calls',
            'work','help','need','wants','want','gets','make',
            'best','good','great','safe','safe','right','wrong',
            'long','last','life','time','test','check','know',
            'install','repair','clean','replace','maintain','check',
            'pump','tank','pipe','line','unit','part','parts',
            'code','codes','permit','permits','license','licensed',
            'reviews','review','ratings','rated','near','around',
            'professional','certified','trained','qualified','expert',
            'old','new','used','average','normal','standard',
            'high','low','full','small','large','heavy','light',
            'electric','electrical','gas','water','solar','smart',
            # 'city','town','village' handled separately as CITY_SUFFIXES in Pattern 3
            # Additional words that are NEVER place names
            'settings','setting','temperature','noise','banging','signs','signal',
            'signals','leaking','dripping','humming','clicking','buzzing','increase',
            'decrease','improve','repair','broken','working','going','coming',
            'thermostat','pressure','heater','weather','winters','summer','spring',
            'energy','money','savings','power','voltage','current','circuit',
            # Question/informational words
            'what','when','where','which','that','this','these',
            'from','come','goes','goes','does','have','make',
        }
        tail = []
        # Words allowed at end of city name (transparent — don't break traversal)
        CITY_SUFFIXES = {'city','town','village','heights','park','beach','springs',
                         'falls','creek','lake','hill','hills','grove','point','bay'}
        # Common words that appear AS PART OF city names (transparent both directions)
        CITY_PARTS = {'new','old','san','los','las','las','el','le','la','du','des',
                      'west','east','north','south','upper','lower','port','fort',
                      'mount','palm','long','grand','great','little','saint','ste',
                      'isle','bay','cape','lake','rio','del','von','van'}
        for w in reversed(words):
            if w in CITY_SUFFIXES or w in CITY_PARTS:
                tail.insert(0, w)  # include but keep going
                continue
            if w in SERVICE_WORDS or w in NOT_LOC or w in COMMON_ENGLISH:
                break
            if len(w) >= 3 and w.isalpha():
                tail.insert(0, w)
            else:
                break
        if tail:
            candidate = ' '.join(tail)
            # Must have at least one non-suffix, non-city-part word to be a real location
            # (e.g. 'york' in 'new york city' — 'new' is CITY_PARTS, 'city' is CITY_SUFFIXES)
            non_structural = [w for w in tail if w not in CITY_SUFFIXES and w not in CITY_PARTS]
            if non_structural and not any(w in COMMON_ENGLISH for w in non_structural):
                return candidate

    return ''


def normalise_subthemes(api_key, subthemes, biz_desc, status_text, progress_bar):
    """Pass 2: merge near-identical sub-theme variants into canonical names."""
    if not subthemes: return {}
    status_text.text("Phase 5b: Normalising sub-theme names for consistency...")
    progress_bar.progress(96)
    batches = [list(subthemes)[i:i+200] for i in range(0, len(subthemes), 200)]
    canon_map = {}
    for batch in batches:
        prompt = f"""SEO strategist. Normalise these sub-theme names into canonical versions.
Business: {biz_desc}
Rules:
1. Merge variants that mean the SAME content topic into ONE canonical name
2. Keep topics that would go on DIFFERENT pages separate
3. Max 4 words, Title Case
Examples:
  "Furnace Repair", "Furnace Repair Service", "Gas Furnace Repair" → all become "Furnace Repair"
  "AC Tune-up", "AC Tune Up", "Air Conditioner Tune-up" → all become "AC Tune-up"
  "Furnace Repair" vs "Furnace Installation" → KEEP SEPARATE
  "Furnace Cost Guide" vs "AC Cost Guide" → KEEP SEPARATE
Return ONLY JSON: {{"input_name": "Canonical Name", ...}}
Sub-themes: {json.dumps(batch)}"""
        for attempt in range(3):
            try:
                result = call_claude(api_key, prompt, 2000, 50)
                if isinstance(result, dict): canon_map.update(result)
                break
            except Exception:
                if attempt < 2: time.sleep(2)
                else:
                    for s in batch: canon_map[s] = s
        time.sleep(0.1)
    for s in subthemes:
        if s not in canon_map: canon_map[s] = s
    return canon_map

def assign_content_groups(mapped, api_key, biz_desc, status_text, progress_bar):
    """
    Assigns Content Group ID, group-level Content Type, Primary Keyword flag.
    Rules:
    1. One Content Group = One page = One action (service page OR blog post, never both)
    2. Location keywords get their own location groups (DS-LOC-001 etc.)
    3. Primary = highest volume, prefer no generic location qualifier (near me)
    4. Action at GROUP level = majority intent vote within the group
    """
    from collections import Counter, defaultdict

    # Step 1: Normalise sub-themes
    # Guard: skip if Phase 5 AI grouping produced clean cluster names
    phase5_grouped = any(r.get('_phase5_grouped') for r in mapped)
    if not phase5_grouped:
        unique_subs = set(r.get('Sub-theme','') for r in mapped if r.get('Sub-theme',''))
        canon_map   = normalise_subthemes(api_key, unique_subs, biz_desc, status_text, progress_bar)
        for r in mapped:
            r['Sub-theme'] = canon_map.get(r.get('Sub-theme',''), r.get('Sub-theme',''))

    # Step 2: Tag each keyword with its location using pattern-based detection
    for r in mapped:
        r['_loc'] = extract_location(r['Keyword'])

    # Step 4: Build groups — location keywords keyed by their specific location
    groups = {}
    for r in mapped:
        theme = r.get('Theme','Other')
        sub   = r.get('Sub-theme','')
        url   = r.get('Landing Page','') or 'GAP'
        loc   = r.get('_loc','')
        key   = (theme, sub, url, loc)
        groups.setdefault(key, []).append(r)

    # Step 5: Determine group-level content type (majority intent by volume)
    def grp_type(rows):
        tv = sum(r['Volume'] for r in rows if r['Intent']=='Transactional')
        iv = sum(r['Volume'] for r in rows if r['Intent']=='Informational')
        return 'Service page' if tv >= iv else 'Blog post'

    # Step 6: Assign IDs
    theme_groups = defaultdict(list)
    for key, rows in groups.items():
        theme_groups[key[0]].append((key, sum(r['Volume'] for r in rows)))

    cg_map = {}; cg_type_map = {}
    for theme, glist in theme_groups.items():
        code = THEME_CODES.get(theme, 'OT')
        loc_seq = 1; gen_seq = 1
        for key, _ in sorted(glist, key=lambda x: -x[1]):
            loc  = key[3]
            rows = groups[key]
            gt   = grp_type(rows)
            if loc:
                cg_id = f"{code}-LOC-{loc_seq:03d}"; loc_seq += 1
            else:
                cg_id = f"{code}-{gen_seq:03d}"; gen_seq += 1
            cg_map[key] = cg_id; cg_type_map[key] = gt

    # Step 7: Write to rows
    for r in mapped:
        key = (r.get('Theme','Other'), r.get('Sub-theme',''),
               r.get('Landing Page','') or 'GAP', r.get('_loc',''))
        r['Content Group']       = cg_map.get(key, '')
        # Guard: don't overwrite _group_type if Phase 5 already set it
        if not r.get('_group_type'):
            r['_group_type'] = cg_type_map.get(key, '')
        r['_is_loc_group']       = bool(r.get('_loc',''))
        r['Primary Keyword']     = ''

    # Step 8: Flag PRIMARY — highest volume, prefer clean keyword (no "near me")
    GENERIC_LOC = ['near me','in my area','close to me']
    for key, rows in groups.items():
        rows_s = sorted(rows, key=lambda x: (
            -x['Volume'],
            1 if any(t in x['Keyword'].lower() for t in GENERIC_LOC) else 0
        ))
        rows_s[0]['Primary Keyword'] = 'PRIMARY'

    progress_bar.progress(97)
    return mapped


def is_location_url(url):
    """
    Check if a URL is a dedicated location page.
    e.g. /buffalo/furnace-repair/ or /chicago/plumber/ → True
         /furnace-repair/ or /blog/post/ → False
    """
    if not url: return False
    path = re.sub(r'https?://[^/]+', '', url.lower()).strip('/')
    segments = [s for s in path.split('/') if s]
    SERVICE_SLUGS = {'repair','service','install','maintenance','emergency','commercial',
                     'residential','plumbing','heating','cooling','electrical','drain',
                     'sewer','water','heater','furnace','boiler','hvac','blog','about',
                     'contact','services','cleaning','replacement','installation',
                     'generator','pump','electric','electrical','ac','blog'}
    if len(segments) >= 2:
        # First segment is a city/area if it's alphabetic and not a service word
        first = re.sub(r'[\-_]', '', segments[0])
        if first.isalpha() and first not in SERVICE_SLUGS and len(first) >= 3:
            return True
    return False

def get_group_action(r, rel):
    """
    Determine action for a keyword based on its Content Group type.
    Location-aware: keywords with location names get location page actions.
    One Content Group = one action. Informational keywords always get blog action.
    """
    url       = r.get('Landing Page', '')
    rs        = r.get('Ranking Status', '')
    src_match = r.get('Match Source', '')
    gtype     = r.get('_group_type', '')
    is_loc    = r.get('_is_loc_group', False)
    kw_intent = r.get('Intent', '')

    # ── Existing page actions ──────────────────────────────────────────────
    if url and rs == 'Ranking p1-10':
        return 'Confirmed existing page', 'Monitor — already ranking well'
    if rs == 'Quick win p11-20':
        return 'Quick win — optimise', 'Optimise title, meta, H1 — almost ranking'
    if rs in ['Weak ranking p21-50', 'Very weak p51-100']:
        return 'Weak ranking', 'Improve page content + internal links'
    if url and 'Claude semantic' in str(src_match):
        return 'Blog exists — optimise', 'Update blog title/meta/H1 for this keyword'
    if url and '/blog/' in url:
        return 'Page exists — optimise', 'Optimise existing blog post for this keyword'
    if url:
        # Informational keyword mapped to any existing page → treat as blog optimisation
        if kw_intent == 'Informational':
            return 'Page exists — optimise', 'Optimise existing blog post for this keyword'
        # Location keyword on existing page
        if is_loc:
            if is_location_url(url):
                return 'Page exists — optimise', 'Optimise existing location page'
            else:
                return 'Page exists — optimise', 'Create dedicated location page for this keyword'
        return 'Page exists — optimise', 'Optimise existing service page for this keyword'

    # ── Gap actions (no existing page) ────────────────────────────────────
    # Rules (in priority order):
    # 1. Informational keyword → ALWAYS blog post (never service/location page)
    #    This is the most important rule — intent defines content type
    # 2. Transactional + location detected → location service page
    # 3. Transactional, no location → service page (use group type if set)
    if rel in ('RELEVANT', 'BORDERLINE'):
        # Rule 1: Informational always → blog post
        if kw_intent == 'Informational':
            return 'Business relevant gap', 'Create new blog post'
        # Rule 2: Transactional + location → location service page
        if is_loc:
            return 'Business relevant gap', 'Create new location service page'
        # Rule 3: Transactional → service page
        effective_type = gtype if gtype else 'Service page'
        if effective_type == 'Blog post':
            return 'Business relevant gap', 'Create new blog post'
        return 'Business relevant gap', 'Create new service page'

    return 'True content gap', 'Evaluate — may need new page'

# ── EXCEL OUTPUT ──────────────────────────────────────────────────────────
def build_excel(gsc_df, mapped, rel_map, clusters, url_clusters, taxonomy=None):
    wb = Workbook()
    bf = Font(size=10); lf = Font(size=10, color='0563C1'); mf = Font(size=10, italic=True, color='888888')
    C = dict(H='1D9E75', GR='EAF3DE', BL='E8F0FE', YL='FFFBEA', RD='FDECEA',
             AM='FFF3CD', PU='F3E8FE', GY='F5F5F5', WH='FFFFFF', DGR='D4EDDA', OR='FFF0E6')

    # TAB 1: GSC Mapping Quality
    ws1 = wb.active; ws1.title = 'GSC Mapping Quality'
    hdr(ws1, 1, ['Query','Mapped URL','Clicks','Impressions','Position','CTR','Content Score','Mapping Status'])
    sf = {'Confirmed': make_fill(C['GR']), 'Plausible': make_fill(C['YL']), 'Suspicious': make_fill(C['RD'])}
    for i, row in gsc_df.iterrows():
        r = i+2; fill = sf.get(row['Mapping Status'], make_fill(C['WH']))
        for col, v in enumerate([row['Query'],row['Mapped URL'],row['Clicks'],row['Impressions'],
                                  row['Position'],row['CTR'],row['Content Score'],row['Mapping Status']], 1):
            c = ws1.cell(row=r, column=col, value=v); c.fill = fill; c.font = lf if col==2 else bf
    cw(ws1, [45,65,10,12,10,8,14,14]); ws1.freeze_panes = 'A2'

    # TAB 2: Keyword Mapping (with Theme + Sub-theme)
    ws2 = wb.create_sheet('Keyword Mapping')
    hdr(ws2, 1, ['Theme','Sub-theme','Content Group','Primary?','Keyword','Volume',
                 'Your Position','Landing Page','Intent','Content Score',
                 'GSC Score','Final Score','Match Source','Ranking Status'])
    for i, r in enumerate(sorted(mapped, key=lambda x: (
            x.get('Theme',''), x.get('Sub-theme',''),
            x.get('Content Group',''), x.get('Primary Keyword','') != 'PRIMARY',
            -x.get('Volume',0)))):
        row = i+2; url = r['Landing Page']; src = r['Match Source']; rel = rel_map.get(r['Keyword'].lower(), '')
        if url and '/blog/' in url:           fill = make_fill(C['BL'])
        elif src == 'GSC fallback':            fill = make_fill(C['YL'])
        elif src and 'Claude semantic' in src: fill = make_fill(C['OR'])
        elif url:                              fill = make_fill(C['GR'])
        elif rel in ('RELEVANT','BORDERLINE'): fill = make_fill(C['PU'])
        else:                                  fill = make_fill(C['WH'])
        is_primary = r.get('Primary Keyword','') == 'PRIMARY'
        vals = [r.get('Theme',''), r.get('Sub-theme',''), r.get('Content Group',''),
                '★ PRIMARY' if is_primary else '', r['Keyword'], r['Volume'],
                r['Your Position'], url or '', r['Intent'],
                r['Content Score'], r['GSC Score'], r['Final Score'], src or '', r['Ranking Status']]
        for col, v in enumerate(vals, 1):
            c = ws2.cell(row=row, column=col, value=v)
            c.fill = fill
            c.font = Font(size=10, bold=is_primary, color='0563C1' if col==8 else '000000')
    cw(ws2, [22,28,12,10,48,12,14,62,15,14,12,12,26,18]); ws2.freeze_panes = 'A2'

    ws3 = wb.create_sheet('Opportunity Classification')
    hdr(ws3, 1, ['Theme','Sub-theme','Content Group','Primary?','Keyword','Volume',
                 'Your Position','Mapped URL','Intent','Final Score','Opportunity Type','Action'])
    OPP_F = {'Confirmed existing page': make_fill(C['GR']), 'Quick win — optimise': make_fill(C['DGR']),
             'Weak ranking': make_fill(C['YL']), 'Page exists — optimise': make_fill(C['BL']),
             'Blog exists — optimise': make_fill(C['OR']), 'Business relevant gap': make_fill(C['PU']),
             'True content gap': make_fill(C['RD'])}
    for r in sorted(mapped, key=lambda x: (
            x.get('Theme',''), x.get('Sub-theme',''), x.get('Content Group',''),
            x.get('Primary Keyword','') != 'PRIMARY', -x.get('Volume',0))):
        rel = rel_map.get(r['Keyword'].lower(), ''); url = r['Landing Page']
        rs  = r['Ranking Status']; fs = r['Final Score']; src = r.get('Match Source', '')
        opp, act = get_group_action(r, rel)
        fill = OPP_F.get(opp, make_fill(C['WH']))
        is_primary = r.get('Primary Keyword','') == 'PRIMARY'
        rn = ws3.max_row + 1
        for col, v in enumerate([r.get('Theme',''), r.get('Sub-theme',''),
                                  r.get('Content Group',''), '★ PRIMARY' if is_primary else '',
                                  r['Keyword'], r['Volume'], r['Your Position'],
                                  url or '', r['Intent'], fs, opp, act], 1):
            c = ws3.cell(row=rn, column=col, value=v)
            c.fill = fill
            c.font = Font(size=10, bold=is_primary, color='0563C1' if col==8 else '000000')
    cw(ws3, [22,28,12,10,48,12,14,62,15,12,28,48]); ws3.freeze_panes = 'A2'

    # TAB 4: Keyword Clusters
    ws4 = wb.create_sheet('Keyword Clusters')
    ws4['A1'] = 'Keyword Clusters — Semantic grouping by entity and intent'
    ws4['A1'].font = Font(bold=True, size=13, color='0F6E56'); ws4.merge_cells('A1:L1')
    ws4.cell(row=3, column=1, value='NEW CONTENT NEEDED').font = Font(bold=True, size=12, color='0F6E56')
    ws4.merge_cells('A3:L3')
    hdr(ws4, 4, ['Theme','Sub-theme','Content Type','Intent','Suggested Title',
                 'Primary Keyword','Primary Volume','Secondary Keywords','Total Volume','Action'], bg='0F6E56')
    r = 5
    for cl in sorted(clusters, key=lambda x: -x.get('total_volume', 0)):
        ct   = cl.get('content_type', 'Blog post')
        fill = make_fill(C['PU']) if 'Blog' in ct else make_fill(C['RD'])
        ent  = cl.get('entity', ''); cln = cl.get('cluster_name', '')
        vals = [cl.get('entity',''), cl.get('cluster_name',''), ct, cl.get('intent',''),
                cl.get('suggested_title',''), cl.get('primary_keyword',''),
                cl.get('primary_volume', 0), ' | '.join(cl.get('secondary_keywords', [])),
                cl.get('total_volume', 0),
                'Create new blog post' if 'Blog' in ct else 'Create new service page']
        for col, v in enumerate(vals, 1):
            c = ws4.cell(row=r, column=col, value=v); c.fill = fill; c.font = bf
        r += 1
    r += 2
    ws4.cell(row=r, column=1, value='EXISTING PAGE OPTIMISATION').font = Font(bold=True, size=12, color='0F6E56')
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10); r += 1
    hdr(ws4, r, ['Theme','Sub-theme','URL','Intent','Primary Keyword','Primary Volume',
                 'Secondary Keywords','Total Volume','# Keywords','Action'], bg='1D9E75'); r += 1
    for cl in sorted(url_clusters, key=lambda x: (x['Theme'], x['Sub-theme'], -x['Total Volume'])):
        fill = make_fill(C['BL']) if 'Blog' in cl['Content Type'] else make_fill(C['GR'])
        vals = [cl['Theme'], cl['Sub-theme'], cl['URL'], cl['Intent'], cl['Primary Keyword'],
                cl['Primary Volume'], cl['Secondary Keywords'], cl['Total Volume'],
                cl['Keyword Count'], 'Optimise for primary keyword — include secondary keywords naturally']
        for col, v in enumerate(vals, 1):
            c = ws4.cell(row=r, column=col, value=v); c.fill = fill; c.font = lf if col==3 else bf
        r += 1
    cw(ws4, [22,28,60,15,42,15,65,14,12,48]); ws4.freeze_panes = 'A5'

    # TAB 5: Business Relevant Gaps
    ws5 = wb.create_sheet('Business Relevant Gaps')
    hdr(ws5, 1, ['Theme','Sub-theme','Content Group','Primary?','Keyword','Volume','Intent','Relevance','Action Needed'])
    gaps = sorted([r for r in mapped if not r['Landing Page']
                   and rel_map.get(r['Keyword'].lower(), '') in ('RELEVANT','BORDERLINE')],
                  key=lambda x: (x.get('Theme',''), x.get('Sub-theme',''),
                                 x.get('Content Group',''),
                                 x.get('Primary Keyword','') != 'PRIMARY', -x.get('Volume',0)))
    for r in gaps:
        rel = rel_map.get(r['Keyword'].lower(), '')
        fill = make_fill(C['PU'] if rel == 'RELEVANT' else C['AM'])
        is_primary = r.get('Primary Keyword','') == 'PRIMARY'
        act = get_group_action(r, rel)[1]
        rn = ws5.max_row + 1
        for col, v in enumerate([r.get('Theme',''), r.get('Sub-theme',''),
                                  r.get('Content Group',''), '★ PRIMARY' if is_primary else '',
                                  r['Keyword'], r['Volume'], r['Intent'], rel, act], 1):
            c = ws5.cell(row=rn, column=col, value=v)
            c.fill = fill; c.font = Font(size=10, bold=is_primary)
    cw(ws5, [22,28,12,10,50,12,15,14,28]); ws5.freeze_panes = 'A2'

    # TAB 6: Priority Roadmap — ONE ROW PER CLUSTER (client-facing strategy)
    ws6 = wb.create_sheet('Priority Roadmap')
    ws6['A1'] = 'SEO Content Strategy Roadmap'
    ws6['A1'].font = Font(bold=True, size=14, color='0F6E56')
    ws6.merge_cells('A1:L1')

    # Summary counts
    qw   = sum(1 for r in mapped if r['Ranking Status'] == 'Quick win p11-20')
    weak = sum(1 for r in mapped if r['Ranking Status'] in ['Weak ranking p21-50','Very weak p51-100'])
    pnr  = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status'] == 'Not ranking')
    bgap = sum(1 for r in mapped if not r['Landing Page']
               and rel_map.get(r['Keyword'].lower(),'') in ('RELEVANT','BORDERLINE'))
    conf = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status'] == 'Ranking p1-10')

    hdr(ws6, 3, ['Opportunity Type','Count','Total Volume','Action','Priority',''])
    summary_rows = [
        ('Quick Wins (p11-20)',      qw,   sum(r['Volume'] for r in mapped if r['Ranking Status']=='Quick win p11-20'), 'Optimise existing pages', 'High', C['DGR']),
        ('Weak Rankings (p21-100)',  weak, sum(r['Volume'] for r in mapped if r['Ranking Status'] in ['Weak ranking p21-50','Very weak p51-100']), 'Improve content + internal links', 'High', C['YL']),
        ('Pages Exist — Not Ranking',pnr,  sum(r['Volume'] for r in mapped if r['Landing Page'] and r['Ranking Status']=='Not ranking'), 'Optimise existing pages', 'Medium', C['BL']),
        ('Business Relevant Gaps',   bgap, sum(r['Volume'] for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'].lower(),'') in ('RELEVANT','BORDERLINE')), 'Create new content', 'Medium', C['PU']),
        ('Already Ranking Well',     conf, sum(r['Volume'] for r in mapped if r['Landing Page'] and r['Ranking Status']=='Ranking p1-10'), 'Monitor only', 'Low', C['GY']),
    ]
    for i, (opp,cnt,vol,act,pri,color) in enumerate(summary_rows):
        r_row = i+4; fill = make_fill(color)
        for col, v in enumerate([opp,cnt,vol,act,pri,''], 1):
            c = ws6.cell(row=r_row, column=col, value=v)
            c.fill = fill; c.font = Font(size=10, bold=(col==5))

    # ── Cluster-level roadmap ─────────────────────────────────────────────
    dr = len(summary_rows) + 7
    ws6.cell(row=dr, column=1,
             value='Content Strategy — One Row Per Cluster').font = Font(bold=True, size=12, color='0F6E56')
    ws6.merge_cells(start_row=dr, start_column=1, end_row=dr, end_column=12); dr += 1

    hdr(ws6, dr, ['Priority','Theme','Sub-theme','Content Group','Content Type',
                  'Primary Keyword','Secondary Keywords','Total Volume','# Keywords',
                  'Action','Mapped URL','Confidence']); dr += 1

    PORD = {'Quick win — optimise':1,'Weak ranking':2,'Page exists — optimise':3,
            'Blog exists — optimise':3,'Business relevant gap':4,'True content gap':5,
            'Confirmed existing page':6}
    AF = {'Quick win — optimise':   make_fill(C['DGR']),
          'Weak ranking':            make_fill(C['YL']),
          'Page exists — optimise':  make_fill(C['BL']),
          'Blog exists — optimise':  make_fill(C['OR']),
          'Business relevant gap':   make_fill(C['PU']),
          'True content gap':        make_fill(C['RD']),
          'Confirmed existing page': make_fill(C['GY'])}

    # Build cluster-level summary from mapped rows
    from collections import defaultdict
    cluster_rows = defaultdict(list)
    for r in mapped:
        cg = r.get('Content Group','')
        if cg:
            cluster_rows[cg].append(r)
        # Keywords without a content group get their own virtual group
        else:
            cluster_rows[f"__{r['Keyword']}__"].append(r)

    cluster_items = []
    for cg, rows in cluster_rows.items():
        # Determine cluster-level action:
        # If ANY keyword has a mapped URL → Optimise that URL
        # Otherwise → Create new
        mapped_rows = [r for r in rows if r['Landing Page']]
        url_for_cluster = ''
        confidence_for_cluster = ''
        if mapped_rows:
            # Pick the highest confidence / highest scoring URL
            best = max(mapped_rows, key=lambda x: x.get('Final Score',0))
            url_for_cluster  = best['Landing Page']
            src_txt = best.get('Match Source','')
            confidence_for_cluster = (
                'High'   if 'High' in src_txt else
                'Medium' if 'Medium' in src_txt else
                'Confirmed' if 'GSC' in src_txt else 'Medium'
            )

        # Group-level action
        rel_any = any(rel_map.get(r['Keyword'].lower(),'') in ('RELEVANT','BORDERLINE') for r in rows)
        first_r = rows[0]
        group_action_opp, group_action_act = get_group_action(
            {**first_r, 'Landing Page': url_for_cluster,
             '_is_loc_group': any(r.get('_is_loc_group') for r in rows)},
            'RELEVANT' if rel_any else 'BORDERLINE'
        )

        # Ranking status for priority
        rs_list = [r['Ranking Status'] for r in rows]
        if any(s == 'Quick win p11-20' for s in rs_list):       po = 1
        elif any(s in ('Weak ranking p21-50','Very weak p51-100') for s in rs_list): po = 2
        elif url_for_cluster:                                    po = 3
        elif rel_any:                                            po = 4
        else:                                                    po = 5
        pri = {1:'High',2:'High',3:'Medium',4:'Medium',5:'Low'}.get(po,'Low')

        # Primary keyword = highest volume
        rows_sorted = sorted(rows, key=lambda x: -x.get('Volume',0))
        primary_kw  = rows_sorted[0]['Keyword']
        secondary   = ' | '.join(r['Keyword'] for r in rows_sorted[1:8])
        total_vol   = sum(r.get('Volume',0) for r in rows)
        theme       = first_r.get('Theme','')
        subtheme    = first_r.get('Sub-theme','')
        gtype       = first_r.get('_group_type','')
        content_type = ('Blog post' if 'Blog' in gtype else
                        'Location page' if any(r.get('_is_loc_group') for r in rows) else
                        'Service page')

        cluster_items.append({
            'cg': cg if not cg.startswith('__') else '',
            'theme': theme, 'subtheme': subtheme,
            'content_type': content_type,
            'primary_kw': primary_kw, 'secondary': secondary,
            'total_vol': total_vol, 'n_kws': len(rows),
            'opp': group_action_opp, 'act': group_action_act,
            'url': url_for_cluster, 'confidence': confidence_for_cluster,
            'pri': pri, 'po': po,
        })

    # Sort: Priority → Theme → Total Volume desc
    cluster_items.sort(key=lambda x: (x['po'], x['theme'], x['subtheme'], -x['total_vol']))

    for item in cluster_items:
        fill = AF.get(item['opp'], make_fill(C['WH']))
        is_optimise = 'Optimise' in item['act'] or 'optimise' in item['act']
        for col, v in enumerate([
            item['pri'], item['theme'], item['subtheme'], item['cg'],
            item['content_type'], item['primary_kw'], item['secondary'],
            item['total_vol'], item['n_kws'], item['act'],
            item['url'] if is_optimise else '',
            item['confidence'] if is_optimise else '',
        ], 1):
            c = ws6.cell(row=dr, column=col, value=v)
            c.fill = fill
            c.font = Font(size=10, bold=(item['po'] <= 2),
                         color='0563C1' if col==11 and v else '000000')
        dr += 1

    cw(ws6, [10,22,28,12,14,42,65,12,10,38,65,12])
    ws6.freeze_panes = f'A{len(summary_rows)+10}'

    # Taxonomy reference tab
    if taxonomy:
        wst = wb.create_sheet('Taxonomy Reference')
        wst['A1'] = 'Locked Taxonomy — used for consistent theme/sub-theme assignment'
        wst['A1'].font = Font(bold=True, size=12, color='0F6E56')
        wst.merge_cells('A1:C1')
        hdr(wst, 2, ['Theme', 'Sub-themes', 'Count'])
        for i, (theme, subs) in enumerate(taxonomy.items()):
            r = i + 3
            wst.cell(row=r, column=1, value=theme).font = Font(size=10, bold=True)
            wst.cell(row=r, column=2, value=', '.join(subs)).font = Font(size=10)
            wst.cell(row=r, column=3, value=len(subs)).font = Font(size=10)
        cw(wst, [28, 80, 8])

    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── UI ────────────────────────────────────────────────────────────────────
st.title("🔍 SEO Content Mapper")
st.markdown("Complete SEO gap analysis — GSC validation + semantic intent + accurate theme clustering.")

with st.sidebar:
    st.header("⚙️ Settings"); st.markdown("---")
    threshold = st.slider("Match Score Threshold", 0.10, 0.40, 0.15, 0.01)
    st.markdown("**Content Weights**")
    st.caption("Page Title, H1 and Meta are primary signals. URL Slug is secondary.")
    tw=st.slider("Page Title",1,8,5); hw=st.slider("H1 Heading",1,6,4)
    mw=st.slider("Meta Description",1,5,3); sw=st.slider("URL Slug",1,4,2)
    weights = (sw,tw,hw,mw)
    st.markdown("---")
    st.markdown("**Score Guide**")
    st.markdown("| Score | Quality |\n|---|---|\n| 0.50+ | Strong ✅ |\n| 0.30–0.50 | Good ✅ |\n| 0.20–0.30 | Acceptable ⚠️ |\n| 0.15–0.20 | Weak — verify |\n| Blank | Content gap |")
    st.markdown("---")
    st.markdown("**Colour Coding**")
    st.markdown("🟢 Content match  🔵 Blog  🟡 GSC fallback  🟠 Claude semantic  🟣 Business gap")

st.markdown('<div class="sec-hdr">📂 File 1 — Screaming Frog Export</div>', unsafe_allow_html=True)
st.caption("Screaming Frog → Bulk Export → All. Required columns: Address, Title 1, Meta Description 1, H1-1")
sf_file = st.file_uploader("Upload Screaming Frog file (.xlsx or .csv)", type=['xlsx','csv'], key='sf')
sf_df = None
if sf_file:
    try:
        sf_df = pd.read_csv(sf_file) if sf_file.name.endswith('.csv') else pd.read_excel(sf_file)
        sf_df.columns = [c.strip() for c in sf_df.columns]
        missing = [c for c in ['Address','Title 1','Meta Description 1','H1-1'] if c not in sf_df.columns]
        if missing:
            st.markdown(f'<div class="warn-box">⚠️ Missing columns: {missing}. Found: {list(sf_df.columns[:8])}</div>', unsafe_allow_html=True); sf_df = None
        else:
            if 'Status Code' in sf_df.columns: sf_df = sf_df[sf_df['Status Code']==200].reset_index(drop=True)
            sf_df = sf_df[sf_df['Address'].notna()].reset_index(drop=True)
            st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(sf_df):,}</strong> pages</div>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)

st.markdown('<div class="sec-hdr">📂 File 2 — GSC Export</div>', unsafe_allow_html=True)
st.caption("Looker Studio export. Required: Query, Landing Page, Clicks, Impressions, Position, CTR")
gsc_file = st.file_uploader("Upload GSC Excel file", type=['xlsx'], key='gsc')
gsc_df = None
if gsc_file:
    try:
        gsc_xl = pd.read_excel(gsc_file, sheet_name=None); gsc_sheets = list(gsc_xl.keys())
        gsc_sheet = st.selectbox("Select GSC sheet", gsc_sheets, key='gs')
        raw = gsc_xl[gsc_sheet]; cols = list(raw.columns)
        with st.expander("Map GSC columns"):
            qc   = st.selectbox("Query",        cols, index=cols.index('Query')        if 'Query'        in cols else 0, key='qc')
            lpc  = st.selectbox("Landing Page", cols, index=cols.index('Landing Page') if 'Landing Page' in cols else 1, key='lpc')
            clkc = st.selectbox("Clicks",       cols, index=next((i for i,c in enumerate(cols) if 'click' in c.lower()),2), key='clkc')
            impc = st.selectbox("Impressions",  cols, index=next((i for i,c in enumerate(cols) if 'impression' in c.lower()),3), key='impc')
            posc = st.selectbox("Position",     cols, index=next((i for i,c in enumerate(cols) if 'position' in c.lower()),min(4,len(cols)-1)), key='posc')
            ctrc = st.selectbox("CTR",          cols, index=next((i for i,c in enumerate(cols) if 'ctr' in c.lower()),min(5,len(cols)-1)), key='ctrc')
        gsc_df = raw.rename(columns={qc:'Query',lpc:'Landing Page',clkc:'Clicks',impc:'Impressions',posc:'Position',ctrc:'CTR'})
        gsc_df['Clicks'] = pd.to_numeric(gsc_df['Clicks'], errors='coerce').fillna(0)
        gsc_df['Impressions'] = pd.to_numeric(gsc_df['Impressions'], errors='coerce').fillna(0)
        st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(gsc_df):,}</strong> GSC rows</div>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)

st.markdown('<div class="sec-hdr">📂 File 3 — Semrush Keyword Gap Export</div>', unsafe_allow_html=True)
st.caption("Full keyword gap export. Required: Keyword, Volume, your domain position column.")
sem_file = st.file_uploader("Upload Semrush file", type=['xlsx'], key='sem')
sem_df = None; your_col = None
if sem_file:
    try:
        sem_xl = pd.read_excel(sem_file, sheet_name=None); sem_sheets = list(sem_xl.keys())
        sem_sheet = st.selectbox("Select Semrush sheet", sem_sheets, key='ss')
        raw = sem_xl[sem_sheet]; cols = list(raw.columns)
        with st.expander("Map Semrush columns"):
            kwc = st.selectbox("Keyword", cols, index=cols.index('Keyword') if 'Keyword' in cols else 0, key='skc')
            vc  = st.selectbox("Volume",  cols, index=cols.index('Volume')  if 'Volume'  in cols else 1, key='svc')
            your_col = st.selectbox("Your domain position column", cols, index=0,
                                    help="Column showing YOUR site's position e.g. yourdomain.com", key='sdc')
        sem_df = raw.rename(columns={kwc:'Keyword', vc:'Volume'})
        sem_df = sem_df[sem_df['Keyword'].notna()].reset_index(drop=True)
        st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(sem_df):,}</strong> keywords</div>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)

st.markdown('<div class="sec-hdr">🏢 Business Context + Claude API</div>', unsafe_allow_html=True)
ca, cb = st.columns(2)
with ca:
    api_key   = st.text_input("Claude API Key", type="password", placeholder="sk-ant-...", help="Required. Get from console.anthropic.com")
    biz_desc  = st.text_area("Business Description *", placeholder="Describe your business, services and location...", height=130)
with cb:
    excl_str  = st.text_area("Services NOT offered (optional — one per line)", placeholder="e.g.\noil boiler\nseptic tank\nwell pump", height=130)
    st.caption("Leave blank if unsure — Claude will use the business description to judge relevance.")

st.markdown("---")
issues = []
if sf_df is None:       issues.append("File 1 (Screaming Frog) not uploaded or has column errors")
if gsc_df is None:      issues.append("File 2 (GSC) not uploaded or has errors")
if sem_df is None:      issues.append("File 3 (Semrush) not uploaded or has errors")
if not api_key.strip(): issues.append("Claude API key is required")
if not biz_desc.strip():issues.append("Business description is required")

for issue in issues:
    st.markdown(f'<div class="warn-box">⚠️ {issue}</div>', unsafe_allow_html=True)

if st.button("🚀 Run Full Analysis", disabled=bool(issues), use_container_width=True, type="primary"):
    progress_bar = st.progress(0)
    status_text  = st.empty()
    t_start      = time.time()

    try:
        # Phase 0: Generate locked taxonomy for this business
        taxonomy, url_theme_map, abbrev_map = generate_taxonomy(
            api_key, biz_desc, sf_df, status_text, progress_bar)

        # Phase 1
        gsc_val, url_df, gsc_dedup, stop, hp_slugs = phase1_gsc(gsc_df, sf_df, weights, status_text, progress_bar)

        # Phase 3: Intent classification (before mapping so mapping uses correct intent)
        all_kws = sem_df['Keyword'].fillna('').tolist()
        intent_map = phase3_intent(all_kws, api_key, biz_desc, status_text, progress_bar)

        # Phase 2
        mapped = phase2_map(sem_df, url_df, gsc_dedup, gsc_val, weights, threshold,
                            your_col, intent_map, hp_slugs, url_theme_map,
                            status_text, progress_bar)

        # Phase 4
        rel_map, mapped = phase4_relevance_blogs(mapped, url_df, api_key, biz_desc, excl_str,
                                                  status_text, progress_bar)

        # Phase 5: Theme + Sub-theme
        mapped = phase5_themes(mapped, url_df, api_key, biz_desc, taxonomy, status_text, progress_bar)

        # Phase 6: Clustering
        clusters, url_clusters = phase6_cluster(mapped, rel_map, api_key, status_text, progress_bar)

        # ── Content Group Assignment ──────────────────────────────────────
        mapped = assign_content_groups(mapped, api_key, biz_desc, status_text, progress_bar)

        # ── GUARANTEED ZERO BLANK SUB-THEMES ─────────────────────────────
        # After all Claude phases, sweep every row and fill any remaining blanks.
        # This runs on mapped list which feeds ALL tabs — no tab can have blanks.
        def guaranteed_subtheme(kw, theme, intent):
            """Generate sub-theme from keyword terms when Claude missed it."""
            kl = kw.lower()
            # Service type from keyword signals
            if any(t in kl for t in ['not working','not heating','not cooling','not turning',
                                      'not draining','blowing warm','blowing cold','wont start',
                                      "won't start","doesn't work",'broken','failed']):
                svc = 'Troubleshooting'
            elif any(t in kl for t in ['cost','price','how much','pricing','average cost',
                                        'worth it','expensive','affordable','cheap']):
                svc = 'Cost Guide'
            elif any(t in kl for t in ['repair','fix','fixing','broken','service call']):
                svc = 'Repair'
            elif any(t in kl for t in ['install','installation','replace','replacement',
                                        'new','put in','set up']):
                svc = 'Installation'
            elif any(t in kl for t in ['maintenance','tune up','tune-up','service',
                                        'clean','flush','inspect','annual','seasonal']):
                svc = 'Maintenance'
            elif any(t in kl for t in ['emergency','urgent','24 hour','24/7','same day',
                                        'after hours']):
                svc = 'Emergency Service'
            elif any(t in kl for t in ['near me','local','in my area','close to me']):
                svc = 'Local Service'
            elif intent == 'Informational':
                if any(t in kl for t in ['how to','how do','how does']):   svc = 'How-to Guide'
                elif any(t in kl for t in ['what is','what are','what does']): svc = 'Explainer'
                elif any(t in kl for t in ['why','cause','reason']):        svc = 'Problem Guide'
                elif any(t in kl for t in ['vs','versus','difference','compare']): svc = 'Comparison'
                elif any(t in kl for t in ['signs','symptoms','detect']):   svc = 'Warning Signs'
                elif any(t in kl for t in ['tips','advice','guide','checklist']): svc = 'Tips Guide'
                else:                                                         svc = 'Guide'
            else:
                svc = 'Service'
            # Entity prefix from theme
            entity = theme.split('/')[0].strip() if '/' in theme else theme
            if entity and entity not in ('Other', ''):
                return f"{entity} {svc}"
            return svc

        blanks_filled = 0
        for r in mapped:
            if not r.get('Sub-theme') or str(r.get('Sub-theme','')).strip() == '':
                r['Sub-theme'] = guaranteed_subtheme(
                    r['Keyword'], r.get('Theme', classify_theme(r['Keyword'])), r['Intent'])
                blanks_filled += 1
            # Also ensure Theme is never blank
            if not r.get('Theme') or str(r.get('Theme','')).strip() == '':
                r['Theme'] = classify_theme(r['Keyword'])

        progress_bar.progress(99)
        status_text.text("Building Excel output...")
        excel_buf = build_excel(gsc_val, mapped, rel_map, clusters, url_clusters, taxonomy)
        progress_bar.progress(100)
        elapsed = round(time.time() - t_start)
        status_text.text(f"✅ Done in {elapsed//60}m {elapsed%60}s")

        # Metrics
        st.markdown("---"); st.subheader("📊 Analysis Complete")
        mapped_kw = sum(1 for r in mapped if r['Landing Page'])
        qw        = sum(1 for r in mapped if r['Ranking Status'] == 'Quick win p11-20')
        bgaps     = sum(1 for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'].lower(),'') in ('RELEVANT','BORDERLINE'))
        themes    = len(set(r.get('Theme','') for r in mapped if r.get('Theme','')))
        n_cl      = len(clusters)
        conf_gsc  = len(gsc_val[gsc_val['Mapping Status']=='Confirmed'])

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        for col,num,label in [(c1,conf_gsc,'GSC confirmed'),(c2,mapped_kw,'Keywords mapped'),
                               (c3,qw,'Quick wins'),(c4,bgaps,'Business gaps'),
                               (c5,themes,'Themes found'),(c6,n_cl,'Topic clusters')]:
            col.markdown(f'<div class="metric-card"><div class="metric-num">{num:,}</div><div class="metric-label">{label}</div></div>', unsafe_allow_html=True)

        st.markdown("")
        t1,t2,t3,t4 = st.tabs(["GSC Validation","Keyword Mapping","Topic Clusters","Business Gaps"])
        with t1: st.dataframe(gsc_val.head(200), use_container_width=True, height=350)
        with t2:
            prev = pd.DataFrame([{'Theme':r.get('Theme',''),'Sub-theme':r.get('Sub-theme',''),
                                   'Keyword':r['Keyword'],'Volume':r['Volume'],
                                   'Landing Page':r['Landing Page'],'Intent':r['Intent'],
                                   'Final Score':r['Final Score']} for r in mapped]).sort_values(
                                   ['Theme','Sub-theme','Final Score'], ascending=[True,True,False])
            st.dataframe(prev.head(300), use_container_width=True, height=350)
        with t3:
            if clusters:
                cl_prev = pd.DataFrame([{'Theme':c.get('entity',''),'Sub-theme':c.get('cluster_name',''),
                                          'Type':c.get('content_type',''),'Title':c.get('suggested_title',''),
                                          'Primary KW':c.get('primary_keyword',''),
                                          'Total Volume':c.get('total_volume',0)} for c in clusters]).sort_values('Total Volume', ascending=False)
                st.dataframe(cl_prev, use_container_width=True, height=350)
        with t4:
            gp = pd.DataFrame([{'Theme':r.get('Theme',''),'Sub-theme':r.get('Sub-theme',''),
                                  'Keyword':r['Keyword'],'Volume':r['Volume'],'Intent':r['Intent'],
                                  'Relevance':rel_map.get(r['Keyword'].lower(),'')} for r in mapped
                                 if not r['Landing Page'] and rel_map.get(r['Keyword'].lower(),'') in ('RELEVANT','BORDERLINE')]).sort_values(['Theme','Volume'], ascending=[True,False])
            st.dataframe(gp.head(300), use_container_width=True, height=350)

        st.markdown("---")
        st.download_button("⬇️ Download Full Excel Output (6 tabs)", data=excel_buf,
                           file_name="seo_content_mapping_output.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
        st.markdown('<div class="info-box"><strong>6 tabs:</strong> GSC Mapping Quality | Keyword Mapping | Opportunity Classification | Keyword Clusters | Business Relevant Gaps | Priority Roadmap — all sorted by Theme → Sub-theme</div>', unsafe_allow_html=True)

    except Exception as e:
        elapsed = round(time.time() - t_start)
        progress_bar.progress(0)
        st.markdown(f'<div class="error-box">❌ Error after {elapsed}s: {str(e)}</div>', unsafe_allow_html=True)
        st.exception(e)
