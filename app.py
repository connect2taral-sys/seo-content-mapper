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

# FIX 2: /tag/ added
EXCL = ['/truck-wrap','/careers','/join-the-team','/contact','/podcast','/coupons',
        '/warranty','/about-us','/meet-the','/technicians','/community','/financing',
        '/solar','/estore','/home-energy','/lancaster-neighbors','/tonawanda-dispatch',
        '/hamburg-service','/buffalo-service','/buffalo-home-improvement',
        '/case_study_category','why-choose-us','/walker','/our-community',
        '/renovation-services','/case-stud','/specials','/tag/']

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
    ('buffalo-ac',                        'AC / Cooling',       'AC Service'),
    ('buffalo-repair',                    'AC / Cooling',       'AC Repair'),
    ('air-conditioning',                  'AC / Cooling',       'Air Conditioning'),
    ('heat-pump-repair',                  'Heat Pump',          'Heat Pump Repair'),
    ('heat-pump-install',                 'Heat Pump',          'Heat Pump Installation'),
    ('heat-pump-maintenance',             'Heat Pump',          'Heat Pump Maintenance'),
    ('heat-pump-services',                'Heat Pump',          'Heat Pump Service'),
    ('heat-pump',                         'Heat Pump',          'Heat Pump Service'),
    ('mitsubishi',                        'Heat Pump',          'Mitsubishi Heat Pump'),
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
    ('trane',                             'HVAC General',       'Trane Systems'),
    ('commercial-hvac',                   'Commercial',         'Commercial HVAC'),
    ('buffalo-commercial',                'Commercial',         'Commercial Service'),
    ('maintenance-plans',                 'HVAC General',       'Maintenance Plans'),
    ('buffalo-ny',                        'Plumbing',           'General Service'),
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
    ('cummins',              'Generator',          'Generator Guide'),
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
]

# Strong transactional signals — override Semrush
TRANS_STRONG = [
    r'\bnear me\b', r'\bin buffalo\b', r'\bin western ny\b',
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
    return 'Ambiguous'

# ── HELPERS ───────────────────────────────────────────────────────────────
def expand(t):
    t = t.lower()
    for p, r in EXPAND: t = re.sub(p, r, t)
    return t

def is_excl(url):
    return any(p in str(url).lower() for p in EXCL)

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

def classify_topic(kw):
    ke = expand(kw.lower())
    TMAP = {
        'Furnace / Heating': ['furnace heating','furnace','boiler heating','boiler','radiant heat','steam heat','forced air'],
        'AC / Cooling':      ['air conditioning','cooling','central air','mini split ductless','ductless','evaporator coil','refrigerant'],
        'Heat Pump':         ['heat pump heating cooling'],
        'Water Heater':      ['water heater hot water','water heater','tankless water heater','hot water'],
        'Drain / Sewer':     ['drain sewer','sewer','clog','unclog','jetting','hydrojet','rooter','sewage','septic'],
        'Plumbing':          ['plumb','pipe','leak','repiping','water line','water main','burst pipe'],
        'Electrical':        ['electric','electrician','panel','wiring','outlet','gfci outlet electrical','circuit','lighting','surge'],
        'Generator':         ['generator backup power'],
        'Sump Pump':         ['sump pump basement'],
        'Air Quality':       ['air quality','dehumidif','humidif','air purif','uv air sanitizer','indoor air'],
        'Water Treatment':   ['soft water','hard water','water treatment','water softener'],
        'Backflow':          ['backflow prevention'],
        'Geothermal':        ['geothermal'],
        'Bathroom / Kitchen':['bathroom','kitchen','toilet','shower','faucet','sink','tub','garbage disposal'],
        'Ventilation':       ['ventilation','hrv','erv','air handler'],
        'EV Charger':        ['ev charger','electric vehicle charg','charging station'],
        'Commercial':        ['commercial'],
    }
    for t, terms in TMAP.items():
        if any(x in ke for x in terms): return t
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
    if any(t in p for t in ['ac-tune','ac-filter','ac-install','evaporator','buffalo-repair','68-ac']): c.add('ac_cooling')
    if p.rstrip('/') == '/air-conditioning': c.update({'ac_cooling','hvac_general'})
    if any(t in p for t in ['ductless','mini-split']): c.update({'ductless','ac_cooling','heat_pump'})
    if any(t in p for t in ['commercial-hvac','duct-cleaning','air-duct','air-handler','trane']): c.add('hvac_general')
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
    if p.rstrip('/') in ['/buffalo-ny','/maintenance-plans','/buffalo']: c.update({'plumbing_general','hvac_general','electrical'})
    for loc in ['cheektowaga','amherst','hamburg','lancaster','west-seneca','orchard-park',
                'east-aurora','springville','alden','akron','batavia','clarence','depew',
                'elma','eden','marilla','holland','wales','darien','medina','attica',
                'boston-ny','grand-island','bennington','west-valley','williamsville','tonawanda']:
        if loc in p: c.add('location'); c.add('loc:'+loc.replace('-',' ')); break
    if '/blog/' in p: c.add('blog')
    for brand in ['trane','mitsubishi','cummins','rheem','bradford','lennox']:
        if brand in p: c.add('brand:'+brand)
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
    for brand in ['trane','mitsubishi','cummins','rheem','bradford','lennox']:
        if brand in kl: c.add('brand:'+brand)
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

def claude_intent(api_key, ambiguous_kws, biz_desc, status_text, progress_bar, p0, p1):
    """Dedicated intent classification — one task only, no mixing."""
    batches = [ambiguous_kws[i:i+100] for i in range(0, len(ambiguous_kws), 100)]
    def make_prompt(batch):
        return f"""You are classifying search intent for a plumbing and HVAC company.

INFORMATIONAL = searcher wants to learn, diagnose, research, or understand something.
They are NOT ready to hire yet. Examples: "ac blows warm air", "furnace smell",
"water heater not hot enough", "how long does a furnace last", "what is radiant heat"

TRANSACTIONAL = searcher wants to hire a company or book a service.
They ARE ready to take action. Examples: "furnace repair service", "plumber",
"AC installation", "emergency drain cleaning"

Business: {biz_desc}

Return ONLY JSON: {{"keyword": "Informational"}} or {{"keyword": "Transactional"}}
No other values allowed.

Keywords:
{json.dumps(batch)}"""
    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=lambda r, b: r if isinstance(r, dict) else {k: 'Transactional' for k in b},
        fallback_fn=lambda b: {k: 'Transactional' for k in b},
        status_prefix="Claude — intent classification",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=1200, timeout=45
    )

def claude_theme_subtheme(api_key, kw_url_pairs, biz_desc, status_text, progress_bar, p0, p1):
    """Assign Theme + Sub-theme for keywords on generic URLs."""
    batches = [kw_url_pairs[i:i+60] for i in range(0, len(kw_url_pairs), 60)]
    def make_prompt(batch):
        items = json.dumps([{"keyword": k, "url": u} for k, u in batch])
        return f"""You are an SEO strategist. For each keyword + URL pair, assign:
1. theme: top-level service category
2. subtheme: specific content topic (what ONE page should cover)

Use these themes only: Furnace / Heating, AC / Cooling, Heat Pump, Water Heater,
Drain / Sewer, Plumbing, Electrical, Generator, Sump Pump, Air Quality,
Water Treatment, Backflow, Geothermal, Bathroom / Kitchen, Ventilation,
EV Charger, Commercial, HVAC General

Sub-theme should be specific enough to represent ONE page.
Examples: "Furnace Repair", "AC Tune-up", "Drain Cleaning", "Water Heater Cost Guide"

Business: {biz_desc}
Return ONLY JSON: {{"keyword": {{"theme": "X", "subtheme": "Y"}},...}}
Pairs: {items}"""
    def parse_result(r, batch):
        out = {}
        for k, u in batch:
            if k in r and isinstance(r[k], dict):
                out[k] = (r[k].get('theme', ''), r[k].get('subtheme', ''))
            else:
                out[k] = ('', '')
        return out
    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=parse_result,
        fallback_fn=lambda b: {k: ('', '') for k, u in b},
        status_prefix="Claude — theme/sub-theme assignment",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=1500, timeout=50
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
        parse_result=lambda r, b: r if isinstance(r, dict) else {k: 'BORDERLINE' for k in b},
        fallback_fn=lambda b: {k: 'BORDERLINE' for k in b},
        status_prefix="Claude — business relevance",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=1200, timeout=45
    )

def claude_validate_blogs(api_key, kw_blog_pairs, status_text, progress_bar, p0, p1):
    batches = [kw_blog_pairs[i:i+50] for i in range(0, len(kw_blog_pairs), 50)]
    def make_prompt(batch):
        return f"""SEO analyst. Does the blog post clearly match the keyword intent?
YES = blog directly answers what the searcher wants to know.
NO = blog topic is different from keyword intent.
Return ONLY JSON: {{"keyword":"YES/NO",...}}
Pairs: {json.dumps([{"keyword":k,"blog":s} for k,s in batch])}"""
    def parse_result(r, batch):
        return {k: r.get(k, 'YES') for k, s in batch}
    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=parse_result,
        fallback_fn=lambda b: {k: 'YES' for k, s in b},
        status_prefix="Claude — blog validation",
        status_text=status_text, progress_bar=progress_bar, p0=p0, p1=p1,
        max_tokens=800, timeout=40
    )

def claude_match_blogs(api_key, keywords, existing_blogs, status_text, progress_bar, p0, p1):
    batches = [keywords[i:i+80] for i in range(0, len(keywords), 80)]
    blog_list = json.dumps([b.split('/blog/')[-1].strip('/') for b in existing_blogs[:60]])
    def make_prompt(batch):
        return f"""SEO analyst. Match each keyword to the best existing blog post slug, or null.
Only return a slug if the blog CLEARLY covers that keyword (90%+ relevance).
Available blog slugs: {blog_list}
Return ONLY JSON: {{"keyword":"slug_or_null",...}}
Keywords: {json.dumps(batch)}"""
    return run_batches(
        api_key, batches,
        make_prompt=make_prompt,
        parse_result=lambda r, b: {k: r.get(k) for k in b},
        fallback_fn=lambda b: {k: None for k in b},
        status_prefix="Claude — blog semantic matching",
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

# ── PHASE 1: GSC VALIDATION ───────────────────────────────────────────────
def phase1_gsc(gsc_df, sf_df, weights, status_text, progress_bar):
    status_text.text("Phase 1: Building URL index...")
    progress_bar.progress(3)
    stop = set()
    if len(sf_df):
        m = re.search(r'https?://([^/]+)', str(get_url(sf_df.iloc[0])))
        if m: stop = set(m.group(1).replace('www.', '').split('.'))
    url_df = sf_df[~sf_df.apply(lambda r: is_excl(get_url(r)), axis=1)].reset_index(drop=True)
    url_df['content'] = url_df.apply(lambda r: build_content(r, weights, stop), axis=1)
    url_df['cats']    = url_df.apply(lambda r: cat_url(get_url(r)), axis=1)
    url_df['ptype']   = url_df.apply(lambda r: 'blog' if '/blog/' in get_url(r).lower() else 'service', axis=1)
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
    return pd.DataFrame(validated), url_df, dedup, stop

# ── PHASE 2: KEYWORD MAPPING ──────────────────────────────────────────────
def phase2_map(sem_df, url_df, gsc_dedup, gsc_val_df, weights, threshold,
               your_col, intent_map, status_text, progress_bar):
    status_text.text("Phase 2: Building TF-IDF models...")
    progress_bar.progress(24)
    svc_df  = url_df[url_df['ptype'] == 'service'].reset_index(drop=True)
    blog_df = url_df[url_df['ptype'] == 'blog'].reset_index(drop=True)
    hp_urls = set()
    for _, r in url_df.iterrows():
        u = get_url(r)
        slug = re.sub(r'https?://[^/]+', '', u).strip('/')
        if slug in ('', 'buffalo', 'buffalo-ny', 'buffalo-ny/'): hp_urls.add(u.lower().strip())
    gsc_status = dict(zip(gsc_val_df['Query'].str.lower().str.strip(), gsc_val_df['Mapping Status']))
    keywords = sem_df['Keyword'].fillna('').tolist()
    volumes  = sem_df['Volume'].fillna(0).tolist()
    your_pos = sem_df[your_col].fillna('').tolist() if your_col in sem_df.columns else [''] * len(keywords)
    cleaned  = [clean_kw(k) for k in keywords]
    def build_sims(df):
        v = TfidfVectorizer(ngram_range=(1,3), min_df=1, sublinear_tf=True)
        v.fit(df['content'].tolist() + cleaned)
        return cosine_similarity(v.transform(cleaned), v.transform(df['content']))
    ss = build_sims(svc_df)
    bs = build_sims(blog_df)
    progress_bar.progress(40)
    status_text.text("Phase 2: Mapping keywords to pages...")
    results = []
    for i, kw in enumerate(keywords):
        if i % 500 == 0:
            progress_bar.progress(min(40 + int((i / len(keywords)) * 25), 64))
            status_text.text(f"Phase 2: Mapping {i:,}/{len(keywords):,} keywords...")
        kl = kw.lower().strip()
        intent = intent_map.get(kl, intent_map.get(kw.lower(), 'Transactional'))
        kc = cat_kw(kw)
        chosen = ''; source = ''; cs = 0.0; gs = 0.0
        if intent == 'Informational':
            for idx in np.argsort(bs[i])[::-1]:
                s = float(bs[i][idx])
                if s < threshold: break
                url = get_url(blog_df.iloc[idx])
                uc  = blog_df.iloc[idx]['cats']
                if is_compat(kc, uc, kw, url):
                    chosen = url; cs = round(s, 4); source = 'Content match (blog)'; break
        else:
            hp_fb = None
            for idx in np.argsort(ss[i])[::-1]:
                s = float(ss[i][idx])
                if s < threshold: break
                url = get_url(svc_df.iloc[idx])
                uc  = svc_df.iloc[idx]['cats']
                if is_compat(kc, uc, kw, url):
                    if url.lower().strip() in hp_urls:
                        if hp_fb is None: hp_fb = (url, round(s, 4))
                        continue
                    chosen = url; cs = round(s, 4); source = 'Content match'; break
            if not chosen and hp_fb:
                chosen, cs = hp_fb; source = 'Content match (homepage fallback)'
        gd  = gsc_dedup.get(kl, {})
        gu  = gd.get('url', '')
        guc = cat_url(gu) if gu else frozenset()
        gst = gsc_status.get(kl, '')
        gv  = gu and not is_excl(gu) and is_compat(kc, guc, kw, gu)
        if gv:
            if intent == 'Informational' and '/blog/' not in gu.lower(): gv = False
            if gst == 'Suspicious' and not chosen: gv = False
        if gv and chosen and gu.lower().strip() == chosen.lower().strip():
            gs = 0.40; source = source.replace('Content match', 'Content + GSC confirmed')
        elif gv and not chosen:
            if gst in ('Confirmed', 'Plausible'):
                chosen = gu; gs = 0.40; source = 'GSC fallback'
        fs = round(cs * 0.7 + gs * 0.3, 4)
        rp = str(your_pos[i]).strip()
        try: pn = float(rp)
        except: pn = None
        if pn and pn > 0:
            if pn <= 10:    rs = 'Ranking p1-10'
            elif pn <= 20:  rs = 'Quick win p11-20'
            elif pn <= 50:  rs = 'Weak ranking p21-50'
            else:           rs = 'Very weak p51-100'
        else: rs = 'Not ranking'
        results.append({'Keyword': kw, 'Volume': int(volumes[i]) if volumes[i] else 0,
                        'Your Position': rp or 'N/A', 'Landing Page': chosen,
                        'Intent': intent, 'Content Score': cs, 'GSC Score': gs,
                        'Final Score': fs, 'Match Source': source, 'Ranking Status': rs,
                        '_cats': kc})
    progress_bar.progress(65)
    return results

# ── PHASE 3: INTENT CLASSIFICATION ───────────────────────────────────────
def phase3_intent(keywords, api_key, biz_desc, status_text, progress_bar):
    status_text.text("Phase 3: Classifying keyword intent...")
    progress_bar.progress(66)
    intent_map = {}
    ambiguous = []
    for kw in keywords:
        result = classify_intent_rules(kw)
        if result == 'Ambiguous':
            ambiguous.append(kw)
        else:
            intent_map[kw.lower()] = result
    progress_bar.progress(68)
    status_text.text(f"Phase 3: Rule-based classified {len(intent_map):,} keywords. Sending {len(ambiguous):,} ambiguous to Claude...")
    if ambiguous:
        claude_intents = claude_intent(api_key, ambiguous, biz_desc, status_text, progress_bar, 68, 75)
        for kw, intent in claude_intents.items():
            intent_map[kw.lower()] = intent if intent in ('Informational','Transactional') else 'Transactional'
    progress_bar.progress(75)
    return intent_map

# ── PHASE 4: BUSINESS RELEVANCE + BLOG FIXES ──────────────────────────────
def phase4_relevance_blogs(mapped, url_df, api_key, biz_desc, excl_str, status_text, progress_bar):
    progress_bar.progress(76)
    # Fix B: validate weak blog matches
    weak_blog = [(r['Keyword'], r['Landing Page'].split('/blog/')[-1].strip('/'))
                 for r in mapped
                 if r['Landing Page'] and '/blog/' in r['Landing Page']
                 and r['Content Score'] < 0.15 and r['Match Source'] != 'GSC fallback']
    blog_valid = {}
    if weak_blog:
        blog_valid = claude_validate_blogs(api_key, weak_blog, status_text, progress_bar, 77, 81)
    # Fix C: semantic blog matching for unmapped informational
    existing_blogs = [get_url(r) for _, r in url_df.iterrows() if '/blog/' in get_url(r).lower()]
    unmapped_info  = [r['Keyword'] for r in mapped if r['Intent'] == 'Informational' and not r['Landing Page']]
    blog_semantic  = {}
    if unmapped_info and existing_blogs:
        top_info = sorted(unmapped_info, key=lambda k: next((r['Volume'] for r in mapped if r['Keyword'] == k), 0), reverse=True)[:400]
        blog_semantic = claude_match_blogs(api_key, top_info, existing_blogs, status_text, progress_bar, 81, 85)
    # Business relevance
    low_scored = [r['Keyword'] for r in mapped if not r['Landing Page'] or r['Final Score'] < 0.15]
    confirmed  = {r['Keyword']: 'RELEVANT' for r in mapped if r['Landing Page'] and r['Final Score'] >= 0.15}
    rel_map = {}
    if low_scored:
        rel_map = claude_relevance(api_key, low_scored, biz_desc, excl_str, status_text, progress_bar, 85, 92)
    all_rel = {**confirmed, **rel_map}
    # Apply blog fixes
    for r in mapped:
        kw = r['Keyword']
        if (r['Landing Page'] and '/blog/' in r['Landing Page']
                and r['Content Score'] < 0.15
                and blog_valid.get(kw, 'YES') == 'NO'):
            r['Landing Page'] = ''; r['Match Source'] = ''; r['Content Score'] = 0.0; r['Final Score'] = 0.0
        if r['Intent'] == 'Informational' and not r['Landing Page']:
            slug = blog_semantic.get(kw)
            if slug:
                full_url = next((u for u in existing_blogs if slug in u), None)
                if full_url:
                    r['Landing Page'] = full_url; r['Match Source'] = 'Claude semantic (blog)'; r['Final Score'] = 0.25
    progress_bar.progress(92)
    return all_rel, mapped

# ── PHASE 5: THEME + SUB-THEME ASSIGNMENT ─────────────────────────────────
def phase5_themes(mapped, api_key, biz_desc, status_text, progress_bar):
    status_text.text("Phase 5: Assigning Theme and Sub-theme...")
    progress_bar.progress(93)
    # Build URL → (theme, subtheme) cache
    url_cache = {}
    generic_url_kw_pairs = []  # needs Claude
    for r in mapped:
        url = r['Landing Page']
        if not url: continue
        if url in url_cache: continue
        theme, sub = url_to_theme_subtheme(url)
        if theme:
            url_cache[url] = (theme, sub)
        else:
            url_cache[url] = None  # needs Claude
    # Collect keyword+URL pairs for generic URLs
    generic_kws_seen = set()
    for r in mapped:
        url = r['Landing Page']
        if url and url_cache.get(url) is None:
            k = r['Keyword']
            if k not in generic_kws_seen:
                generic_kw_url_pairs = []
                generic_kws_seen.add(k)
    # Get unique generic URL+keyword pairs (sample top 5 per generic URL)
    generic_url_samples = {}
    for r in mapped:
        url = r['Landing Page']
        if url and url_cache.get(url) is None:
            if url not in generic_url_samples:
                generic_url_samples[url] = []
            if len(generic_url_samples[url]) < 5:
                generic_url_samples[url].append(r['Keyword'])
    # Build pairs for Claude
    pairs_for_claude = []
    for url, kws in generic_url_samples.items():
        for kw in kws:
            pairs_for_claude.append((kw, url.replace('https://cellinoplumbing.com','').replace('https://','')[:60]))
    theme_map = {}
    if pairs_for_claude:
        raw = claude_theme_subtheme(api_key, pairs_for_claude, biz_desc, status_text, progress_bar, 93, 96)
        theme_map = raw  # keyword → (theme, subtheme)
        # Propagate URL-level theme to all keywords on same URL
        for url in generic_url_samples:
            representative = generic_url_samples[url][0]
            if representative in theme_map:
                url_cache[url] = theme_map[representative]
    # Assign theme + subtheme to every keyword
    for r in mapped:
        url = r['Landing Page']
        kw  = r['Keyword']
        if url:
            ts = url_cache.get(url)
            if ts:
                r['Theme'], r['Sub-theme'] = ts
            elif kw in theme_map:
                r['Theme'], r['Sub-theme'] = theme_map[kw]
            else:
                r['Theme'] = classify_topic(kw); r['Sub-theme'] = ''
        else:
            r['Theme'] = classify_topic(kw); r['Sub-theme'] = ''
    progress_bar.progress(96)
    return mapped

# ── PHASE 6: SEMANTIC CLUSTERING ──────────────────────────────────────────
def phase6_cluster(mapped, rel_map, api_key, status_text, progress_bar):
    status_text.text("Phase 6: Semantic clustering...")
    progress_bar.progress(97)
    unmapped_info  = sorted([(r['Keyword'], r['Volume']) for r in mapped
                              if not r['Landing Page'] and rel_map.get(r['Keyword'], '') in ('RELEVANT','BORDERLINE')
                              and r['Intent'] == 'Informational'], key=lambda x: -x[1])[:800]
    unmapped_trans = sorted([(r['Keyword'], r['Volume']) for r in mapped
                              if not r['Landing Page'] and rel_map.get(r['Keyword'], '') in ('RELEVANT','BORDERLINE')
                              and r['Intent'] == 'Transactional'], key=lambda x: -x[1])[:1200]
    clusters = []
    if unmapped_info:
        clusters += claude_cluster(api_key, unmapped_info, "blog post", status_text, progress_bar, 97, 98)
    if unmapped_trans:
        clusters += claude_cluster(api_key, unmapped_trans, "service page", status_text, progress_bar, 98, 99)
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

# ── EXCEL OUTPUT ──────────────────────────────────────────────────────────
def build_excel(gsc_df, mapped, rel_map, clusters, url_clusters):
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
    hdr(ws2, 1, ['Theme','Sub-theme','Keyword','Volume','Your Position','Landing Page','Intent',
                 'Intent Source','Content Score','GSC Score','Final Score','Match Source','Ranking Status'])
    for i, r in enumerate(sorted(mapped, key=lambda x: (x.get('Theme',''), x.get('Sub-theme',''), -x.get('Final Score',0)))):
        row = i+2; url = r['Landing Page']; src = r['Match Source']; rel = rel_map.get(r['Keyword'], '')
        if url and '/blog/' in url:           fill = make_fill(C['BL'])
        elif src == 'GSC fallback':            fill = make_fill(C['YL'])
        elif src and 'Claude semantic' in src: fill = make_fill(C['OR'])
        elif url:                              fill = make_fill(C['GR'])
        elif rel in ('RELEVANT','BORDERLINE'): fill = make_fill(C['PU'])
        else:                                  fill = make_fill(C['WH'])
        intent_src = 'Claude' if r.get('_intent_source','') == 'claude' else 'Rule-based'
        vals = [r.get('Theme',''), r.get('Sub-theme',''), r['Keyword'], r['Volume'],
                r['Your Position'], url or '', r['Intent'], intent_src,
                r['Content Score'], r['GSC Score'], r['Final Score'], src or '', r['Ranking Status']]
        for col, v in enumerate(vals, 1):
            c = ws2.cell(row=row, column=col, value=v); c.fill = fill; c.font = lf if col==6 else bf
    cw(ws2, [22,28,48,12,14,62,15,14,14,12,12,26,18]); ws2.freeze_panes = 'A2'

    # TAB 3: Opportunity Classification (with Theme + Sub-theme)
    ws3 = wb.create_sheet('Opportunity Classification')
    hdr(ws3, 1, ['Theme','Sub-theme','Keyword','Volume','Your Position','Landing Page',
                 'Intent','Final Score','Opportunity Type','Action'])
    OPP_F = {'Confirmed existing page': make_fill(C['GR']), 'Quick win — optimise': make_fill(C['DGR']),
             'Weak ranking': make_fill(C['YL']), 'Page exists — optimise': make_fill(C['BL']),
             'Blog exists — optimise': make_fill(C['OR']), 'Business relevant gap': make_fill(C['PU']),
             'True content gap': make_fill(C['RD'])}
    for r in sorted(mapped, key=lambda x: (x.get('Theme',''), x.get('Sub-theme',''), -x.get('Volume',0))):
        rel = rel_map.get(r['Keyword'], ''); url = r['Landing Page']
        rs  = r['Ranking Status']; fs = r['Final Score']; src = r.get('Match Source', '')
        if url and rs == 'Ranking p1-10':         opp = 'Confirmed existing page'; act = 'Monitor — already ranking well'
        elif rs == 'Quick win p11-20':             opp = 'Quick win — optimise';    act = 'Optimise title, meta, H1 — almost ranking'
        elif rs in ['Weak ranking p21-50','Very weak p51-100']: opp = 'Weak ranking'; act = 'Improve page content + internal links'
        elif url and 'Claude semantic' in src:     opp = 'Blog exists — optimise';  act = 'Update blog title/meta/H1 for this keyword'
        elif url and '/blog/' in url:              opp = 'Page exists — optimise';  act = 'Optimise existing blog post for this keyword'
        elif url:                                  opp = 'Page exists — optimise';  act = 'Optimise existing service page for this keyword'
        elif rel in ('RELEVANT','BORDERLINE'):
            opp = 'Business relevant gap'
            act = 'Create new blog post' if r['Intent']=='Informational' else 'Create new service page'
        else:                                      opp = 'True content gap'; act = 'Evaluate — may need new page'
        fill = OPP_F.get(opp, make_fill(C['WH']))
        rn = ws3.max_row + 1
        for col, v in enumerate([r.get('Theme',''), r.get('Sub-theme',''), r['Keyword'], r['Volume'],
                                  r['Your Position'], url or '', r['Intent'], fs, opp, act], 1):
            c = ws3.cell(row=rn, column=col, value=v); c.fill = fill; c.font = lf if col==6 else bf
    cw(ws3, [22,28,48,12,14,62,15,12,28,48]); ws3.freeze_panes = 'A2'

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

    # TAB 5: Business Relevant Gaps (with Theme + Sub-theme)
    ws5 = wb.create_sheet('Business Relevant Gaps')
    hdr(ws5, 1, ['Theme','Sub-theme','Keyword','Volume','Intent','Relevance','Action Needed'])
    gaps = sorted([r for r in mapped if not r['Landing Page']
                   and rel_map.get(r['Keyword'], '') in ('RELEVANT','BORDERLINE')],
                  key=lambda x: (x.get('Theme',''), x.get('Sub-theme',''), -x.get('Volume',0)))
    for r in gaps:
        rel = rel_map.get(r['Keyword'], '')
        fill = make_fill(C['PU'] if rel == 'RELEVANT' else C['AM'])
        act = 'Create new blog post' if r['Intent']=='Informational' else 'Create new service page'
        rn = ws5.max_row + 1
        for col, v in enumerate([r.get('Theme',''), r.get('Sub-theme',''), r['Keyword'],
                                  r['Volume'], r['Intent'], rel, act], 1):
            c = ws5.cell(row=rn, column=col, value=v); c.fill = fill; c.font = bf
    cw(ws5, [22,28,50,12,15,14,28]); ws5.freeze_panes = 'A2'

    # TAB 6: Priority Roadmap (with Theme + Sub-theme)
    ws6 = wb.create_sheet('Priority Roadmap')
    ws6['A1'] = 'SEO Content Opportunity Roadmap'
    ws6['A1'].font = Font(bold=True, size=14, color='0F6E56'); ws6.merge_cells('A1:J1')
    qw   = sum(1 for r in mapped if r['Ranking Status'] == 'Quick win p11-20')
    weak = sum(1 for r in mapped if r['Ranking Status'] in ['Weak ranking p21-50','Very weak p51-100'])
    pnr  = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status'] == 'Not ranking')
    bgap = len(gaps)
    conf = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status'] == 'Ranking p1-10')
    hdr(ws6, 3, ['Opportunity Type','Count','Total Volume','Action','Priority',''])
    summary = [
        ('Quick Wins (p11-20)', qw, sum(r['Volume'] for r in mapped if r['Ranking Status']=='Quick win p11-20'), 'Optimise existing pages', 'High', C['DGR']),
        ('Weak Rankings (p21-100)', weak, sum(r['Volume'] for r in mapped if r['Ranking Status'] in['Weak ranking p21-50','Very weak p51-100']), 'Improve content + internal links', 'High', C['YL']),
        ('Pages Exist — Not Ranking', pnr, sum(r['Volume'] for r in mapped if r['Landing Page'] and r['Ranking Status']=='Not ranking'), 'Optimise existing pages', 'Medium', C['BL']),
        ('Business Relevant Gaps', bgap, sum(r['Volume'] for r in gaps), 'Create new pages / blog posts', 'Medium', C['PU']),
        ('Already Ranking Well', conf, sum(r['Volume'] for r in mapped if r['Landing Page'] and r['Ranking Status']=='Ranking p1-10'), 'Monitor only', 'Low', C['GY']),
    ]
    for i, (opp,cnt,vol,act,pri,color) in enumerate(summary):
        r = i+4; fill = make_fill(color)
        for col, v in enumerate([opp,cnt,vol,act,pri,''], 1):
            c = ws6.cell(row=r, column=col, value=v); c.fill = fill; c.font = Font(size=10, bold=(col==5))
    dr = len(summary) + 7
    ws6.cell(row=dr, column=1, value='Detailed Action List — sorted by Theme > Sub-theme > Priority').font = Font(bold=True, size=12, color='0F6E56')
    ws6.merge_cells(start_row=dr, start_column=1, end_row=dr, end_column=10); dr += 1
    hdr(ws6, dr, ['Theme','Sub-theme','Keyword','Volume','Your Position','Intent',
                  'Final Score','Opportunity','Action','Priority']); dr += 1
    PORD = {'Quick win — optimise':1,'Weak ranking':2,'Page exists — optimise':3,
            'Blog exists — optimise':3,'Business relevant gap':4,'True content gap':5,
            'Confirmed existing page':6}
    AF = {'Quick win — optimise':make_fill(C['DGR']),'Weak ranking':make_fill(C['YL']),
          'Page exists — optimise':make_fill(C['BL']),'Blog exists — optimise':make_fill(C['OR']),
          'Business relevant gap':make_fill(C['PU']),'True content gap':make_fill(C['RD']),
          'Confirmed existing page':make_fill(C['GY'])}
    all_items = []
    for r in mapped:
        rel = rel_map.get(r['Keyword'],''); url = r['Landing Page']
        rs  = r['Ranking Status']; fs = r['Final Score']; src = r.get('Match Source','')
        if url and rs=='Ranking p1-10':           opp='Confirmed existing page'
        elif rs=='Quick win p11-20':              opp='Quick win — optimise'
        elif rs in['Weak ranking p21-50','Very weak p51-100']: opp='Weak ranking'
        elif url and 'Claude semantic' in src:    opp='Blog exists — optimise'
        elif url:                                 opp='Page exists — optimise'
        elif rel in('RELEVANT','BORDERLINE'):     opp='Business relevant gap'
        else:                                     opp='True content gap'
        act = {'Confirmed existing page':'Monitor','Quick win — optimise':'Optimise title + meta + H1',
               'Weak ranking':'Improve content + internal links',
               'Page exists — optimise':'Optimise existing page',
               'Blog exists — optimise':'Update blog title/meta/H1',
               'Business relevant gap':'Create new blog post' if r['Intent']=='Informational' else 'Create new service page',
               'True content gap':'Evaluate for new page'}.get(opp,'')
        pri = {1:'High',2:'High',3:'Medium',4:'Medium',5:'Low',6:'Low'}.get(PORD.get(opp,5),'Low')
        all_items.append({**r,'opp':opp,'act':act,'pri':pri,'po':PORD.get(opp,5)})
    # Sort by Theme → Sub-theme → Priority → Volume
    all_items.sort(key=lambda x: (x.get('Theme','zzz'), x.get('Sub-theme','zzz'), x['po'], -x['Volume']))
    for item in all_items:
        fill = AF.get(item['opp'], make_fill(C['WH']))
        for col, v in enumerate([item.get('Theme',''), item.get('Sub-theme',''), item['Keyword'],
                                  item['Volume'], item['Your Position'], item['Intent'],
                                  item['Final Score'], item['opp'], item['act'], item['pri']], 1):
            c = ws6.cell(row=dr, column=col, value=v); c.fill = fill; c.font = Font(size=10, bold=(col==10))
        dr += 1
    cw(ws6, [22,28,48,12,14,15,12,28,42,10]); ws6.freeze_panes = f'A{len(summary)+10}'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── UI ────────────────────────────────────────────────────────────────────
st.title("🔍 SEO Content Mapper")
st.markdown("Complete SEO gap analysis — GSC validation + semantic intent + accurate theme clustering.")

with st.sidebar:
    st.header("⚙️ Settings"); st.markdown("---")
    threshold = st.slider("Match Score Threshold", 0.10, 0.40, 0.15, 0.01)
    st.markdown("**Content Weights**")
    sw=st.slider("URL Slug",1,8,5); tw=st.slider("Page Title",1,6,3)
    hw=st.slider("H1 Heading",1,4,2); mw=st.slider("Meta Description",1,3,1)
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
                                    help="Column showing YOUR site's position e.g. cellinoplumbing.com", key='sdc')
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
        # Phase 1
        gsc_val, url_df, gsc_dedup, stop = phase1_gsc(gsc_df, sf_df, weights, status_text, progress_bar)

        # Phase 3: Intent classification (before mapping so mapping uses correct intent)
        all_kws = sem_df['Keyword'].fillna('').tolist()
        intent_map = phase3_intent(all_kws, api_key, biz_desc, status_text, progress_bar)

        # Phase 2
        mapped = phase2_map(sem_df, url_df, gsc_dedup, gsc_val, weights, threshold,
                            your_col, intent_map, status_text, progress_bar)

        # Phase 4
        rel_map, mapped = phase4_relevance_blogs(mapped, url_df, api_key, biz_desc, excl_str,
                                                  status_text, progress_bar)

        # Phase 5: Theme + Sub-theme
        mapped = phase5_themes(mapped, api_key, biz_desc, status_text, progress_bar)

        # Phase 6: Clustering
        clusters, url_clusters = phase6_cluster(mapped, rel_map, api_key, status_text, progress_bar)

        progress_bar.progress(99)
        status_text.text("Building Excel output...")
        excel_buf = build_excel(gsc_val, mapped, rel_map, clusters, url_clusters)
        progress_bar.progress(100)
        elapsed = round(time.time() - t_start)
        status_text.text(f"✅ Done in {elapsed//60}m {elapsed%60}s")

        # Metrics
        st.markdown("---"); st.subheader("📊 Analysis Complete")
        mapped_kw = sum(1 for r in mapped if r['Landing Page'])
        qw        = sum(1 for r in mapped if r['Ranking Status'] == 'Quick win p11-20')
        bgaps     = sum(1 for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in ('RELEVANT','BORDERLINE'))
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
                                  'Relevance':rel_map.get(r['Keyword'],'')} for r in mapped
                                 if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in ('RELEVANT','BORDERLINE')]).sort_values(['Theme','Volume'], ascending=[True,False])
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
