import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SEO Content Mapper",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main { padding-top: 1rem; }
    .stProgress > div > div > div { background-color: #1D9E75; }
    .metric-card {
        background: #f0faf5;
        border: 1px solid #1D9E75;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .metric-num { font-size: 2rem; font-weight: 700; color: #1D9E75; }
    .metric-label { font-size: 0.85rem; color: #555; margin-top: 4px; }
    .info-box {
        background: #e8f5f0;
        border-left: 4px solid #1D9E75;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 8px 0;
    }
    .warn-box {
        background: #fffbea;
        border-left: 4px solid #f59e0b;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 8px 0;
    }
    .error-box {
        background: #fef2f2;
        border-left: 4px solid #ef4444;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 8px 0;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAXONOMY & RULES (all logic from the full workflow)
# ─────────────────────────────────────────────────────────────────────────────

HVAC_GROUP       = {'furnace','boiler','heat_pump','ac_cooling','ductless','hvac_general','geothermal'}
PLUMBING_GROUP   = {'water_pipes','drain_sewer','water_heater','toilet','faucet_sink','shower_tub',
                    'bathroom','kitchen','plumbing_general','backflow','sump_pump','sprinkler'}
ELECTRICAL_GROUP = {'electrical','generator','ev_charger'}
ALL_EXCLUSIVE    = (HVAC_GROUP | PLUMBING_GROUP | ELECTRICAL_GROUP | {'gas_utility','air_quality'})

EXCLUDE_PATTERNS = [
    '/truck-wrap','/careers','/join-the-team','/contact','/podcast',
    '/coupons','/warranty','/about-us','/meet-the','/technicians',
    '/community','/financing','/solar','/estore','/home-energy',
    '/lancaster-neighbors','/tonawanda-dispatch','/hamburg-service',
    '/buffalo-service','/buffalo-home-improvement','/case_study_category',
    'why-choose-us','/walker','/our-community','/renovation-services',
    '/case-stud','/specials'
]

EXPAND_RULES = [
    (r'\bac\b','air conditioning'),(r'\ba/c\b','air conditioning'),
    (r'\bhvac\b','heating cooling air conditioning'),
    (r'\bfurnace\b','furnace heating'),(r'\bheat pump\b','heat pump heating cooling'),
    (r'\bboiler\b','boiler heating'),(r'\bgenerator\b','generator backup power'),
    (r'\bdrain\b','drain sewer'),(r'\bsump pump\b','sump pump basement'),
    (r'\bgfci\b','gfci outlet electrical'),(r'\bbackflow\b','backflow prevention'),
    (r'\buv\b','uv air sanitizer'),(r'\brepiping\b','repiping pipe replacement'),
    (r'\bwater heater\b','water heater hot water'),(r'\btankless\b','tankless water heater'),
    (r'\bmini.?split\b','mini split ductless heating cooling'),
]

INFO_SIGNALS = [
    r'^what\b',r'^how\b',r'^why\b',r'^when\b',r'^does\b',r'^do\b',
    r'^is\b',r'^are\b',r'^can\b',r'^should\b',r'^which\b',r'^who\b',
    r'\bvs\b',r'\bversus\b',r'\bdifference between\b',r'\btips\b',
    r'\bbenefits of\b',r'\bsigns\b',r'\bcauses\b',r'\btypes of\b',
    r'\bhow to\b',r'\bwhat is\b',r'\bwhy is\b',r'\bhow does\b',
    r'\bwhat does\b',r'\badvantages\b',r'\bproblems with\b'
]

SLUG_STOP = {'com','cellinoplumbing','buffalo','ny','https','http',
             'in','and','the','a','of','for','to','near','me'}

TOPIC_CLUSTERS = {
    'Furnace / Heating':   ['furnace heating','furnace','heating','heat','boiler heating','boiler','steam heat','radiant'],
    'AC / Cooling':        ['air conditioning','cooling','central air','mini split ductless','ductless'],
    'Heat Pump':           ['heat pump heating cooling'],
    'Water Heater':        ['water heater hot water','water heater','tankless water heater','hot water'],
    'Drain / Sewer':       ['drain sewer','sewer','clog','unclog','jetting','hydrojet','rooter'],
    'Plumbing General':    ['plumb','pipe','leak','repiping','water line','water main'],
    'Electrical':          ['electric','electrician','panel','wiring','outlet','gfci outlet electrical',
                            'circuit','lighting','generator backup power','surge'],
    'HVAC General':        ['heating cooling air conditioning'],
    'Sump Pump':           ['sump pump basement'],
    'Air Quality':         ['air quality','dehumidif','humidif','air purif','uv air sanitizer','indoor air'],
    'Backflow':            ['backflow prevention'],
    'Geothermal':          ['geothermal'],
    'Bathroom / Kitchen':  ['bathroom','kitchen','toilet','shower','faucet','sink','tub','vanity','remodel','renovation'],
    'Commercial':          ['commercial'],
    'Location-Specific':   ['cheektowaga','amherst','hamburg','lancaster','west seneca','orchard park',
                            'east aurora','clarence','depew','tonawanda','williamsville','batavia',
                            'grand island','rochester','syracuse','albany'],
}

# ─────────────────────────────────────────────────────────────────────────────
# CORE FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def expand(t):
    t = t.lower()
    for p, r in EXPAND_RULES:
        t = re.sub(p, r, t)
    return t

def is_excluded(url):
    return any(p in str(url).lower() for p in EXCLUDE_PATTERNS)

def is_informational(kw):
    return any(re.search(p, kw.lower()) for p in INFO_SIGNALS)

def slug_terms(url, custom_stop=None):
    stop = custom_stop or SLUG_STOP
    path  = re.sub(r'https?://[^/]+', '', str(url))
    parts = re.split(r'[/\-_]', path)
    return ' '.join(p for p in parts if p and p.lower() not in stop)

def build_url_content(row, weights):
    slug  = expand(slug_terms(row['Landing Page']))
    title = expand(str(row.get('Page Title','') or '').replace('| Cellino Plumbing','').replace('- Cellino Plumbing',''))
    meta  = expand(str(row.get('Meta Description','') or ''))
    h1r   = str(row.get('H1','') or '')
    h1    = '' if h1r.lower() in ('nan','none','') else expand(h1r)
    slug_w, title_w, h1_w, meta_w = weights
    return f"{slug} " * slug_w + f"{title} " * title_w + f"{h1} " * h1_w + meta * meta_w

def categorize_url(url):
    p    = url.lower().replace('https://','').replace('http://','')
    # strip domain
    p    = re.sub(r'^[^/]+', '', p)
    cats = set()
    if any(t in p for t in ['gas-leak','gas-line','gas-pipe','gas-service','gas-detection']): cats.add('gas_utility')
    if any(t in p for t in ['water-leak','water-line','water-main','main-water','burst-pipe','slab-leak','repiping','frozen-pipe','underground','water-shut']): cats.add('water_pipes')
    if any(t in p for t in ['drain','sewer','hydro','jetting','rooter','bubbler','clogged-kitchen','flooded-basement']): cats.add('drain_sewer')
    if any(t in p for t in ['water-heater','tankless']): cats.add('water_heater')
    if any(t in p for t in ['furnace','68-furnace','heater-install']): cats.add('furnace')
    if p.rstrip('/') in ['/heating','']: cats.update({'furnace','boiler','heat_pump','hvac_general'})
    if any(t in p for t in ['boiler','steam-boiler']): cats.add('boiler')
    if 'heat-pump' in p: cats.add('heat_pump')
    if any(t in p for t in ['68-ac','ac-tune','ac-filter','ac-install','evaporator','buffalo-repair']): cats.add('ac_cooling')
    if p.rstrip('/') in ['/air-conditioning']: cats.update({'ac_cooling','hvac_general'})
    if any(t in p for t in ['ductless','mini-split']): cats.update({'ductless','ac_cooling','heat_pump'})
    if any(t in p for t in ['commercial-hvac','duct-cleaning','air-duct','air-handler','trane']): cats.add('hvac_general')
    if 'geothermal' in p: cats.update({'geothermal','hvac_general'})
    if any(t in p for t in ['electric','electrical','circuit-breaker','surge-protection','energy-efficient-upgrade']): cats.add('electrical')
    if 'generator' in p: cats.update({'generator','electrical'})
    if 'car-home-charging' in p: cats.update({'ev_charger','electrical'})
    if 'sump' in p or 'basepump' in p: cats.add('sump_pump')
    if any(t in p for t in ['air-quality','dehumidif','uv-air','iaq','indoor-air']): cats.add('air_quality')
    if 'backflow' in p: cats.add('backflow')
    if 'toilet' in p: cats.update({'toilet','drain_sewer'})
    if any(t in p for t in ['garbage-disposal','faucet']): cats.update({'faucet_sink','plumbing_general'})
    if any(t in p for t in ['bathroom-remodel','bathroom-renov']): cats.update({'bathroom','plumbing_general'})
    if any(t in p for t in ['kitchen-remodel','kitchen-renov']): cats.update({'kitchen','plumbing_general'})
    if 'sprinkler' in p: cats.add('sprinkler')
    if any(t in p for t in ['residential-plumbing','emergency-plumber','commercial-plumbing','plumbing-inspection']): cats.add('plumbing_general')
    # detect location pages
    for loc in ['cheektowaga','amherst','hamburg','lancaster','west-seneca','orchard-park',
                'east-aurora','springville','alden','akron','batavia','clarence','depew',
                'elma','eden','marilla','holland','wales','darien','medina','elba',
                'attica','boston-ny','grand-island','bennington','west-valley','williamsville','tonawanda']:
        if loc in p:
            cats.add('location')
            cats.add('location:'+loc.replace('-',' '))
            break
    if '/blog/' in p: cats.add('blog')
    if not (cats - {'blog'}): cats.add('plumbing_general')
    return frozenset(cats)

def categorize_kw(kw):
    kw_l = expand(kw.lower())
    cats = set()
    GAS_EXACT = ['gas leak','gas line','gas pipe','natural gas line','gas service line',
                 'gas detector','gas fireplace','gas stove','gas dryer','gas appliance',
                 'gas connection','gas hookup','gas meter','carbon monoxide detector',
                 'gas shut off','propane line','natural gas leak']
    if any(t in kw_l for t in GAS_EXACT):
        if not any(x in kw_l for x in ['furnace heating','boiler heating','water heater hot water','heater']):
            cats.add('gas_utility')
    WH = ['water heater hot water','water heater','hot water heater','tankless water heater',
          'hot water tank','leaking water heater','water heater leak','water heater drip',
          'water heater temp','water heater not','no hot water','water not heating']
    if any(t in kw_l for t in WH): cats.add('water_heater')
    if 'water_heater' not in cats:
        WPIPE = ['water leak','water line','water main','water pipe','burst pipe','pipe leak',
                 'slab leak','repiping','main water','frozen pipe','water service line',
                 'water shut off valve','main line leak','pipe burst','whole home repipe','water pressure']
        if any(t in kw_l for t in WPIPE): cats.add('water_pipes')
    DRAIN = ['drain sewer','sewer','clogged drain sewer','unclog','hydro jet','hydrojett',
             'rooter','sewage','septic','main line clog','main sewer','main drain sewer',
             'backed up drain','slow drain','floor drain','bubbler','toilet drain sewer',
             'storm drain','overflow drain','exterior drain','interior drain',
             'basement drain','laundry drain','bathroom drain','kitchen drain']
    if any(t in kw_l for t in DRAIN): cats.add('drain_sewer')
    FURN = ['furnace heating','furnace tune','furnace service','furnace repair','furnace install',
            'furnace maint','furnace clean','furnace filter','furnace not working',
            'furnace not heating','annual furnace','gas furnace','electric furnace',
            'new furnace','furnace replace','forced air']
    if any(t in kw_l for t in FURN): cats.add('furnace')
    BOIL = ['boiler heating','boiler repair','boiler service','boiler install','boiler maint',
            'boiler replace','boiler tune','boiler cleaning','annual boiler',
            'gas boiler','oil boiler','steam heat','steam boiler','radiant heat']
    if any(t in kw_l for t in BOIL): cats.add('boiler')
    if 'heat pump heating cooling' in kw_l: cats.add('heat_pump')
    AC = ['air conditioning','air conditioner','central air','cooling system',
          'evaporator coil','refrigerant','freon','cooling near me']
    if any(t in kw_l for t in AC): cats.add('ac_cooling')
    if 'mini split ductless' in kw_l: cats.update({'ductless','ac_cooling','heat_pump'})
    if 'heating cooling air conditioning' in kw_l: cats.update({'hvac_general','furnace','ac_cooling'})
    elif 'heating cooling' in kw_l: cats.update({'hvac_general','furnace','ac_cooling'})
    if 'geothermal' in kw_l: cats.update({'geothermal','hvac_general'})
    is_combo = any(x in kw_l for x in ['water heater','furnace heating','boiler heating','heat pump heating'])
    ELEC = ['electrician','electric panel','breaker box','circuit breaker','electrical panel',
            'electrical wire','wiring','rewiring','outlet','gfci outlet electrical',
            'light fixture','lighting install','lighting repair','surge protect',
            'electrical inspect','electrical troubl','electrical service','electrical repair',
            'electrical install','electrical work','electrical contractor','electric repair',
            'electric install','residential electric','outdoor lighting','electrical permit',
            'electrical company','local electrician','licensed electrician','home electrical']
    if any(t in kw_l for t in ELEC) and not is_combo: cats.add('electrical')
    if 'generator backup power' in kw_l: cats.update({'generator','electrical'})
    if any(t in kw_l for t in ['ev charger','electric vehicle charg','charging station','car charging','home charging']):
        cats.update({'ev_charger','electrical'})
    if 'sump pump basement' in kw_l: cats.add('sump_pump')
    AIR_Q = ['air quality','dehumidif','humidif','air purif','air filter service',
             'uv air sanitizer','indoor air','air cleaner','air scrubber']
    if any(t in kw_l for t in AIR_Q): cats.add('air_quality')
    if 'backflow prevention' in kw_l: cats.add('backflow')
    if 'toilet' in kw_l: cats.update({'toilet','drain_sewer'})
    if any(t in kw_l for t in ['faucet','garbage disposal','kitchen sink','disposal unit']):
        cats.update({'faucet_sink','plumbing_general'})
    if any(t in kw_l for t in ['shower','bathtub','bath tub']):
        cats.update({'shower_tub','drain_sewer','plumbing_general'})
    if any(t in kw_l for t in ['bathroom renov','bathroom remodel','bathroom addition','bathroom plumb']):
        cats.update({'bathroom','plumbing_general'})
    if any(t in kw_l for t in ['kitchen renov','kitchen remodel','kitchen plumb']):
        cats.update({'kitchen','plumbing_general'})
    if 'sprinkler' in kw_l: cats.add('sprinkler')
    PLUMB = ['plumbing','plumber','pipe install','pipe replac','pipe repair','pipe clean',
             'residential plumb','commercial plumb','plumbing inspect','emergency plumb',
             '24 hour plumb','plumbing company','plumbing service','local plumb',
             'home plumb','plumb contractor','licensed plumb','best plumb','professional plumb']
    if any(t in kw_l for t in PLUMB): cats.add('plumbing_general')
    if 'heater' in kw_l and 'water heater' not in kw_l and 'water_heater' not in cats:
        cats.update({'furnace','hvac_general'})
    if re.search(r'\bheating\b', kw_l) and not (cats & {'furnace','boiler','heat_pump','hvac_general','water_heater'}):
        cats.update({'furnace','hvac_general'})
    if re.search(r'\bcooling\b', kw_l) and not (cats & {'ac_cooling','ductless','heat_pump','hvac_general'}):
        cats.update({'ac_cooling','hvac_general'})
    return frozenset(cats)

def is_compatible(kw_cats, url_cats, kw, url):
    if 'location' in url_cats:
        kw_l      = expand(kw.lower())
        loc_tags  = [t for t in url_cats if t.startswith('location:')]
        return bool(loc_tags) and loc_tags[0].replace('location:','') in kw_l
    if 'gas_utility' in url_cats and 'gas_utility' not in kw_cats: return False
    if 'gas_utility' in kw_cats and 'gas_utility' not in url_cats: return False
    kw_ex  = kw_cats  & ALL_EXCLUSIVE
    url_ex = url_cats & ALL_EXCLUSIVE
    if not kw_ex or not url_ex: return True
    if kw_ex & url_ex: return True
    kw_specific_hvac   = kw_cats  & (HVAC_GROUP - {'hvac_general'})
    url_specific_hvac  = url_cats & (HVAC_GROUP - {'hvac_general'})
    if 'hvac_general' in kw_cats  and not kw_specific_hvac  and url_ex & HVAC_GROUP: return True
    if 'hvac_general' in url_cats and not url_specific_hvac and kw_ex  & HVAC_GROUP: return True
    kw_specific_plumb  = kw_cats  & (PLUMBING_GROUP - {'plumbing_general'})
    url_specific_plumb = url_cats & (PLUMBING_GROUP - {'plumbing_general'})
    if 'plumbing_general' in kw_cats  and not kw_specific_plumb  and url_ex & PLUMBING_GROUP: return True
    if 'plumbing_general' in url_cats and not url_specific_plumb and kw_ex  & PLUMBING_GROUP: return True
    if 'electrical' in kw_cats  and url_ex & ELECTRICAL_GROUP: return True
    if 'electrical' in url_cats and kw_ex  & ELECTRICAL_GROUP: return True
    return False

def clean_kw(kw):
    kw = expand(kw.lower().strip())
    for p in [r'\bbuffalo,?\s*ny\b', r'\bnear me\b', r'\bwestern ny\b', r'\bnear\b']:
        kw = re.sub(p, '', kw)
    return re.sub(r'\s+', ' ', kw).strip()

def classify_topic(kw):
    kw_e = expand(kw.lower())
    for t, terms in TOPIC_CLUSTERS.items():
        if any(x in kw_e for x in terms): return t
    return 'Other'

# ─────────────────────────────────────────────────────────────────────────────
# MAIN MAPPING FUNCTION
# ─────────────────────────────────────────────────────────────────────────────

def run_mapping(keywords, vol_map, url_df, gsc_map, threshold, weights, progress_bar, status_text):

    url_df = url_df[~url_df['Landing Page'].apply(is_excluded)].reset_index(drop=True)
    url_df['content'] = url_df.apply(lambda r: build_url_content(r, weights), axis=1)
    url_df['cats']    = url_df['Landing Page'].apply(categorize_url)
    url_df['ptype']   = url_df['Landing Page'].apply(lambda u: 'blog' if '/blog/' in u.lower() else 'service')

    service_df = url_df[url_df['ptype']=='service'].reset_index(drop=True)
    blog_df    = url_df[url_df['ptype']=='blog'].reset_index(drop=True)

    cleaned_kw = [clean_kw(k) for k in keywords]

    status_text.text("Building TF-IDF models...")
    progress_bar.progress(10)

    def build_sims(df):
        vec = TfidfVectorizer(ngram_range=(1,3), min_df=1, sublinear_tf=True)
        vec.fit(df['content'].tolist() + cleaned_kw)
        return cosine_similarity(vec.transform(cleaned_kw), vec.transform(df['content']))

    svc_sims  = build_sims(service_df)
    blog_sims = build_sims(blog_df)
    progress_bar.progress(40)
    status_text.text("Matching keywords to pages...")

    GSC_BOOST       = 0.10
    GSC_FALLBACK_MIN = 0.40

    mapped_urls, sources, scores_out, intents = [], [], [], []
    n = len(keywords)

    for i, kw in enumerate(keywords):
        if i % 500 == 0:
            pct = 40 + int((i / n) * 55)
            progress_bar.progress(min(pct, 94))
            status_text.text(f"Matching keywords... {i:,} / {n:,}")

        kw_low  = kw.lower().strip()
        intent  = 'Informational' if is_informational(kw) else 'Transactional'
        kw_cats = categorize_kw(kw)
        chosen  = ''
        source  = ''
        score   = 0.0

        # Step 1: Content match (PRIMARY)
        sims_data = blog_sims[i] if intent == 'Informational' else svc_sims[i]
        pool_df   = blog_df      if intent == 'Informational' else service_df
        ranked    = np.argsort(sims_data)[::-1]

        for idx in ranked:
            s = float(sims_data[idx])
            if s < threshold: break
            url      = pool_df.iloc[idx]['Landing Page']
            url_cats = pool_df.iloc[idx]['cats']
            if is_compatible(kw_cats, url_cats, kw, url):
                chosen = url
                score  = round(s, 4)
                source = 'Content match'
                break

        # Step 2: Check GSC confirmation or fallback
        gsc_url      = gsc_map.get(kw_low, '')
        gsc_url_cats = categorize_url(gsc_url) if gsc_url else frozenset()
        gsc_valid    = (gsc_url
                        and not is_excluded(gsc_url)
                        and is_compatible(kw_cats, gsc_url_cats, kw, gsc_url))

        if chosen and gsc_valid and gsc_url == chosen:
            score  = min(round(score + GSC_BOOST, 4), 0.99)
            source = 'Content + GSC confirmed'

        if not chosen and gsc_valid:
            chosen = gsc_url
            source = 'GSC fallback'
            score  = GSC_FALLBACK_MIN

        mapped_urls.append(chosen)
        sources.append(source)
        scores_out.append(score)
        intents.append(intent)

    progress_bar.progress(95)
    status_text.text("Building output...")
    return mapped_urls, sources, scores_out, intents

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL OUTPUT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(keywords, mapped_urls, sources, scores_out, intents, vol_map):
    wb  = __import__('openpyxl').Workbook()
    ws  = wb.active
    ws.title = 'Keyword Mapping'

    gh  = PatternFill(start_color='1D9E75', end_color='1D9E75', fill_type='solid')
    gr  = PatternFill(start_color='EAF3DE', end_color='EAF3DE', fill_type='solid')
    bl  = PatternFill(start_color='E8F0FE', end_color='E8F0FE', fill_type='solid')
    yl  = PatternFill(start_color='FFFBEA', end_color='FFFBEA', fill_type='solid')
    hf  = Font(bold=True, color='FFFFFF', size=11)
    lf  = Font(size=10, color='0563C1')
    bf  = Font(size=10)
    mf  = Font(size=10, italic=True, color='888888')

    headers = ['Keyword','Search Volume','Landing Page','Intent','Match Source','Match Score']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = gh; c.font = hf; c.alignment = Alignment(horizontal='left')

    for i, (kw, url, src, sc, intent) in enumerate(zip(keywords, mapped_urls, sources, scores_out, intents)):
        r   = i + 2
        vol = vol_map.get(kw.lower(), 0)
        if url and '/blog/' in url:       fill = bl
        elif src == 'GSC fallback':        fill = yl
        elif url:                          fill = gr
        else:                              fill = PatternFill(fill_type=None)

        ws.cell(row=r, column=1, value=kw).font   = bf
        ws.cell(row=r, column=2, value=vol).font   = bf
        ws.cell(row=r, column=3, value=url or None).font = lf if url else bf
        ws.cell(row=r, column=4, value=intent).font = Font(size=10, color='555555', italic=(intent=='Informational'))
        ws.cell(row=r, column=5, value=src or None).font = mf
        ws.cell(row=r, column=6, value=sc if sc else None).font = bf
        for col in range(1, 7):
            ws.cell(row=r, column=col).fill = fill

    ws.column_dimensions['A'].width = 52
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 68
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 13
    ws.freeze_panes = 'A2'

    # Content Gap Analysis tab
    ws2 = wb.create_sheet('Content Gap Analysis')
    ws2['A1'] = 'Content Gap Analysis — New Pages & Blog Posts Needed'
    ws2['A1'].font = Font(bold=True, size=13, color='0F6E56')
    ws2.merge_cells('A1:F1')

    for col, h in enumerate(['Topic','Intent','# Keywords','Total Volume','Priority','Page Type Needed'], 1):
        c = ws2.cell(row=3, column=col, value=h)
        c.fill = gh; c.font = hf

    unmapped = [(keywords[i], vol_map.get(keywords[i].lower(), 0),
                 classify_topic(keywords[i]), intents[i])
                for i in range(len(keywords)) if not mapped_urls[i]]

    if unmapped:
        gap_df      = pd.DataFrame(unmapped, columns=['Keyword','Volume','Topic','Intent'])
        gap_summary = (gap_df.groupby(['Topic','Intent'])
                       .agg(Keywords=('Keyword','count'), Total_Volume=('Volume','sum'))
                       .sort_values('Total_Volume', ascending=False).reset_index())

        amber = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
        red   = PatternFill(start_color='FDECEA', end_color='FDECEA', fill_type='solid')

        for i, row in gap_summary.iterrows():
            r   = i + 4
            vol = int(row['Total_Volume'])
            pri = 'High' if vol > 200000 else ('Medium' if vol > 50000 else 'Low')
            pn  = 'Blog / Guide' if row['Intent'] == 'Informational' else 'Service Page'
            fill = red if pri == 'High' else (amber if pri == 'Medium' else gr)
            ws2.cell(row=r, column=1, value=row['Topic']).font   = bf
            ws2.cell(row=r, column=2, value=row['Intent']).font  = bf
            ws2.cell(row=r, column=3, value=int(row['Keywords'])).font = bf
            ws2.cell(row=r, column=4, value=vol).font            = bf
            ws2.cell(row=r, column=5, value=pri).font            = Font(size=10, bold=True)
            ws2.cell(row=r, column=6, value=pn).font             = Font(size=10, color='0563C1' if 'Blog' in pn else '333333')
            for col in range(1, 7): ws2.cell(row=r, column=col).fill = fill

        dr = len(gap_summary) + 6
        ws2.cell(row=dr, column=1, value='All Unmapped Keywords').font = Font(bold=True, size=12, color='0F6E56')
        ws2.merge_cells(start_row=dr, start_column=1, end_row=dr, end_column=4)
        dr += 1
        for col, h in enumerate(['Topic','Intent','Keyword','Search Volume'], 1):
            c = ws2.cell(row=dr, column=col, value=h); c.fill = gh; c.font = hf
        dr += 1
        for _, row in gap_df.sort_values(['Topic','Intent','Volume'], ascending=[True,True,False]).iterrows():
            ws2.cell(row=dr, column=1, value=row['Topic']).font  = bf
            ws2.cell(row=dr, column=2, value=row['Intent']).font = Font(size=10, italic=(row['Intent']=='Informational'))
            ws2.cell(row=dr, column=3, value=row['Keyword']).font = bf
            ws2.cell(row=dr, column=4, value=int(row['Volume'])).font = bf
            dr += 1

    for col, w in zip('ABCDEF', [28,16,52,18,12,16]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

st.title("🔍 SEO Content Mapper")
st.markdown("Map untapped keywords to existing pages based on URL, Title, Meta Description and H1.")

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    st.markdown("---")

    threshold = st.slider(
        "Match Score Threshold",
        min_value=0.10, max_value=0.40, value=0.15, step=0.01,
        help="Minimum similarity score to accept a match. Lower = more matches but weaker. Higher = fewer but stronger."
    )

    st.markdown("**Content Weights**")
    st.caption("How much each element influences matching")
    slug_w  = st.slider("URL Slug",          1, 8, 5)
    title_w = st.slider("Page Title",        1, 6, 3)
    h1_w    = st.slider("H1 Heading",        1, 4, 2)
    meta_w  = st.slider("Meta Description",  1, 3, 1)
    weights = (slug_w, title_w, h1_w, meta_w)

    st.markdown("---")
    st.markdown("**Score Guide**")
    st.markdown("""
| Score | Quality |
|-------|---------|
| 0.50+ | Strong ✅ |
| 0.30–0.50 | Good ✅ |
| 0.20–0.30 | Acceptable ⚠️ |
| 0.15–0.20 | Weak — verify |
| 0.40 (GSC) | Fallback |
""")

    st.markdown("---")
    st.markdown("**Colour Coding**")
    st.markdown("🟢 Green = Service page match")
    st.markdown("🔵 Blue = Blog match")
    st.markdown("🟡 Yellow = GSC fallback")
    st.markdown("⬜ White = Content gap")

# ── MAIN AREA — FILE UPLOADS ──────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 File 1 — GSC Data")
    st.caption("Your Google Search Console export. Used to cross-reference existing rankings.")
    gsc_file = st.file_uploader("Upload GSC Excel file", type=['xlsx'], key='gsc')

    if gsc_file:
        try:
            gsc_xl     = pd.read_excel(gsc_file, sheet_name=None)
            gsc_sheets = list(gsc_xl.keys())
            gsc_sheet  = st.selectbox("Select GSC sheet", gsc_sheets, key='gsc_sheet')
            gsc_df_raw = gsc_xl[gsc_sheet]
            gsc_cols   = list(gsc_df_raw.columns)

            with st.expander("Map columns"):
                q_col   = st.selectbox("Query column",         gsc_cols, index=gsc_cols.index('Query')         if 'Query'         in gsc_cols else 0, key='q')
                lp_col  = st.selectbox("Landing Page column",  gsc_cols, index=gsc_cols.index('Landing Page')  if 'Landing Page'  in gsc_cols else 1, key='lp')
                clk_col = st.selectbox("Clicks column",        gsc_cols, index=gsc_cols.index('Url Clicks')    if 'Url Clicks'    in gsc_cols else 2, key='clk')
                imp_col = st.selectbox("Impressions column",   gsc_cols, index=gsc_cols.index('Impressions')   if 'Impressions'   in gsc_cols else 3, key='imp')

            st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(gsc_df_raw):,}</strong> GSC rows from <strong>{gsc_sheet}</strong></div>', unsafe_allow_html=True)
        except Exception as e:
            st.markdown(f'<div class="error-box">❌ Error reading file: {e}</div>', unsafe_allow_html=True)
            gsc_file = None

with col2:
    st.subheader("📂 File 2 — Keywords + URLs")
    st.caption("Sheet 1: keywords to map. Sheet 2: your URL index with Title, Meta, H1.")
    kw_file = st.file_uploader("Upload Keywords + URLs Excel file", type=['xlsx'], key='kw')

    if kw_file:
        try:
            kw_xl      = pd.read_excel(kw_file, sheet_name=None)
            kw_sheets  = list(kw_xl.keys())

            st.markdown("**Keywords sheet**")
            kw_sheet   = st.selectbox("Select keywords sheet", kw_sheets, key='kw_sheet')
            kw_df_raw  = kw_xl[kw_sheet]
            kw_cols    = list(kw_df_raw.columns)

            with st.expander("Map keyword columns"):
                kw_col  = st.selectbox("Keyword column", kw_cols,
                    index=kw_cols.index('Keyword') if 'Keyword' in kw_cols else 0, key='kwc')
                vol_options = ['(none — no volume column)'] + kw_cols
                vol_default = (kw_cols.index('Volume') + 1) if 'Volume' in kw_cols else \
                              (kw_cols.index('Search Volume') + 1) if 'Search Volume' in kw_cols else 0
                vol_col_sel = st.selectbox("Search Volume column (optional)", vol_options,
                    index=vol_default, key='vc')
                vol_col = None if vol_col_sel == '(none — no volume column)' else vol_col_sel

            st.markdown("**URLs sheet**")
            # Default to 'Unique URLs' sheet if it exists, else second sheet
            url_sheet_default = 0
            if 'Unique URLs' in kw_sheets:
                url_sheet_default = kw_sheets.index('Unique URLs')
            elif len(kw_sheets) > 1:
                url_sheet_default = 1
            url_sheet  = st.selectbox("Select URL index sheet", kw_sheets, index=url_sheet_default, key='url_sheet')
            url_df_raw = kw_xl[url_sheet]
            url_cols   = list(url_df_raw.columns)

            with st.expander("Map URL columns"):
                ulp_col  = st.selectbox("Landing Page column",     url_cols, index=url_cols.index('Landing Page')     if 'Landing Page'     in url_cols else 0, key='ulp')
                utt_col  = st.selectbox("Page Title column",       url_cols, index=url_cols.index('Page Title')       if 'Page Title'       in url_cols else 1, key='utt')
                umd_col  = st.selectbox("Meta Description column", url_cols, index=url_cols.index('Meta Description') if 'Meta Description' in url_cols else 2, key='umd')
                uh1_col  = st.selectbox("H1 column",               url_cols, index=url_cols.index('H1')               if 'H1'               in url_cols else 3, key='uh1')

            st.markdown(f'<div class="info-box">✅ <strong>{len(kw_df_raw):,}</strong> keywords | <strong>{len(url_df_raw):,}</strong> URLs</div>', unsafe_allow_html=True)
        except Exception as e:
            st.markdown(f'<div class="error-box">❌ Error reading file: {e}</div>', unsafe_allow_html=True)
            kw_file = None

# ── RUN BUTTON ────────────────────────────────────────────────────────────────
st.markdown("---")

can_run = gsc_file is not None and kw_file is not None
if not can_run:
    st.markdown('<div class="warn-box">⚠️ Upload both files above to enable mapping.</div>', unsafe_allow_html=True)

if st.button("🚀 Run Mapping", disabled=not can_run, use_container_width=True, type="primary"):

    with st.spinner("Preparing data..."):

        # Build GSC map
        gsc_clean = gsc_df_raw.rename(columns={q_col:'Query', lp_col:'Landing Page',
                                                clk_col:'Url Clicks', imp_col:'Impressions'})
        gsc_clean['kw_lower'] = gsc_clean['Query'].str.lower().str.strip()
        gsc_map = {}
        for kw_l, grp in gsc_clean.groupby('kw_lower'):
            best = grp.sort_values(['Url Clicks','Impressions'], ascending=False).iloc[0]
            gsc_map[kw_l] = best['Landing Page']

        # Build keyword list + volume map
        kw_df_clean = kw_df_raw.rename(columns={kw_col: 'Keyword'})
        if vol_col:
            kw_df_clean = kw_df_clean.rename(columns={vol_col: 'Volume'})
        else:
            kw_df_clean['Volume'] = 0
        keywords = kw_df_clean['Keyword'].fillna('').tolist()
        vol_map  = dict(zip(kw_df_clean['Keyword'].str.lower(), kw_df_clean['Volume'].fillna(0)))

        # Build URL df
        url_df_clean = url_df_raw.rename(columns={
            ulp_col:'Landing Page', utt_col:'Page Title',
            umd_col:'Meta Description', uh1_col:'H1'
        })

    progress_bar = st.progress(0)
    status_text  = st.empty()

    mapped_urls, sources, scores_out, intents = run_mapping(
        keywords, vol_map, url_df_clean, gsc_map,
        threshold, weights, progress_bar, status_text
    )

    progress_bar.progress(100)
    status_text.text("Done!")

    # ── RESULTS METRICS ──────────────────────────────────────────────────────
    mapped_count   = sum(1 for u in mapped_urls if u)
    gsc_conf       = sum(1 for s in sources if s == 'Content + GSC confirmed')
    content_only   = sum(1 for s in sources if s == 'Content match')
    gsc_fb         = sum(1 for s in sources if s == 'GSC fallback')
    unmapped_count = len(keywords) - mapped_count

    st.markdown("---")
    st.subheader("📊 Results")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.markdown(f'<div class="metric-card"><div class="metric-num">{len(keywords):,}</div><div class="metric-label">Total Keywords</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="metric-card"><div class="metric-num">{mapped_count:,}</div><div class="metric-label">Mapped</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="metric-card"><div class="metric-num">{unmapped_count:,}</div><div class="metric-label">Content Gaps</div></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="metric-card"><div class="metric-num">{gsc_conf:,}</div><div class="metric-label">Content + GSC</div></div>', unsafe_allow_html=True)
    c5.markdown(f'<div class="metric-card"><div class="metric-num">{gsc_fb:,}</div><div class="metric-label">GSC Fallback</div></div>', unsafe_allow_html=True)

    st.markdown("")

    # ── PREVIEW TABLE ────────────────────────────────────────────────────────
    preview_data = []
    for kw, url, src, sc, intent in zip(keywords, mapped_urls, sources, scores_out, intents):
        preview_data.append({
            'Keyword':      kw,
            'Volume':       vol_map.get(kw.lower(), 0),
            'Landing Page': url or '',
            'Intent':       intent,
            'Source':       src or '',
            'Score':        sc if sc else ''
        })

    preview_df = pd.DataFrame(preview_data)
    mapped_df  = preview_df[preview_df['Landing Page'] != ''].sort_values('Score', ascending=False)
    gaps_df    = preview_df[preview_df['Landing Page'] == '']

    tab1, tab2 = st.tabs([f"✅ Mapped ({mapped_count:,})", f"🔴 Content Gaps ({unmapped_count:,})"])

    with tab1:
        st.dataframe(mapped_df, use_container_width=True, height=400)

    with tab2:
        gap_topics = gaps_df['Keyword'].apply(classify_topic)
        gap_vols   = gaps_df['Volume']
        summary    = pd.DataFrame({'Topic': gap_topics, 'Volume': gap_vols})
        summary    = summary.groupby('Topic').agg(
            Keywords=('Topic','count'), Total_Volume=('Volume','sum')
        ).sort_values('Total_Volume', ascending=False).reset_index()
        st.dataframe(summary, use_container_width=True, height=300)
        st.caption("Full list of unmapped keywords is in the downloaded Excel file.")

    # ── DOWNLOAD ─────────────────────────────────────────────────────────────
    st.markdown("---")
    excel_buf = build_excel(keywords, mapped_urls, sources, scores_out, intents, vol_map)

    st.download_button(
        label="⬇️ Download Full Excel Output",
        data=excel_buf,
        file_name="seo_content_mapping_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    st.markdown("""
    <div class="info-box">
    <strong>Output file contains:</strong><br>
    • <strong>Keyword Mapping</strong> tab — all keywords with Landing Page, Intent, Match Source, Match Score. Sorted by score.<br>
    • <strong>Content Gap Analysis</strong> tab — unmapped keywords grouped by topic with total search volume and priority.
    </div>
    """, unsafe_allow_html=True)
