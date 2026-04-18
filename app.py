import streamlit as st
import pandas as pd
import numpy as np
import re, io, json, time, urllib.request
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
HVAC_GRP  = {'furnace','boiler','heat_pump','ac_cooling','ductless','hvac_general','geothermal'}
PLMB_GRP  = {'water_pipes','drain_sewer','water_heater','toilet','faucet_sink','shower_tub','bathroom','kitchen','plumbing_general','backflow','sump_pump','sprinkler'}
ELEC_GRP  = {'electrical','generator','ev_charger'}
ALL_EX    = HVAC_GRP | PLMB_GRP | ELEC_GRP | {'gas_utility','air_quality'}

# FIX 2: Added /tag/ to exclusion list
EXCL = ['/truck-wrap','/careers','/join-the-team','/contact','/podcast','/coupons',
        '/warranty','/about-us','/meet-the','/technicians','/community','/financing',
        '/solar','/estore','/home-energy','/lancaster-neighbors','/tonawanda-dispatch',
        '/hamburg-service','/buffalo-service','/buffalo-home-improvement',
        '/case_study_category','why-choose-us','/walker','/our-community',
        '/renovation-services','/case-stud','/specials','/tag/']  # FIX 2

EXPAND = [(r'\bac\b','air conditioning'),(r'\ba/c\b','air conditioning'),
          (r'\bhvac\b','heating cooling air conditioning'),(r'\bfurnace\b','furnace heating'),
          (r'\bheat pump\b','heat pump heating cooling'),(r'\bboiler\b','boiler heating'),
          (r'\bgenerator\b','generator backup power'),(r'\bdrain\b','drain sewer'),
          (r'\bsump pump\b','sump pump basement'),(r'\bgfci\b','gfci outlet electrical'),
          (r'\bbackflow\b','backflow prevention'),(r'\buv\b','uv air sanitizer'),
          (r'\brepiping\b','repiping pipe replacement'),(r'\bwater heater\b','water heater hot water'),
          (r'\btankless\b','tankless water heater'),(r'\bmini.?split\b','mini split ductless heating cooling')]

# FIX 9: Cost/price keywords are informational
INFO_SIG = [r'^what\b',r'^how\b',r'^why\b',r'^when\b',r'^does\b',r'^do\b',
            r'^is\b',r'^are\b',r'^can\b',r'^should\b',r'^which\b',r'^who\b',
            r'\bvs\b',r'\bversus\b',r'\bdifference between\b',r'\btips\b',
            r'\bbenefits of\b',r'\bsigns\b',r'\bcauses\b',r'\btypes of\b',
            r'\bhow to\b',r'\bwhat is\b',r'\bwhy is\b',r'\bhow does\b',
            r'\bwhat does\b',r'\badvantages\b',r'\bproblems with\b',
            r'\bcost of\b',r'\bcost to\b',r'\bhow much\b',r'\bprice of\b',  # FIX 9
            r'\bpricing\b',r'\baverage cost\b',r'\bwhat does .+ cost\b']    # FIX 9

TOPIC_MAP = {
    'Furnace / Heating': ['furnace heating','furnace','boiler heating','boiler','radiant heat','steam heat','forced air','heating system'],
    'AC / Cooling':      ['air conditioning','cooling','central air','mini split ductless','ductless','evaporator coil','refrigerant'],
    'Heat Pump':         ['heat pump heating cooling'],
    'Water Heater':      ['water heater hot water','water heater','tankless water heater','hot water'],
    'Drain / Sewer':     ['drain sewer','sewer','clog','unclog','jetting','hydrojet','rooter','sewage','septic'],
    'Plumbing General':  ['plumb','pipe','leak','repiping','water line','water main','burst pipe'],
    'Electrical':        ['electric','electrician','panel','wiring','outlet','gfci outlet electrical','circuit breaker','lighting','surge'],
    'Generator':         ['generator backup power'],
    'Sump Pump':         ['sump pump basement'],
    'Air Quality':       ['air quality','dehumidif','humidif','air purif','uv air sanitizer','indoor air'],
    'Water Treatment':   ['soft water','hard water','water treatment','water softener','water filter'],
    'Backflow':          ['backflow prevention'],
    'Geothermal':        ['geothermal'],
    'Bathroom/Kitchen':  ['bathroom','kitchen','toilet','shower','faucet','sink','tub','garbage disposal'],
    'Ventilation':       ['ventilation','hrv','erv','air handler','duct'],
    'EV Charger':        ['ev charger','electric vehicle charg','charging station','car charging'],
    'Commercial':        ['commercial'],
    'Location':          ['cheektowaga','amherst','hamburg','lancaster','west seneca','orchard park',
                          'east aurora','clarence','depew','tonawanda','williamsville','batavia','grand island'],
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def expand(t):
    t = t.lower()
    for p,r in EXPAND: t = re.sub(p,r,t)
    return t

def is_excl(url):
    return any(p in str(url).lower() for p in EXCL)

def is_info(kw):
    return any(re.search(p, kw.lower()) for p in INFO_SIG)

def get_url(row):
    return str(row.get('Address', row.get('Landing Page','')))

def slug_terms(url, stop=None):
    s = {'com','https','http','in','and','the','a','of','for','to','near','me','www'} | (stop or set())
    p = re.sub(r'https?://[^/]+','',str(url))
    return ' '.join(x for x in re.split(r'[/\-_]',p) if x and x.lower() not in s)

def build_content(row, weights, stop=None):
    url   = get_url(row)
    slug  = expand(slug_terms(url, stop))
    title = expand(str(row.get('Title 1', row.get('Page Title','')) or ''))
    meta  = expand(str(row.get('Meta Description 1', row.get('Meta Description','')) or ''))
    h1r   = str(row.get('H1-1', row.get('H1','')) or '')
    h1    = expand(h1r) if h1r.lower() not in ('nan','none','') else ''
    sw,tw,hw,mw = weights
    return f"{slug} "*sw + f"{title} "*tw + f"{h1} "*hw + meta*mw

def classify_topic(kw):
    ke = expand(kw.lower())
    for t,terms in TOPIC_MAP.items():
        if any(x in ke for x in terms): return t
    return 'Other'

# FIX 9: Enhanced intent — cost/price = informational
def classify_intent(kw):
    kl = kw.lower()
    cost_terms = ['cost','price','pricing','how much','average cost','cost of','cost to','what does it cost']
    if any(t in kl for t in cost_terms): return 'Informational'
    return 'Informational' if is_info(kw) else 'Transactional'

def make_fill(c): return PatternFill(start_color=c,end_color=c,fill_type='solid')
def hdr(ws,row,cols,bg='1D9E75',fg='FFFFFF'):
    for i,v in enumerate(cols,1):
        c=ws.cell(row=row,column=i,value=v)
        c.fill=make_fill(bg); c.font=Font(bold=True,color=fg,size=11)
        c.alignment=Alignment(horizontal='left',vertical='center')
def cw(ws,ws_list):
    for col,w in zip('ABCDEFGHIJKLMNOPQRST',ws_list): ws.column_dimensions[col].width=w

# ─────────────────────────────────────────────────────────────────────────────
# URL TAXONOMY
# ─────────────────────────────────────────────────────────────────────────────
def cat_url(url):
    p = re.sub(r'https?://[^/]+','',url.lower())
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
    # FIX 8: commercial-hvac restricted to commercial context
    if any(t in p for t in ['commercial-hvac','duct-cleaning','air-duct','air-handler','trane']): c.add('hvac_general')
    if 'commercial-hvac' in p: c.add('commercial')  # FIX 8
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
    for loc in ['cheektowaga','amherst','hamburg','lancaster','west-seneca','orchard-park','east-aurora',
                'springville','alden','akron','batavia','clarence','depew','elma','eden','marilla',
                'holland','wales','darien','medina','attica','boston-ny','grand-island','bennington',
                'west-valley','williamsville','tonawanda']:
        if loc in p: c.add('location'); c.add('loc:'+loc.replace('-',' ')); break
    if '/blog/' in p: c.add('blog')
    # FIX 7: Brand detection
    for brand in ['trane','mitsubishi','cummins','rheem','bradford','lennox']:
        if brand in p: c.add('brand:'+brand)
    if not (c - {'blog'}): c.add('plumbing_general')
    return frozenset(c)

def cat_kw(kw):
    kl = expand(kw.lower())
    c  = set()
    GAS = ['gas leak','gas line','gas pipe','natural gas line','gas service line','gas detector',
           'gas fireplace','gas stove','gas dryer','gas connection','gas hookup','gas meter',
           'carbon monoxide detector','gas shut off','propane line','natural gas leak']
    if any(t in kl for t in GAS):
        if not any(x in kl for x in ['furnace heating','boiler heating','water heater hot water']): c.add('gas_utility')
    WH = ['water heater hot water','water heater','hot water heater','tankless water heater','hot water tank',
          'water heater leak','water heater drip','water heater temp','water heater not','no hot water','water not heating']
    if any(t in kl for t in WH): c.add('water_heater')
    if 'water_heater' not in c:
        WP = ['water leak','water line','water main','water pipe','burst pipe','pipe leak','slab leak',
              'repiping','main water','frozen pipe','water service line','water shut off','water pressure']
        if any(t in kl for t in WP): c.add('water_pipes')
    DR = ['drain sewer','sewer','unclog','hydro jet','rooter','sewage','septic','main line clog',
          'backed up drain','slow drain','floor drain','bubbler','storm drain','exterior drain',
          'interior drain','basement drain','laundry drain','bathroom drain','kitchen drain']
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
    # FIX 8: commercial keyword detection
    if any(t in kl for t in ['commercial','industrial','business facility','office building','warehouse']): c.add('commercial')
    is_combo = any(x in kl for x in ['water heater','furnace heating','boiler heating','heat pump heating'])
    EL = ['electrician','electric panel','breaker box','circuit breaker','electrical panel','wiring',
          'rewiring','outlet','gfci outlet electrical','light fixture','lighting','surge protect',
          'electrical inspect','electrical service','electrical repair','electrical install',
          'electrical work','electrical contractor','electric repair','residential electric',
          'home electrical','electrical permit','electrical company','local electrician']
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
    # FIX 7: Brand keyword detection
    for brand in ['trane','mitsubishi','cummins','rheem','bradford','lennox']:
        if brand in kl: c.add('brand:'+brand)
    if 'heater' in kl and 'water heater' not in kl and 'water_heater' not in c: c.update({'furnace','hvac_general'})
    if re.search(r'\bheating\b',kl) and not (c & {'furnace','boiler','heat_pump','hvac_general','water_heater'}): c.update({'furnace','hvac_general'})
    if re.search(r'\bcooling\b',kl) and not (c & {'ac_cooling','ductless','heat_pump','hvac_general'}): c.update({'ac_cooling','hvac_general'})
    return frozenset(c)

def is_compat(kc, uc, kw, url):
    # Location pages
    if 'location' in uc:
        kl = expand(kw.lower())
        locs = [t for t in uc if t.startswith('loc:')]
        return bool(locs) and locs[0].replace('loc:','') in kl
    # FIX 7: Brand pages only match brand keywords
    url_brands = {t for t in uc if t.startswith('brand:')}
    kw_brands  = {t for t in kc if t.startswith('brand:')}
    if url_brands and not (url_brands & kw_brands): return False
    # FIX 8: Commercial HVAC page only for commercial keywords
    if 'commercial' in uc and 'commercial-hvac' in url.lower():
        if 'commercial' not in kc: return False
    # Gas exclusive
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
    for p in [r'\bnear me\b',r'\bnear\b']: kw = re.sub(p,'',kw)
    return re.sub(r'\s+',' ',kw).strip()

# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE API
# ─────────────────────────────────────────────────────────────────────────────
def call_claude(api_key, prompt, max_tokens=2000):
    payload = json.dumps({"model":"claude-sonnet-4-20250514","max_tokens":max_tokens,
                          "messages":[{"role":"user","content":prompt}]}).encode()
    req = urllib.request.Request("https://api.anthropic.com/v1/messages", data=payload,
        headers={"Content-Type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"},method="POST")
    with urllib.request.urlopen(req,timeout=60) as r:
        data = json.loads(r.read())
    txt = data['content'][0]['text'].strip()
    txt = re.sub(r'^```json?\s*','',txt); txt = re.sub(r'\s*```$','',txt)
    return json.loads(txt.strip())

def claude_relevance(api_key, keywords, biz_desc, excl_str, status_text, progress_bar, p0, p1):
    out = {}
    batches = [keywords[i:i+100] for i in range(0,len(keywords),100)]
    for bi,batch in enumerate(batches):
        progress_bar.progress(min(p0+int((bi/len(batches))*(p1-p0)),p1-1))
        status_text.text(f"Claude API — relevance check: batch {bi+1}/{len(batches)}...")
        excl_note = f"\nServices NOT offered: {excl_str}" if excl_str.strip() else ""
        prompt = f"""SEO analyst. Classify each keyword relevance for this business.
Business: {biz_desc}{excl_note}
Labels: RELEVANT | NOT_RELEVANT | BORDERLINE
Return ONLY JSON: {{"keyword":"LABEL",...}}
Keywords: {json.dumps(batch)}"""
        for attempt in range(3):
            try: out.update(call_claude(api_key,prompt,1500)); break
            except:
                if attempt<2: time.sleep(3)
                else:
                    for k in batch: out[k]='BORDERLINE'
        time.sleep(0.1)
    return out

def claude_validate_blogs(api_key, kw_blog_pairs, status_text, progress_bar, p0, p1):
    """Validate whether a keyword actually matches the blog it was mapped to."""
    out = {}
    batches = [kw_blog_pairs[i:i+50] for i in range(0,len(kw_blog_pairs),50)]
    for bi,batch in enumerate(batches):
        progress_bar.progress(min(p0+int((bi/len(batches))*(p1-p0)),p1-1))
        status_text.text(f"Claude API — blog validation: batch {bi+1}/{len(batches)}...")
        pairs_str = json.dumps([{"keyword":k,"blog_slug":s} for k,s in batch])
        prompt = f"""SEO analyst. For each keyword-blog pair, answer if the blog post would be a good match for that keyword.
Answer YES if the blog post clearly covers what the searcher wants to know.
Answer NO if the blog post topic is different from the keyword intent.
Return ONLY JSON: {{"keyword":"YES/NO",...}}
Pairs: {pairs_str}"""
        for attempt in range(3):
            try: out.update(call_claude(api_key,prompt,1000)); break
            except:
                if attempt<2: time.sleep(3)
                else:
                    for k,s in batch: out[k]='YES'
        time.sleep(0.1)
    return out

def claude_match_blogs(api_key, keywords_info, existing_blogs, status_text, progress_bar, p0, p1):
    """FIX C: Semantically match unmapped informational kws to existing blogs."""
    out = {}
    blog_list = json.dumps([b.replace('https://cellinoplumbing.com','') for b in existing_blogs[:60]])
    batches = [keywords_info[i:i+80] for i in range(0,len(keywords_info),80)]
    for bi,batch in enumerate(batches):
        progress_bar.progress(min(p0+int((bi/len(batches))*(p1-p0)),p1-1))
        status_text.text(f"Claude API — blog semantic matching: batch {bi+1}/{len(batches)}...")
        prompt = f"""SEO analyst. For each keyword, find the best matching blog URL from the list, or return null if none fits well.
Only return a URL if the blog CLEARLY covers that keyword topic (90%+ match).
Return ONLY JSON: {{"keyword":"blog_url_or_null",...}}
Available blogs: {blog_list}
Keywords: {json.dumps(batch)}"""
        for attempt in range(3):
            try: out.update(call_claude(api_key,prompt,1500)); break
            except:
                if attempt<2: time.sleep(3)
                else:
                    for k in batch: out[k]=None
        time.sleep(0.1)
    return out

def claude_cluster(api_key, keywords_with_vol, content_type, status_text, progress_bar, p0, p1):
    """Semantic clustering of unmapped keywords into topic clusters."""
    out = []
    batches = [keywords_with_vol[i:i+120] for i in range(0,len(keywords_with_vol),120)]
    for bi,batch in enumerate(batches):
        progress_bar.progress(min(p0+int((bi/len(batches))*(p1-p0)),p1-1))
        status_text.text(f"Claude API — semantic clustering ({content_type}): batch {bi+1}/{len(batches)}...")
        kw_list = [{"keyword":k,"volume":v} for k,v in batch]
        prompt = f"""SEO content strategist. Group these keywords into topic clusters for {content_type} creation.

Rules:
1. Each cluster = exactly ONE piece of content (one blog post or one service page)
2. Group keywords that describe the SAME entity AND SAME searcher intent
3. Aim for 3-8 keywords per cluster. Never put 20+ keywords in one cluster.
4. Primary keyword = highest volume with clearest intent for that content piece
5. Content type options: "Blog post" or "Service page"
6. Intent options: "Educational" | "Problem-aware" | "Comparison" | "Transactional" | "Cost/Pricing"
7. Suggested title = an SEO-optimised title for that content piece

Return ONLY valid JSON array:
[{{"cluster_name":"short topic name","entity":"main service entity","intent":"intent type","content_type":"Blog post or Service page","suggested_title":"SEO title for the content","primary_keyword":"best kw","primary_volume":1000,"secondary_keywords":["kw2","kw3"],"total_volume":3000}}]

Keywords to cluster:
{json.dumps(kw_list)}"""
        for attempt in range(3):
            try:
                result = call_claude(api_key,prompt,3000)
                if isinstance(result,list): out.extend(result)
                break
            except:
                if attempt<2: time.sleep(5)
        time.sleep(0.2)
    return out

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 1 — VALIDATE GSC MAPPINGS
# ─────────────────────────────────────────────────────────────────────────────
def phase1_gsc(gsc_df, sf_df, weights, status_text, progress_bar):
    status_text.text("Phase 1: Building URL index from Screaming Frog...")
    progress_bar.progress(3)

    stop = set()
    if len(sf_df):
        m = re.search(r'https?://([^/]+)',str(get_url(sf_df.iloc[0])))
        if m: stop = set(m.group(1).replace('www.','').split('.'))

    url_df = sf_df[~sf_df.apply(lambda r: is_excl(get_url(r)),axis=1)].reset_index(drop=True)
    url_df['content'] = url_df.apply(lambda r: build_content(r,weights,stop),axis=1)
    url_df['cats']    = url_df.apply(lambda r: cat_url(get_url(r)),axis=1)
    url_df['ptype']   = url_df.apply(lambda r: 'blog' if '/blog/' in get_url(r).lower() else 'service',axis=1)

    content_map = {get_url(r).lower().strip(): url_df.at[i,'content'] for i,r in url_df.iterrows()}

    # Deduplicate GSC — best URL per query by clicks then impressions
    gsc_df['kl'] = gsc_df['Query'].str.lower().str.strip()
    dedup = {}
    for kl,grp in gsc_df.groupby('kl'):
        best = grp.sort_values(['Clicks','Impressions'],ascending=False).iloc[0]
        dedup[kl] = {
            'url':   best['Landing Page'],
            'clicks': best['Clicks'],
            'imps':   best['Impressions'],
            'pos':    best.get('Position', np.nan),
            'ctr':    best.get('CTR', 0),
        }

    progress_bar.progress(8)
    status_text.text("Phase 1: Scoring GSC queries against page content...")

    queries   = list(dedup.keys())
    cleaned_q = [clean_kw(q) for q in queries]
    all_c     = list(content_map.values())
    url_keys  = list(content_map.keys())

    vec = TfidfVectorizer(ngram_range=(1,3),min_df=1,sublinear_tf=True)
    vec.fit(all_c + cleaned_q)
    uv = vec.transform(all_c)
    qv = vec.transform(cleaned_q)
    sims = cosine_similarity(qv,uv)
    uk   = {k:i for i,k in enumerate(url_keys)}

    validated = []
    for i,(q,cq) in enumerate(zip(queries,cleaned_q)):
        if i%2000==0:
            progress_bar.progress(min(8+int((i/len(queries))*15),22))
            status_text.text(f"Phase 1: Validating {i:,}/{len(queries):,} GSC queries...")
        d = dedup[q]
        gu = str(d['url']).lower().strip()
        idx = uk.get(gu)
        score = round(float(sims[i,idx]),4) if idx is not None else 0.0
        status = 'Confirmed' if score>=0.30 else ('Plausible' if score>=0.15 else 'Suspicious')
        pos = d['pos']
        try: pos_val = round(float(pos),1)
        except: pos_val = None
        validated.append({'Query':q,'Mapped URL':d['url'],'Clicks':d['clicks'],
                          'Impressions':d['imps'],'Position':pos_val,'CTR':d['ctr'],
                          'Content Score':score,'Mapping Status':status})

    progress_bar.progress(23)
    return pd.DataFrame(validated), url_df, dedup, stop

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 2 — MAP SEMRUSH KEYWORDS (all fixes applied)
# ─────────────────────────────────────────────────────────────────────────────
def phase2_semrush(sem_df, url_df, gsc_dedup, gsc_val_df, weights, threshold, your_col,
                   status_text, progress_bar):
    status_text.text("Phase 2: Building TF-IDF models...")
    progress_bar.progress(25)

    svc_df  = url_df[url_df['ptype']=='service'].reset_index(drop=True)
    blog_df = url_df[url_df['ptype']=='blog'].reset_index(drop=True)

    # Get list of homepage URLs to deprioritise (FIX 3)
    hp_urls = set()
    for i,r in url_df.iterrows():
        u = get_url(r)
        slug = re.sub(r'https?://[^/]+','',u).strip('/')
        if slug in ('','buffalo','buffalo-ny','buffalo-ny/'): hp_urls.add(u.lower().strip())

    # GSC status map for fix 5
    gsc_status = dict(zip(gsc_val_df['Query'].str.lower().str.strip(), gsc_val_df['Mapping Status']))

    keywords  = sem_df['Keyword'].fillna('').tolist()
    volumes   = sem_df['Volume'].fillna(0).tolist()
    your_pos  = sem_df[your_col].fillna('') if your_col in sem_df.columns else pd.Series(['']*len(keywords))
    your_pos  = your_pos.tolist()
    cleaned   = [clean_kw(k) for k in keywords]

    def build_sims(df):
        v = TfidfVectorizer(ngram_range=(1,3),min_df=1,sublinear_tf=True)
        v.fit(df['content'].tolist()+cleaned)
        return cosine_similarity(v.transform(cleaned),v.transform(df['content']))

    ss = build_sims(svc_df)
    bs = build_sims(blog_df)
    progress_bar.progress(42)
    status_text.text("Phase 2: Mapping keywords to pages...")

    results = []
    for i,kw in enumerate(keywords):
        if i%500==0:
            progress_bar.progress(min(42+int((i/len(keywords))*28),69))
            status_text.text(f"Phase 2: Mapping {i:,}/{len(keywords):,} keywords...")

        kl     = kw.lower().strip()
        intent = classify_intent(kw)  # FIX 9
        kc     = cat_kw(kw)
        chosen=''; source=''; cs=0.0; gs=0.0

        # FIX A: same threshold for blogs as service pages (0.15)
        thr_blog = threshold  # was lower before, now same

        # FIX E: informational keywords never get GSC fallback onto service pages
        # Content matching first
        if intent == 'Informational':
            ranked = np.argsort(bs[i])[::-1]
            for idx in ranked:
                s = float(bs[i][idx])
                if s < thr_blog: break  # FIX A
                url = get_url(blog_df.iloc[idx])
                uc  = blog_df.iloc[idx]['cats']
                if is_compat(kc,uc,kw,url):
                    chosen=url; cs=round(s,4); source='Content match (blog)'; break
        else:
            # FIX 3: Try specific service pages first (exclude homepage), then allow homepage as fallback
            ranked = np.argsort(ss[i])[::-1]
            hp_fallback = None
            for idx in ranked:
                s = float(ss[i][idx])
                if s < threshold: break
                url = get_url(svc_df.iloc[idx])
                uc  = svc_df.iloc[idx]['cats']
                if is_compat(kc,uc,kw,url):
                    if url.lower().strip() in hp_urls:
                        if hp_fallback is None: hp_fallback=(url,round(s,4))
                        continue  # try to find a better specific page first
                    chosen=url; cs=round(s,4); source='Content match'; break
            # FIX 3: Only use homepage if no specific page found
            if not chosen and hp_fallback:
                chosen,cs = hp_fallback; source='Content match (homepage fallback)'

        # GSC cross-reference
        gd  = gsc_dedup.get(kl,{})
        gu  = gd.get('url','')
        guc = cat_url(gu) if gu else frozenset()
        gst = gsc_status.get(kl,'')
        gv  = (gu and not is_excl(gu) and is_compat(kc,guc,kw,gu))

        # FIX 5: GSC fallback with content score 0.0 AND Suspicious → do not use
        # FIX 6: GSC fallback never maps informational keywords to service pages
        if gv:
            gsc_is_blog = '/blog/' in gu.lower()
            # FIX 6: informational → only accept GSC if it points to a blog
            if intent == 'Informational' and not gsc_is_blog:
                gv = False
            # FIX 5: Suspicious GSC with no content match → reject
            if gst == 'Suspicious' and not chosen:
                gv = False

        if gv and chosen and gu.lower().strip()==chosen.lower().strip():
            gs=0.40; source=source.replace('Content match','Content + GSC confirmed')
        elif gv and not chosen:
            # FIX 5: Only use GSC fallback if Confirmed or Plausible
            if gst in ('Confirmed','Plausible'):
                chosen=gu; gs=0.40; source='GSC fallback'

        fs = round(cs*0.7+gs*0.3,4)

        rp = str(your_pos[i]).strip()
        try: pn=float(rp)
        except: pn=None
        if pn and pn>0:
            if pn<=10:    rs='Ranking p1-10'
            elif pn<=20:  rs='Quick win p11-20'
            elif pn<=50:  rs='Weak ranking p21-50'
            else:         rs='Very weak p51-100'
        else: rs='Not ranking'

        results.append({'Keyword':kw,'Volume':int(volumes[i]) if volumes[i] else 0,
                        'Your Position':rp or 'N/A','Landing Page':chosen,
                        'Intent':intent,'Content Score':cs,'GSC Score':gs,
                        'Final Score':fs,'Match Source':source,'Ranking Status':rs,
                        '_cats':kc})

    progress_bar.progress(70)
    return results

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 3 — BUSINESS RELEVANCE + BLOG VALIDATION + BLOG MATCHING
# ─────────────────────────────────────────────────────────────────────────────
def phase3_claude(mapped, url_df, api_key, biz_desc, excl_str, status_text, progress_bar):
    status_text.text("Phase 3: Preparing Claude API checks...")
    progress_bar.progress(72)

    # FIX B: Validate weak blog matches (content score < 0.15)
    weak_blog_pairs = [(r['Keyword'], re.sub(r'https?://[^/]+','',r['Landing Page']))
                       for r in mapped
                       if r['Landing Page'] and '/blog/' in r['Landing Page']
                       and r['Content Score'] < 0.15 and r['Match Source'] != 'GSC fallback']

    blog_valid = {}
    if weak_blog_pairs:
        status_text.text(f"Phase 3a: Validating {len(weak_blog_pairs)} weak blog matches...")
        blog_valid = claude_validate_blogs(api_key, weak_blog_pairs, status_text, progress_bar, 73, 78)

    # FIX C: Semantically match unmapped informational to existing blogs
    existing_blogs = [get_url(r) for _,r in url_df.iterrows() if get_url(r) and '/blog/' in get_url(r).lower()]
    unmapped_info  = [r['Keyword'] for r in mapped
                      if r['Intent']=='Informational' and not r['Landing Page']]

    blog_semantic = {}
    if unmapped_info and existing_blogs:
        # Only run semantic matching on top 400 by volume to keep runtime manageable
        unmapped_info_top = sorted(unmapped_info, key=lambda k: next((r['Volume'] for r in mapped if r['Keyword']==k),0), reverse=True)[:400]
        status_text.text(f"Phase 3b: Semantically matching top {len(unmapped_info_top)} informational keywords to existing blogs...")
        blog_semantic = claude_match_blogs(api_key, unmapped_info_top, existing_blogs, status_text, progress_bar, 78, 84)

    # Business relevance — only truly unmapped + score below 0.15 (not 0.20)
    # Keywords with score 0.15-0.20 are acceptable matches, don't need Claude check
    low_scored = [r['Keyword'] for r in mapped
                  if not r['Landing Page'] or r['Final Score'] < 0.15]
    confirmed  = {r['Keyword']:'RELEVANT' for r in mapped if r['Landing Page'] and r['Final Score'] >= 0.15}

    rel_map = {}
    if low_scored:
        status_text.text(f"Phase 3c: Business relevance check on {len(low_scored):,} keywords...")
        rel_map = claude_relevance(api_key, low_scored, biz_desc, excl_str, status_text, progress_bar, 84, 92)

    all_rel = {**confirmed, **rel_map}

    # Apply fixes to mapped results
    for r in mapped:
        kw = r['Keyword']
        # FIX B: Remove weak blog mappings that Claude says don't match
        if (r['Landing Page'] and '/blog/' in r['Landing Page']
                and r['Content Score'] < 0.15
                and blog_valid.get(kw,'YES') == 'NO'):
            r['Landing Page'] = ''
            r['Match Source'] = ''
            r['Content Score'] = 0.0
            r['Final Score'] = 0.0

        # FIX C: Apply semantic blog matches for previously unmapped informational keywords
        if r['Intent']=='Informational' and not r['Landing Page']:
            sem_url = blog_semantic.get(kw)
            if sem_url:
                # Convert slug back to full URL if needed
                if sem_url and not sem_url.startswith('http'):
                    base = next((get_url(row) for _,row in url_df.iterrows()
                                 if sem_url in get_url(row)), None)
                    if base: sem_url = base
                if sem_url:
                    r['Landing Page'] = sem_url
                    r['Match Source'] = 'Claude semantic match (blog)'
                    r['Final Score'] = 0.25  # assign moderate confidence score

    progress_bar.progress(93)
    return all_rel, mapped

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 4 — SEMANTIC CLUSTERING
# ─────────────────────────────────────────────────────────────────────────────
def phase4_cluster(mapped, rel_map, api_key, status_text, progress_bar):
    status_text.text("Phase 4: Semantic clustering of unmapped keywords...")
    progress_bar.progress(94)

    # Unmapped relevant keywords split by intent
    unmapped_info_all = [(r['Keyword'],r['Volume']) for r in mapped
                      if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in ('RELEVANT','BORDERLINE')
                      and r['Intent']=='Informational']
    unmapped_info = sorted(unmapped_info_all, key=lambda x:-x[1])[:1000]
    # Limit to top 1500 transactional by volume — beyond that, low-volume long tail
    unmapped_trans_all = [(r['Keyword'],r['Volume']) for r in mapped
                      if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in ('RELEVANT','BORDERLINE')
                      and r['Intent']=='Transactional']
    unmapped_trans = sorted(unmapped_trans_all, key=lambda x:-x[1])[:1500]

    clusters = []
    if unmapped_info:
        status_text.text(f"Phase 4a: Clustering {len(unmapped_info)} informational keywords into blog topics...")
        clusters += claude_cluster(api_key, unmapped_info, "blog post", status_text, progress_bar, 94, 97)

    if unmapped_trans:
        status_text.text(f"Phase 4b: Clustering {len(unmapped_trans)} transactional keywords into service topics...")
        clusters += claude_cluster(api_key, unmapped_trans, "service page", status_text, progress_bar, 97, 99)

    # Also build URL-based clusters for mapped keywords
    url_clusters = {}
    for r in mapped:
        if r['Landing Page']:
            url = r['Landing Page']
            if url not in url_clusters:
                url_clusters[url] = []
            url_clusters[url].append(r)

    # Build structured URL cluster list
    url_cluster_list = []
    for url, rows in url_clusters.items():
        rows_sorted = sorted(rows, key=lambda x: (-x.get('Final Score',0), -x.get('Volume',0)))
        # FIX 10/11: Split by intent and sub-intent within URL group
        trans_rows = [r for r in rows_sorted if r['Intent']=='Transactional']
        info_rows  = [r for r in rows_sorted if r['Intent']=='Informational']

        for group, group_intent in [(trans_rows,'Transactional'),(info_rows,'Informational')]:
            if not group: continue
            primary = group[0]
            secondaries = [r['Keyword'] for r in group[1:]]
            total_vol = sum(r['Volume'] for r in group)
            url_cluster_list.append({
                'URL': url,
                'Intent': group_intent,
                'Primary Keyword': primary['Keyword'],
                'Primary Volume': primary['Volume'],
                'Secondary Keywords': ' | '.join(secondaries[:10]),
                'Total Volume': total_vol,
                'Keyword Count': len(group),
                'Best Score': primary.get('Final Score',0),
                'Content Type': 'Blog post' if '/blog/' in url else 'Service page',
            })

    progress_bar.progress(99)
    return clusters, url_cluster_list

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 5 — BUILD EXCEL
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(gsc_df, mapped, rel_map, clusters, url_clusters):
    wb = Workbook()
    bf = Font(size=10); lf = Font(size=10,color='0563C1'); mf = Font(size=10,italic=True,color='888888')
    C = dict(H='1D9E75',GR='EAF3DE',BL='E8F0FE',YL='FFFBEA',RD='FDECEA',
             AM='FFF3CD',PU='F3E8FE',GY='F5F5F5',WH='FFFFFF',DGR='D4EDDA',OR='FFF0E6')

    # ── TAB 1: GSC Mapping Quality ──────────────────────────────────────────
    ws1 = wb.active; ws1.title='GSC Mapping Quality'
    hdr(ws1,1,['Query','Mapped URL','Clicks','Impressions','Position','CTR','Content Score','Mapping Status'])
    sf = {'Confirmed':make_fill(C['GR']),'Plausible':make_fill(C['YL']),'Suspicious':make_fill(C['RD'])}
    for i,row in gsc_df.iterrows():
        r=i+2; fill=sf.get(row['Mapping Status'],make_fill(C['WH']))
        vals=[row['Query'],row['Mapped URL'],row['Clicks'],row['Impressions'],
              row['Position'],row['CTR'],row['Content Score'],row['Mapping Status']]
        for col,v in enumerate(vals,1):
            c=ws1.cell(row=r,column=col,value=v); c.fill=fill; c.font=lf if col==2 else bf
    cw(ws1,[45,65,10,12,10,8,14,14]); ws1.freeze_panes='A2'

    # ── TAB 2: Keyword Mapping ──────────────────────────────────────────────
    ws2=wb.create_sheet('Keyword Mapping')
    hdr(ws2,1,['Keyword','Volume','Your Position','Landing Page','Intent',
               'Content Score','GSC Score','Final Score','Match Source','Ranking Status'])
    for i,r in enumerate(mapped):
        row=i+2; url=r['Landing Page']; src=r['Match Source']; rel=rel_map.get(r['Keyword'],'')
        if url and '/blog/' in url:          fill=make_fill(C['BL'])
        elif src=='GSC fallback':             fill=make_fill(C['YL'])
        elif src and 'Claude semantic' in src:fill=make_fill(C['OR'])
        elif url:                             fill=make_fill(C['GR'])
        elif rel in('RELEVANT','BORDERLINE'): fill=make_fill(C['PU'])
        else:                                 fill=make_fill(C['WH'])
        vals=[r['Keyword'],r['Volume'],r['Your Position'],url or '',r['Intent'],
              r['Content Score'],r['GSC Score'],r['Final Score'],src or '',r['Ranking Status']]
        for col,v in enumerate(vals,1):
            c=ws2.cell(row=row,column=col,value=v); c.fill=fill; c.font=lf if col==4 else bf
    cw(ws2,[50,12,14,65,15,14,12,12,28,18]); ws2.freeze_panes='A2'

    # ── TAB 3: Opportunity Classification ───────────────────────────────────
    ws3=wb.create_sheet('Opportunity Classification')
    hdr(ws3,1,['Keyword','Volume','Your Position','Landing Page','Intent','Final Score','Opportunity Type','Action'])
    OPP_FILL = {'Confirmed existing page':make_fill(C['GR']),'Quick win — optimise':make_fill(C['DGR']),
                'Weak ranking':make_fill(C['YL']),'Page exists — optimise':make_fill(C['BL']),
                'Business relevant gap':make_fill(C['PU']),'True content gap':make_fill(C['RD']),
                'Blog exists — optimise':make_fill(C['OR'])}

    for r in mapped:
        rel=rel_map.get(r['Keyword'],''); url=r['Landing Page']; rs=r['Ranking Status']; fs=r['Final Score']
        src=r.get('Match Source','')
        is_blog = url and '/blog/' in url

        # FIX 1 (revised): Correct action logic
        if url and rs=='Ranking p1-10':
            opp='Confirmed existing page'; action='Monitor — already ranking well'
        elif rs=='Quick win p11-20':
            opp='Quick win — optimise'; action='Optimise title, meta, H1 — almost ranking'
        elif rs in ['Weak ranking p21-50','Very weak p51-100']:
            opp='Weak ranking'; action='Improve page content and build internal links'
        elif url and 'Claude semantic' in src:
            opp='Blog exists — optimise'; action='Update blog title, meta and H1 to better target this keyword'
        elif url and is_blog:
            opp='Page exists — optimise'; action='Optimise existing blog post for this keyword'
        elif url and fs>=0.15:
            # FIX 1: Has correct page → optimise, not create new
            opp='Page exists — optimise'; action='Optimise existing service page for this keyword'
        elif url and fs<0.15:
            # FIX 1: Has page but weak match → still optimise, flag as weak
            opp='Page exists — optimise'; action='Weak match — verify page is correct then optimise'
        elif rel in('RELEVANT','BORDERLINE'):
            opp='Business relevant gap'
            if r['Intent']=='Informational':
                action='Create new blog post'
            else:
                action='Create new service page'
        else:
            opp='True content gap'; action='Evaluate — may need new page'

        fill=OPP_FILL.get(opp,make_fill(C['WH']))
        rn=ws3.max_row+1
        for col,v in enumerate([r['Keyword'],r['Volume'],r['Your Position'],url or '',r['Intent'],fs,opp,action],1):
            c=ws3.cell(row=rn,column=col,value=v); c.fill=fill; c.font=lf if col==4 else bf
    cw(ws3,[50,12,14,65,15,12,28,48]); ws3.freeze_panes='A2'

    # ── TAB 4: Keyword Clusters (NEW) ───────────────────────────────────────
    ws4=wb.create_sheet('Keyword Clusters')
    ws4['A1']='Keyword Clusters — Semantic grouping of all keywords by topic and intent'
    ws4['A1'].font=Font(bold=True,size=13,color='0F6E56'); ws4.merge_cells('A1:L1')

    # Section A: Unmapped keyword clusters (new content needed)
    ws4.cell(row=3,column=1,value='NEW CONTENT NEEDED — Unmapped Keywords Clustered').font=Font(bold=True,size=12,color='0F6E56')
    ws4.merge_cells('A3:L3')
    hdr(ws4,4,['Cluster Name','Entity','Intent Type','Content Type','Suggested Title',
               'Primary Keyword','Primary Volume','Secondary Keywords','Total Volume','Existing URL','Action'],bg='0F6E56')
    r=5
    for cl in sorted(clusters, key=lambda x:-x.get('total_volume',0)):
        ct = cl.get('content_type','Blog post')
        fill = make_fill(C['PU']) if 'Blog' in ct else make_fill(C['RD'])
        vals=[cl.get('cluster_name',''),cl.get('entity',''),cl.get('intent',''),
              ct,cl.get('suggested_title',''),cl.get('primary_keyword',''),
              cl.get('primary_volume',0),' | '.join(cl.get('secondary_keywords',[])),
              cl.get('total_volume',0),'',
              'Create new blog post' if 'Blog' in ct else 'Create new service page']
        for col,v in enumerate(vals,1):
            c=ws4.cell(row=r,column=col,value=v); c.fill=fill; c.font=bf
        r+=1

    # Section B: Mapped URL clusters
    r+=2
    ws4.cell(row=r,column=1,value='EXISTING PAGE OPTIMISATION — Keywords grouped by URL').font=Font(bold=True,size=12,color='0F6E56')
    ws4.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11); r+=1
    hdr(ws4,r,['URL','Intent','Primary Keyword','Primary Volume','Secondary Keywords',
               'Total Volume','# Keywords','Best Score','Content Type','','Action'],bg='1D9E75'); r+=1
    for cl in sorted(url_clusters, key=lambda x:-x['Total Volume']):
        url=cl['URL']; ct=cl['Content Type']
        fill = make_fill(C['BL']) if 'Blog' in ct else make_fill(C['GR'])
        vals=[url,cl['Intent'],cl['Primary Keyword'],cl['Primary Volume'],
              cl['Secondary Keywords'],cl['Total Volume'],cl['Keyword Count'],
              cl['Best Score'],ct,'','Optimise for primary keyword, include secondary keywords naturally']
        for col,v in enumerate(vals,1):
            c=ws4.cell(row=r,column=col,value=v); c.fill=fill; c.font=lf if col==1 else bf
        r+=1

    cw(ws4,[28,15,40,15,65,14,12,12,14,5,45]); ws4.freeze_panes='A5'

    # ── TAB 5: Business Relevant Gaps ───────────────────────────────────────
    ws5=wb.create_sheet('Business Relevant Gaps')
    hdr(ws5,1,['Keyword','Volume','Intent','Service Topic','Relevance','Cluster Name','Action Needed'])
    # Build cluster name lookup
    cluster_lookup = {}
    for cl in clusters:
        for kw2 in [cl.get('primary_keyword','')] + cl.get('secondary_keywords',[]):
            if kw2: cluster_lookup[kw2.lower()] = cl.get('cluster_name','')
    gaps=sorted([r for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in('RELEVANT','BORDERLINE')],key=lambda x:-x['Volume'])
    for r in gaps:
        rel=rel_map.get(r['Keyword'],''); fill=make_fill(C['PU'] if rel=='RELEVANT' else C['AM'])
        action='Create new service page' if r['Intent']=='Transactional' else 'Create new blog post'
        cln=cluster_lookup.get(r['Keyword'].lower(),'')
        rn=ws5.max_row+1
        for col,v in enumerate([r['Keyword'],r['Volume'],r['Intent'],classify_topic(r['Keyword']),rel,cln,action],1):
            c=ws5.cell(row=rn,column=col,value=v); c.fill=fill; c.font=bf
    cw(ws5,[52,12,15,25,14,30,28]); ws5.freeze_panes='A2'

    # ── TAB 6: Priority Roadmap ──────────────────────────────────────────────
    ws6=wb.create_sheet('Priority Roadmap')
    ws6['A1']='SEO Content Opportunity Roadmap'; ws6['A1'].font=Font(bold=True,size=14,color='0F6E56'); ws6.merge_cells('A1:I1')

    qw    = sum(1 for r in mapped if r['Ranking Status']=='Quick win p11-20')
    weak  = sum(1 for r in mapped if r['Ranking Status'] in['Weak ranking p21-50','Very weak p51-100'])
    pnr   = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status']=='Not ranking')
    bgaps = len(gaps)
    conf  = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status']=='Ranking p1-10')
    dup_flag = _find_duplicates(mapped)

    hdr(ws6,3,['Opportunity Type','Count','Total Volume','Action','Priority',''])
    summary=[
        ('Quick Wins (p11-20)',qw,sum(r['Volume'] for r in mapped if r['Ranking Status']=='Quick win p11-20'),'Optimise existing pages','High',C['DGR']),
        ('Weak Rankings (p21-100)',weak,sum(r['Volume'] for r in mapped if r['Ranking Status'] in['Weak ranking p21-50','Very weak p51-100']),'Improve content + internal links','High',C['YL']),
        ('Pages Exist — Not Ranking',pnr,sum(r['Volume'] for r in mapped if r['Landing Page'] and r['Ranking Status']=='Not ranking'),'Optimise existing pages for these keywords','Medium',C['BL']),
        ('Business Relevant Gaps',bgaps,sum(r['Volume'] for r in gaps),'Create new pages / blog posts','Medium',C['PU']),
        ('Already Ranking Well',conf,sum(r['Volume'] for r in mapped if r['Landing Page'] and r['Ranking Status']=='Ranking p1-10'),'Monitor only','Low',C['GY']),
    ]
    for i,(opp,cnt,vol,act,pri,color) in enumerate(summary):
        r=i+4; fill=make_fill(color)
        for col,v in enumerate([opp,cnt,vol,act,pri,''],1):
            c=ws6.cell(row=r,column=col,value=v); c.fill=fill; c.font=Font(size=10,bold=(col==5))

    # FIX 13: Flag duplicate pages
    if dup_flag:
        dr=len(summary)+6
        ws6.cell(row=dr,column=1,value='⚠ DUPLICATE PAGE OPPORTUNITIES — Consider consolidating').font=Font(bold=True,size=11,color='E65100')
        ws6.merge_cells(start_row=dr,start_column=1,end_row=dr,end_column=6); dr+=1
        hdr(ws6,dr,['Service','Page 1','KW Count 1','Page 2','KW Count 2','Recommendation'],bg='E65100'); dr+=1
        for service,p1,c1,p2,c2 in dup_flag:
            for col,v in enumerate([service,p1,c1,p2,c2,'Consider 301 redirecting smaller page to main page'],1):
                c=ws6.cell(row=dr,column=col,value=v); c.fill=make_fill(C['AM']); c.font=bf
            dr+=1

    # Detail section
    dr2=ws6.max_row+3
    ws6.cell(row=dr2,column=1,value='Detailed Action List — sorted by priority then volume').font=Font(bold=True,size=12,color='0F6E56')
    ws6.merge_cells(start_row=dr2,start_column=1,end_row=dr2,end_column=9); dr2+=1
    hdr(ws6,dr2,['Keyword','Volume','Your Position','Intent','Final Score','Opportunity','Action','Priority','Cluster']); dr2+=1

    PORD={'Quick win — optimise':1,'Weak ranking':2,'Page exists — optimise':3,'Blog exists — optimise':3,
          'Business relevant gap':4,'True content gap':5,'Confirmed existing page':6}
    AFILL={'Quick win — optimise':make_fill(C['DGR']),'Weak ranking':make_fill(C['YL']),
           'Page exists — optimise':make_fill(C['BL']),'Blog exists — optimise':make_fill(C['OR']),
           'Business relevant gap':make_fill(C['PU']),'True content gap':make_fill(C['RD']),
           'Confirmed existing page':make_fill(C['GY'])}

    all_items=[]
    for r in mapped:
        rel=rel_map.get(r['Keyword'],''); url=r['Landing Page']; rs=r['Ranking Status']; fs=r['Final Score']; src=r.get('Match Source','')
        if url and rs=='Ranking p1-10': opp='Confirmed existing page'
        elif rs=='Quick win p11-20': opp='Quick win — optimise'
        elif rs in['Weak ranking p21-50','Very weak p51-100']: opp='Weak ranking'
        elif url and 'Claude semantic' in src: opp='Blog exists — optimise'
        elif url: opp='Page exists — optimise'
        elif rel in('RELEVANT','BORDERLINE'): opp='Business relevant gap'
        else: opp='True content gap'
        act={'Confirmed existing page':'Monitor','Quick win — optimise':'Optimise title + meta + H1',
             'Weak ranking':'Improve content + internal links',
             'Page exists — optimise':'Optimise existing page',
             'Blog exists — optimise':'Update blog title/meta/H1',
             'Business relevant gap':'Create new service page' if r['Intent']=='Transactional' else 'Create new blog post',
             'True content gap':'Evaluate for new page'}.get(opp,'')
        pri={1:'High',2:'High',3:'Medium',4:'Medium',5:'Low',6:'Low'}.get(PORD.get(opp,5),'Low')
        cln=cluster_lookup.get(r['Keyword'].lower(),'') if 'cluster_lookup' in dir() else ''
        all_items.append({**r,'opp':opp,'act':act,'pri':pri,'po':PORD.get(opp,5),'cln':cln})

    all_items.sort(key=lambda x:(x['po'],-x['Volume']))
    for item in all_items:
        fill=AFILL.get(item['opp'],make_fill(C['WH']))
        for col,v in enumerate([item['Keyword'],item['Volume'],item['Your Position'],item['Intent'],
                                 item['Final Score'],item['opp'],item['act'],item['pri'],item['cln']],1):
            c=ws6.cell(row=dr2,column=col,value=v); c.fill=fill; c.font=Font(size=10,bold=(col==8))
        dr2+=1

    cw(ws6,[50,12,14,15,12,28,42,10,28]); ws6.freeze_panes=f'A{ws6.max_row-len(all_items)+2}'

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def _find_duplicates(mapped):
    """FIX 13: Detect pages covering the same service topic."""
    url_counts={}
    for r in mapped:
        if r['Landing Page']:
            url=r['Landing Page']; url_counts[url]=url_counts.get(url,0)+1
    service_cats={}
    for url,cnt in url_counts.items():
        slug=re.sub(r'https?://[^/]+','',url).strip('/')
        cats=cat_url(url)
        for cat in cats:
            if cat not in ('blog','location','plumbing_general','hvac_general') and not cat.startswith('loc:') and not cat.startswith('brand:'):
                if cat not in service_cats: service_cats[cat]=[]
                service_cats[cat].append((url,cnt))
    dups=[]
    for cat,pages in service_cats.items():
        if len(pages)>=2:
            pages_sorted=sorted(pages,key=lambda x:-x[1])
            for j in range(1,min(3,len(pages_sorted))):
                p1,c1=pages_sorted[0]; p2,c2=pages_sorted[j]
                if c2>=5:
                    s1=re.sub(r'https?://[^/]+','',p1); s2=re.sub(r'https?://[^/]+','',p2)
                    if s1!=s2 and (s1,s2) not in [(d[1],d[3]) for d in dups]:
                        dups.append((cat.replace('_',' ').title(),s1,c1,s2,c2))
    return dups[:10]

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.title("🔍 SEO Content Mapper")
st.markdown("Complete SEO gap analysis — GSC validation + keyword mapping + semantic clustering.")

with st.sidebar:
    st.header("⚙️ Settings"); st.markdown("---")
    threshold=st.slider("Match Score Threshold",0.10,0.40,0.15,0.01)
    st.markdown("**Content Weights**")
    sw=st.slider("URL Slug",1,8,5); tw=st.slider("Page Title",1,6,3)
    hw=st.slider("H1 Heading",1,4,2); mw=st.slider("Meta Description",1,3,1)
    weights=(sw,tw,hw,mw)
    st.markdown("---"); st.markdown("**Score Guide**")
    st.markdown("| Score | Quality |\n|---|---|\n| 0.50+ | Strong ✅ |\n| 0.30–0.50 | Good ✅ |\n| 0.20–0.30 | Acceptable ⚠️ |\n| 0.15–0.20 | Weak — verify |\n| Blank | Content gap |")
    st.markdown("---"); st.markdown("**Colour Coding**")
    st.markdown("🟢 Content match  🔵 Blog  🟡 GSC fallback  🟠 Claude semantic  🟣 Business gap")

# FILE 1 — Screaming Frog
st.markdown('<div class="sec-hdr">📂 File 1 — Screaming Frog Export</div>', unsafe_allow_html=True)
st.caption("Screaming Frog → Bulk Export → All. Required columns: Address, Title 1, Meta Description 1, H1-1")
sf_file=st.file_uploader("Upload Screaming Frog file (.xlsx or .csv)",type=['xlsx','csv'],key='sf')
sf_df=None
if sf_file:
    try:
        sf_df=pd.read_csv(sf_file) if sf_file.name.endswith('.csv') else pd.read_excel(sf_file)
        sf_df.columns=[c.strip() for c in sf_df.columns]
        missing=[c for c in ['Address','Title 1','Meta Description 1','H1-1'] if c not in sf_df.columns]
        if missing:
            st.markdown(f'<div class="warn-box">⚠️ Missing columns: {missing}. Found: {list(sf_df.columns[:8])}</div>',unsafe_allow_html=True); sf_df=None
        else:
            if 'Status Code' in sf_df.columns: sf_df=sf_df[sf_df['Status Code']==200].reset_index(drop=True)
            sf_df=sf_df[sf_df['Address'].notna()].reset_index(drop=True)
            st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(sf_df):,}</strong> pages</div>',unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>',unsafe_allow_html=True)

# FILE 2 — GSC
st.markdown('<div class="sec-hdr">📂 File 2 — GSC Export</div>', unsafe_allow_html=True)
st.caption("Looker Studio export. Required: Query, Landing Page, Clicks, Impressions, Position, CTR")
gsc_file=st.file_uploader("Upload GSC Excel file",type=['xlsx'],key='gsc')
gsc_df=None
if gsc_file:
    try:
        gsc_xl=pd.read_excel(gsc_file,sheet_name=None); gsc_sheets=list(gsc_xl.keys())
        gsc_sheet=st.selectbox("Select GSC sheet",gsc_sheets,key='gs')
        raw=gsc_xl[gsc_sheet]; cols=list(raw.columns)
        with st.expander("Map GSC columns"):
            qc  =st.selectbox("Query",cols,index=cols.index('Query') if 'Query' in cols else 0,key='qc')
            lpc =st.selectbox("Landing Page",cols,index=cols.index('Landing Page') if 'Landing Page' in cols else 1,key='lpc')
            clkc=st.selectbox("Clicks",cols,index=next((i for i,c in enumerate(cols) if 'click' in c.lower()),2),key='clkc')
            impc=st.selectbox("Impressions",cols,index=next((i for i,c in enumerate(cols) if 'impression' in c.lower()),3),key='impc')
            posc=st.selectbox("Position",cols,index=next((i for i,c in enumerate(cols) if 'position' in c.lower()),min(4,len(cols)-1)),key='posc')
            ctrc=st.selectbox("CTR",cols,index=next((i for i,c in enumerate(cols) if 'ctr' in c.lower()),min(5,len(cols)-1)),key='ctrc')
        gsc_df=raw.rename(columns={qc:'Query',lpc:'Landing Page',clkc:'Clicks',impc:'Impressions',posc:'Position',ctrc:'CTR'})
        gsc_df['Clicks']=pd.to_numeric(gsc_df['Clicks'],errors='coerce').fillna(0)
        gsc_df['Impressions']=pd.to_numeric(gsc_df['Impressions'],errors='coerce').fillna(0)
        st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(gsc_df):,}</strong> GSC rows</div>',unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>',unsafe_allow_html=True)

# FILE 3 — Semrush
st.markdown('<div class="sec-hdr">📂 File 3 — Semrush Keyword Gap Export</div>', unsafe_allow_html=True)
st.caption("Full keyword gap export. Required: Keyword, Volume, your domain position column.")
sem_file=st.file_uploader("Upload Semrush file",type=['xlsx'],key='sem')
sem_df=None; your_col=None
if sem_file:
    try:
        sem_xl=pd.read_excel(sem_file,sheet_name=None); sem_sheets=list(sem_xl.keys())
        sem_sheet=st.selectbox("Select Semrush sheet",sem_sheets,key='ss')
        raw=sem_xl[sem_sheet]; cols=list(raw.columns)
        with st.expander("Map Semrush columns"):
            kwc =st.selectbox("Keyword",cols,index=cols.index('Keyword') if 'Keyword' in cols else 0,key='skc')
            vc  =st.selectbox("Volume",cols,index=cols.index('Volume') if 'Volume' in cols else 1,key='svc')
            your_col=st.selectbox("Your domain position column",cols,index=0,help="Column showing YOUR site's position",key='sdc')
        sem_df=raw.rename(columns={kwc:'Keyword',vc:'Volume'})
        sem_df=sem_df[sem_df['Keyword'].notna()].reset_index(drop=True)
        st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(sem_df):,}</strong> keywords</div>',unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>',unsafe_allow_html=True)

# BUSINESS CONTEXT
st.markdown('<div class="sec-hdr">🏢 Business Context + Claude API</div>', unsafe_allow_html=True)
ca,cb=st.columns(2)
with ca:
    api_key=st.text_input("Claude API Key",type="password",placeholder="sk-ant-...",help="Required. Get from console.anthropic.com")
    biz_desc=st.text_area("Business Description *",placeholder="Describe your business, services offered and location...",height=130)
with cb:
    excl_str=st.text_area("Services NOT offered (optional — one per line)",placeholder="e.g.\noil boiler\nseptic tank\nwell pump",height=130)
    st.caption("Leave blank if unsure — Claude will use the business description to judge relevance.")

# RUN
st.markdown("---")
issues=[]
if sf_df is None:        issues.append("File 1 (Screaming Frog) not uploaded or has errors")
if gsc_df is None:       issues.append("File 2 (GSC) not uploaded or has errors")
if sem_df is None:       issues.append("File 3 (Semrush) not uploaded or has errors")
if not api_key.strip():  issues.append("Claude API key is required")
if not biz_desc.strip(): issues.append("Business description is required")

for issue in issues:
    st.markdown(f'<div class="warn-box">⚠️ {issue}</div>',unsafe_allow_html=True)

if st.button("🚀 Run Full Analysis",disabled=bool(issues),use_container_width=True,type="primary"):
    progress_bar=st.progress(0); status_text=st.empty()
    try:
        # Phase 1
        gsc_val, url_df, gsc_dedup, stop = phase1_gsc(gsc_df, sf_df, weights, status_text, progress_bar)

        # Phase 2
        mapped = phase2_semrush(sem_df, url_df, gsc_dedup, gsc_val, weights, threshold,
                                 your_col, status_text, progress_bar)

        # Phase 3
        rel_map, mapped = phase3_claude(mapped, url_df, api_key, biz_desc, excl_str,
                                         status_text, progress_bar)

        # Phase 4
        clusters, url_clusters = phase4_cluster(mapped, rel_map, api_key, status_text, progress_bar)

        # Build cluster lookup for roadmap
        cluster_lookup = {}
        for cl in clusters:
            for kw2 in [cl.get('primary_keyword','')] + cl.get('secondary_keywords',[]):
                if kw2: cluster_lookup[kw2.lower()] = cl.get('cluster_name','')

        progress_bar.progress(99); status_text.text("Building Excel output...")
        excel_buf = build_excel(gsc_val, mapped, rel_map, clusters, url_clusters)
        progress_bar.progress(100); status_text.text("Done!")

        # Metrics
        st.markdown("---"); st.subheader("📊 Analysis Complete")
        conf_gsc  = len(gsc_val[gsc_val['Mapping Status']=='Confirmed'])
        susp_gsc  = len(gsc_val[gsc_val['Mapping Status']=='Suspicious'])
        mapped_kw = sum(1 for r in mapped if r['Landing Page'])
        qw        = sum(1 for r in mapped if r['Ranking Status']=='Quick win p11-20')
        biz_gaps  = sum(1 for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in('RELEVANT','BORDERLINE'))
        n_clusters= len(clusters)

        c1,c2,c3,c4,c5,c6=st.columns(6)
        for col,num,label in [(c1,len(gsc_val),'GSC queries'),(c2,conf_gsc,'GSC confirmed'),
                               (c3,susp_gsc,'GSC suspicious'),(c4,mapped_kw,'Keywords mapped'),
                               (c5,biz_gaps,'Business gaps'),(c6,n_clusters,'Topic clusters')]:
            col.markdown(f'<div class="metric-card"><div class="metric-num">{num:,}</div><div class="metric-label">{label}</div></div>',unsafe_allow_html=True)

        st.markdown("")
        t1,t2,t3,t4=st.tabs(["GSC Validation","Keyword Mapping","Topic Clusters","Business Gaps"])
        with t1: st.dataframe(gsc_val.head(200),use_container_width=True,height=350)
        with t2:
            prev=pd.DataFrame([{'Keyword':r['Keyword'],'Volume':r['Volume'],'Your Position':r['Your Position'],
                                 'Landing Page':r['Landing Page'],'Final Score':r['Final Score'],
                                 'Source':r['Match Source'],'Ranking':r['Ranking Status']} for r in mapped]).sort_values('Final Score',ascending=False)
            st.dataframe(prev.head(200),use_container_width=True,height=350)
        with t3:
            if clusters:
                cl_prev=pd.DataFrame([{'Cluster':c.get('cluster_name',''),'Entity':c.get('entity',''),
                                        'Intent':c.get('intent',''),'Type':c.get('content_type',''),
                                        'Suggested Title':c.get('suggested_title',''),
                                        'Primary KW':c.get('primary_keyword',''),
                                        'Total Volume':c.get('total_volume',0)} for c in clusters]).sort_values('Total Volume',ascending=False)
                st.dataframe(cl_prev,use_container_width=True,height=350)
            else:
                st.info("No clusters generated.")
        with t4:
            gp=pd.DataFrame([{'Keyword':r['Keyword'],'Volume':r['Volume'],'Intent':r['Intent'],
                               'Relevance':rel_map.get(r['Keyword'],''),'Topic':classify_topic(r['Keyword'])} for r in mapped
                              if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in('RELEVANT','BORDERLINE')]).sort_values('Volume',ascending=False)
            st.dataframe(gp.head(200),use_container_width=True,height=350)

        st.markdown("---")
        st.download_button("⬇️ Download Full Excel Output (6 tabs)",data=excel_buf,
                           file_name="seo_content_mapping_output.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True,type="primary")
        st.markdown('<div class="info-box"><strong>6 tabs:</strong> GSC Mapping Quality | Keyword Mapping | Opportunity Classification | Keyword Clusters | Business Relevant Gaps | Priority Roadmap</div>',unsafe_allow_html=True)

    except Exception as e:
        progress_bar.progress(0)
        st.markdown(f'<div class="error-box">❌ Error: {str(e)}</div>',unsafe_allow_html=True)
        st.exception(e)
