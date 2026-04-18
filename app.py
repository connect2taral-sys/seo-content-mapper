import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import json
import time
import urllib.request
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="SEO Content Mapper", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.main{padding-top:1rem;}
.stProgress>div>div>div{background-color:#1D9E75;}
.metric-card{background:#f0faf5;border:1px solid #1D9E75;border-radius:8px;padding:16px;text-align:center;}
.metric-num{font-size:2rem;font-weight:700;color:#1D9E75;}
.metric-label{font-size:0.85rem;color:#555;margin-top:4px;}
.info-box{background:#e8f5f0;border-left:4px solid #1D9E75;padding:12px 16px;border-radius:4px;margin:8px 0;}
.warn-box{background:#fffbea;border-left:4px solid #f59e0b;padding:12px 16px;border-radius:4px;margin:8px 0;}
.error-box{background:#fef2f2;border-left:4px solid #ef4444;padding:12px 16px;border-radius:4px;margin:8px 0;}
.section-header{background:#1D9E75;color:white;padding:8px 16px;border-radius:4px;font-weight:600;margin:16px 0 8px 0;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAXONOMY
# ─────────────────────────────────────────────────────────────────────────────
HVAC_GROUP       = {'furnace','boiler','heat_pump','ac_cooling','ductless','hvac_general','geothermal'}
PLUMBING_GROUP   = {'water_pipes','drain_sewer','water_heater','toilet','faucet_sink','shower_tub','bathroom','kitchen','plumbing_general','backflow','sump_pump','sprinkler'}
ELECTRICAL_GROUP = {'electrical','generator','ev_charger'}
ALL_EXCLUSIVE    = HVAC_GROUP | PLUMBING_GROUP | ELECTRICAL_GROUP | {'gas_utility','air_quality'}

EXCLUDE_PATTERNS = ['/truck-wrap','/careers','/join-the-team','/contact','/podcast','/coupons','/warranty','/about-us','/meet-the','/technicians','/community','/financing','/solar','/estore','/home-energy','/lancaster-neighbors','/tonawanda-dispatch','/hamburg-service','/buffalo-service','/buffalo-home-improvement','/case_study_category','why-choose-us','/walker','/our-community','/renovation-services','/case-stud','/specials']

EXPAND_RULES = [(r'\bac\b','air conditioning'),(r'\ba/c\b','air conditioning'),(r'\bhvac\b','heating cooling air conditioning'),(r'\bfurnace\b','furnace heating'),(r'\bheat pump\b','heat pump heating cooling'),(r'\bboiler\b','boiler heating'),(r'\bgenerator\b','generator backup power'),(r'\bdrain\b','drain sewer'),(r'\bsump pump\b','sump pump basement'),(r'\bgfci\b','gfci outlet electrical'),(r'\bbackflow\b','backflow prevention'),(r'\buv\b','uv air sanitizer'),(r'\brepiping\b','repiping pipe replacement'),(r'\bwater heater\b','water heater hot water'),(r'\btankless\b','tankless water heater'),(r'\bmini.?split\b','mini split ductless heating cooling')]

INFO_SIGNALS = [r'^what\b',r'^how\b',r'^why\b',r'^when\b',r'^does\b',r'^do\b',r'^is\b',r'^are\b',r'^can\b',r'^should\b',r'^which\b',r'^who\b',r'\bvs\b',r'\bversus\b',r'\bdifference between\b',r'\btips\b',r'\bbenefits of\b',r'\bsigns\b',r'\bcauses\b',r'\btypes of\b',r'\bhow to\b',r'\bwhat is\b',r'\bwhy is\b',r'\bhow does\b',r'\bwhat does\b',r'\badvantages\b',r'\bproblems with\b']

TOPIC_CLUSTERS = {'Furnace / Heating':['furnace heating','furnace','heating','heat','boiler heating','boiler','steam heat','radiant'],'AC / Cooling':['air conditioning','cooling','central air','mini split ductless','ductless'],'Heat Pump':['heat pump heating cooling'],'Water Heater':['water heater hot water','water heater','tankless water heater','hot water'],'Drain / Sewer':['drain sewer','sewer','clog','unclog','jetting','hydrojet','rooter'],'Plumbing General':['plumb','pipe','leak','repiping','water line','water main'],'Electrical':['electric','electrician','panel','wiring','outlet','gfci outlet electrical','circuit','lighting','generator backup power','surge'],'HVAC General':['heating cooling air conditioning'],'Sump Pump':['sump pump basement'],'Air Quality':['air quality','dehumidif','humidif','air purif','uv air sanitizer','indoor air'],'Backflow':['backflow prevention'],'Geothermal':['geothermal'],'Bathroom / Kitchen':['bathroom','kitchen','toilet','shower','faucet','sink','tub','vanity','remodel','renovation'],'Commercial':['commercial'],'Location-Specific':['cheektowaga','amherst','hamburg','lancaster','west seneca','orchard park','east aurora','clarence','depew','tonawanda','williamsville','batavia','grand island','rochester','syracuse','albany']}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def expand(t):
    t = t.lower()
    for p,r in EXPAND_RULES: t = re.sub(p,r,t)
    return t

def is_excluded(url):
    return any(p in str(url).lower() for p in EXCLUDE_PATTERNS)

def is_informational(kw):
    return any(re.search(p,kw.lower()) for p in INFO_SIGNALS)

def slug_terms(url, stop=None):
    base_stop = {'com','https','http','in','and','the','a','of','for','to','near','me','www'}
    if stop: base_stop |= stop
    path = re.sub(r'https?://[^/]+','',str(url))
    parts = re.split(r'[/\-_]',path)
    return ' '.join(p for p in parts if p and p.lower() not in base_stop)

def build_content(row, weights, stop=None):
    url   = str(row.get('Address', row.get('Landing Page','')))
    slug  = expand(slug_terms(url, stop))
    title = expand(str(row.get('Title 1', row.get('Page Title','')) or '').replace('| Cellino Plumbing','').replace('- Cellino Plumbing',''))
    meta  = expand(str(row.get('Meta Description 1', row.get('Meta Description','')) or ''))
    h1r   = str(row.get('H1-1', row.get('H1','')) or '')
    h1    = '' if h1r.lower() in ('nan','none','') else expand(h1r)
    sw,tw,hw,mw = weights
    return f"{slug} "*sw + f"{title} "*tw + f"{h1} "*hw + meta*mw

def get_url(row):
    return str(row.get('Address', row.get('Landing Page','')))

def categorize_url(url):
    p = re.sub(r'https?://[^/]+','',url.lower())
    cats = set()
    if any(t in p for t in ['gas-leak','gas-line','gas-pipe','gas-service','gas-detection']): cats.add('gas_utility')
    if any(t in p for t in ['water-leak','water-line','water-main','main-water','burst-pipe','slab-leak','repiping','frozen-pipe','underground','water-shut']): cats.add('water_pipes')
    if any(t in p for t in ['drain','sewer','hydro','jetting','rooter','bubbler','clogged-kitchen','flooded-basement']): cats.add('drain_sewer')
    if any(t in p for t in ['water-heater','tankless']): cats.add('water_heater')
    if any(t in p for t in ['furnace','heater-install']): cats.add('furnace')
    if p.rstrip('/') == '/heating': cats.update({'furnace','boiler','heat_pump','hvac_general'})
    if any(t in p for t in ['boiler','steam-boiler']): cats.add('boiler')
    if 'heat-pump' in p: cats.add('heat_pump')
    if any(t in p for t in ['ac-tune','ac-filter','ac-install','evaporator','buffalo-repair','68-ac']): cats.add('ac_cooling')
    if p.rstrip('/') == '/air-conditioning': cats.update({'ac_cooling','hvac_general'})
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
    if p.rstrip('/') in ['/buffalo-ny','/maintenance-plans','/buffalo']: cats.update({'plumbing_general','hvac_general','electrical'})
    for loc in ['cheektowaga','amherst','hamburg','lancaster','west-seneca','orchard-park','east-aurora','springville','alden','akron','batavia','clarence','depew','elma','eden','marilla','holland','wales','darien','medina','elba','attica','boston-ny','grand-island','bennington','west-valley','williamsville','tonawanda']:
        if loc in p: cats.add('location'); cats.add('location:'+loc.replace('-',' ')); break
    if '/blog/' in p: cats.add('blog')
    if not (cats - {'blog'}): cats.add('plumbing_general')
    return frozenset(cats)

def categorize_kw(kw):
    kw_l = expand(kw.lower())
    cats = set()
    GAS = ['gas leak','gas line','gas pipe','natural gas line','gas service line','gas detector','gas fireplace','gas stove','gas dryer','gas appliance','gas connection','gas hookup','gas meter','carbon monoxide detector','gas shut off','propane line','natural gas leak']
    if any(t in kw_l for t in GAS):
        if not any(x in kw_l for x in ['furnace heating','boiler heating','water heater hot water','heater']): cats.add('gas_utility')
    WH = ['water heater hot water','water heater','hot water heater','tankless water heater','hot water tank','leaking water heater','water heater leak','water heater drip','water heater temp','water heater not','no hot water','water not heating']
    if any(t in kw_l for t in WH): cats.add('water_heater')
    if 'water_heater' not in cats:
        WP = ['water leak','water line','water main','water pipe','burst pipe','pipe leak','slab leak','repiping','main water','frozen pipe','water service line','water shut off valve','main line leak','pipe burst','whole home repipe','water pressure']
        if any(t in kw_l for t in WP): cats.add('water_pipes')
    DR = ['drain sewer','sewer','unclog','hydro jet','hydrojett','rooter','sewage','septic','main line clog','main sewer','main drain sewer','backed up drain','slow drain','floor drain','bubbler','storm drain','overflow drain','exterior drain','interior drain','basement drain','laundry drain','bathroom drain','kitchen drain']
    if any(t in kw_l for t in DR): cats.add('drain_sewer')
    FU = ['furnace heating','furnace tune','furnace service','furnace repair','furnace install','furnace maint','furnace clean','furnace filter','furnace not working','furnace not heating','annual furnace','gas furnace','electric furnace','new furnace','furnace replace','forced air']
    if any(t in kw_l for t in FU): cats.add('furnace')
    BO = ['boiler heating','boiler repair','boiler service','boiler install','boiler maint','boiler replace','boiler tune','boiler cleaning','annual boiler','gas boiler','oil boiler','steam heat','steam boiler','radiant heat']
    if any(t in kw_l for t in BO): cats.add('boiler')
    if 'heat pump heating cooling' in kw_l: cats.add('heat_pump')
    AC = ['air conditioning','air conditioner','central air','cooling system','evaporator coil','refrigerant','freon','cooling near me']
    if any(t in kw_l for t in AC): cats.add('ac_cooling')
    if 'mini split ductless' in kw_l: cats.update({'ductless','ac_cooling','heat_pump'})
    if 'heating cooling air conditioning' in kw_l: cats.update({'hvac_general','furnace','ac_cooling'})
    elif 'heating cooling' in kw_l: cats.update({'hvac_general','furnace','ac_cooling'})
    if 'geothermal' in kw_l: cats.update({'geothermal','hvac_general'})
    is_combo = any(x in kw_l for x in ['water heater','furnace heating','boiler heating','heat pump heating'])
    EL = ['electrician','electric panel','breaker box','circuit breaker','electrical panel','electrical wire','wiring','rewiring','outlet','gfci outlet electrical','light fixture','lighting install','lighting repair','surge protect','electrical inspect','electrical troubl','electrical service','electrical repair','electrical install','electrical work','electrical contractor','electric repair','electric install','residential electric','outdoor lighting','electrical permit','electrical company','local electrician','licensed electrician','home electrical']
    if any(t in kw_l for t in EL) and not is_combo: cats.add('electrical')
    if 'generator backup power' in kw_l: cats.update({'generator','electrical'})
    if any(t in kw_l for t in ['ev charger','electric vehicle charg','charging station','car charging','home charging']): cats.update({'ev_charger','electrical'})
    if 'sump pump basement' in kw_l: cats.add('sump_pump')
    AQ = ['air quality','dehumidif','humidif','air purif','air filter service','uv air sanitizer','indoor air','air cleaner','air scrubber']
    if any(t in kw_l for t in AQ): cats.add('air_quality')
    if 'backflow prevention' in kw_l: cats.add('backflow')
    if 'toilet' in kw_l: cats.update({'toilet','drain_sewer'})
    if any(t in kw_l for t in ['faucet','garbage disposal','kitchen sink','disposal unit']): cats.update({'faucet_sink','plumbing_general'})
    if any(t in kw_l for t in ['shower','bathtub','bath tub']): cats.update({'shower_tub','drain_sewer','plumbing_general'})
    if any(t in kw_l for t in ['bathroom renov','bathroom remodel','bathroom addition','bathroom plumb']): cats.update({'bathroom','plumbing_general'})
    if any(t in kw_l for t in ['kitchen renov','kitchen remodel','kitchen plumb']): cats.update({'kitchen','plumbing_general'})
    if 'sprinkler' in kw_l: cats.add('sprinkler')
    PL = ['plumbing','plumber','pipe install','pipe replac','pipe repair','pipe clean','residential plumb','commercial plumb','plumbing inspect','emergency plumb','24 hour plumb','plumbing company','plumbing service','local plumb','home plumb','plumb contractor','licensed plumb','best plumb','professional plumb']
    if any(t in kw_l for t in PL): cats.add('plumbing_general')
    if 'heater' in kw_l and 'water heater' not in kw_l and 'water_heater' not in cats: cats.update({'furnace','hvac_general'})
    if re.search(r'\bheating\b',kw_l) and not (cats & {'furnace','boiler','heat_pump','hvac_general','water_heater'}): cats.update({'furnace','hvac_general'})
    if re.search(r'\bcooling\b',kw_l) and not (cats & {'ac_cooling','ductless','heat_pump','hvac_general'}): cats.update({'ac_cooling','hvac_general'})
    return frozenset(cats)

def is_compatible(kw_cats, url_cats, kw, url):
    if 'location' in url_cats:
        kw_l = expand(kw.lower())
        loc_tags = [t for t in url_cats if t.startswith('location:')]
        return bool(loc_tags) and loc_tags[0].replace('location:','') in kw_l
    if 'gas_utility' in url_cats and 'gas_utility' not in kw_cats: return False
    if 'gas_utility' in kw_cats and 'gas_utility' not in url_cats: return False
    kw_ex = kw_cats & ALL_EXCLUSIVE; url_ex = url_cats & ALL_EXCLUSIVE
    if not kw_ex or not url_ex: return True
    if kw_ex & url_ex: return True
    kw_sh  = kw_cats  & (HVAC_GROUP-{'hvac_general'}); url_sh = url_cats & (HVAC_GROUP-{'hvac_general'})
    if 'hvac_general' in kw_cats  and not kw_sh  and url_ex & HVAC_GROUP: return True
    if 'hvac_general' in url_cats and not url_sh  and kw_ex  & HVAC_GROUP: return True
    kw_sp  = kw_cats  & (PLUMBING_GROUP-{'plumbing_general'}); url_sp = url_cats & (PLUMBING_GROUP-{'plumbing_general'})
    if 'plumbing_general' in kw_cats  and not kw_sp  and url_ex & PLUMBING_GROUP: return True
    if 'plumbing_general' in url_cats and not url_sp  and kw_ex  & PLUMBING_GROUP: return True
    if 'electrical' in kw_cats  and url_ex & ELECTRICAL_GROUP: return True
    if 'electrical' in url_cats and kw_ex  & ELECTRICAL_GROUP: return True
    return False

def clean_kw(kw):
    kw = expand(kw.lower().strip())
    for p in [r'\bnear me\b',r'\bnear\b']: kw = re.sub(p,'',kw)
    return re.sub(r'\s+',' ',kw).strip()

def classify_topic(kw):
    kw_e = expand(kw.lower())
    for t,terms in TOPIC_CLUSTERS.items():
        if any(x in kw_e for x in terms): return t
    return 'Other'

def make_fill(c): return PatternFill(start_color=c,end_color=c,fill_type='solid')
def hdr(ws,row,cols,bg='1D9E75',fg='FFFFFF'):
    for col,val in enumerate(cols,1):
        c=ws.cell(row=row,column=col,value=val)
        c.fill=make_fill(bg); c.font=Font(bold=True,color=fg,size=11); c.alignment=Alignment(horizontal='left')
def col_w(ws,widths):
    for col,w in zip('ABCDEFGHIJKLMNOP',widths): ws.column_dimensions[col].width=w

# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE API
# ─────────────────────────────────────────────────────────────────────────────
def call_claude(api_key, prompt, max_tokens=1500):
    payload = json.dumps({"model":"claude-sonnet-4-20250514","max_tokens":max_tokens,"messages":[{"role":"user","content":prompt}]}).encode()
    req = urllib.request.Request("https://api.anthropic.com/v1/messages",data=payload,
        headers={"Content-Type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"},method="POST")
    with urllib.request.urlopen(req,timeout=60) as resp:
        data = json.loads(resp.read())
    text = data['content'][0]['text'].strip()
    text = re.sub(r'^```json?\s*','',text); text = re.sub(r'\s*```$','',text)
    return json.loads(text.strip())

def claude_relevance(api_key, keywords, biz_desc, exclusions, status_text, progress_bar, p0, p1):
    excl = f"\nServices NOT offered: {exclusions}" if exclusions.strip() else ""
    results = {}
    batches = [keywords[i:i+50] for i in range(0,len(keywords),50)]
    for b_idx,batch in enumerate(batches):
        pct = p0+int((b_idx/len(batches))*(p1-p0))
        progress_bar.progress(min(pct,p1-1))
        status_text.text(f"Claude API — relevance check: batch {b_idx+1}/{len(batches)} ({len(results):,} done)...")
        prompt = f"""You are an SEO analyst. Classify each keyword as relevant to this business.

Business: {biz_desc}{excl}

Return ONE label per keyword:
- RELEVANT: clearly relates to a service this business offers
- NOT_RELEVANT: unrelated to this business  
- BORDERLINE: loosely related, worth reviewing

Return ONLY valid JSON (no explanation):
{{"keyword":"RELEVANT/NOT_RELEVANT/BORDERLINE",...}}

Keywords:
{json.dumps(batch)}"""
        for attempt in range(3):
            try: batch_res = call_claude(api_key,prompt); results.update(batch_res); break
            except: 
                if attempt<2: time.sleep(3)
                else:
                    for kw in batch: results[kw]='BORDERLINE'
        time.sleep(0.3)
    return results

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 1 — VALIDATE GSC MAPPINGS
# ─────────────────────────────────────────────────────────────────────────────
def phase1_validate_gsc(gsc_df, sf_df, weights, status_text, progress_bar):
    status_text.text("Phase 1: Preparing URL index from Screaming Frog data...")
    progress_bar.progress(5)

    # Extract domain stop words
    stop = set()
    if len(sf_df):
        sample = str(sf_df.iloc[0].get('Address',''))
        m = re.search(r'https?://([^/]+)',sample)
        if m: stop = set(m.group(1).replace('www.','').split('.'))

    # Clean URL df
    url_df = sf_df[~sf_df.apply(lambda r: is_excluded(get_url(r)),axis=1)].reset_index(drop=True)
    url_df['content'] = url_df.apply(lambda r: build_content(r,weights,stop),axis=1)
    url_df['cats']    = url_df.apply(lambda r: categorize_url(get_url(r)),axis=1)
    url_df['ptype']   = url_df.apply(lambda r: 'blog' if '/blog/' in get_url(r).lower() else 'service',axis=1)
    url_content_map   = {get_url(r).lower().strip(): url_df.at[i,'content'] for i,r in url_df.iterrows()}

    # Deduplicate GSC
    gsc_df['kw_lower'] = gsc_df['Query'].str.lower().str.strip()
    gsc_dedup = {}
    for kw,grp in gsc_df.groupby('kw_lower'):
        best = grp.sort_values(['Clicks','Impressions'],ascending=False).iloc[0]
        gsc_dedup[kw] = {'url':best['Landing Page'],'clicks':best['Clicks'],'impressions':best['Impressions'],'position':best.get('Position',0),'ctr':best.get('CTR',0)}

    progress_bar.progress(10)
    status_text.text("Phase 1: Scoring GSC query vs page content...")

    queries = list(gsc_dedup.keys())
    cleaned_q = [clean_kw(q) for q in queries]
    all_contents = list(url_content_map.values())
    url_keys = list(url_content_map.keys())

    vec = TfidfVectorizer(ngram_range=(1,3),min_df=1,sublinear_tf=True)
    vec.fit(all_contents+cleaned_q)
    url_vecs   = vec.transform(all_contents)
    q_vecs     = vec.transform(cleaned_q)
    sims       = cosine_similarity(q_vecs,url_vecs)
    uk_idx     = {k:i for i,k in enumerate(url_keys)}

    validated = []
    for i,(q,cq) in enumerate(zip(queries,cleaned_q)):
        if i%1000==0:
            progress_bar.progress(min(10+int((i/len(queries))*15),24))
            status_text.text(f"Phase 1: Validating {i:,}/{len(queries):,} GSC queries...")
        d = gsc_dedup[q]
        gsc_url = str(d['url']).lower().strip()
        idx = uk_idx.get(gsc_url)
        score = round(float(sims[i,idx]),4) if idx is not None else 0.0
        status = 'Confirmed' if score>=0.30 else ('Plausible' if score>=0.15 else 'Suspicious')
        validated.append({'Query':q,'Mapped URL':d['url'],'Clicks':d['clicks'],'Impressions':d['impressions'],'Position':round(float(d['position']),1) if d['position'] else '','CTR':d['ctr'],'Content Score':score,'Mapping Status':status})

    progress_bar.progress(25)
    return pd.DataFrame(validated), url_df, gsc_dedup, stop

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 2 — MAP SEMRUSH KEYWORDS
# ─────────────────────────────────────────────────────────────────────────────
def phase2_map_keywords(sem_df, url_df, gsc_dedup, weights, threshold, your_col, status_text, progress_bar):
    status_text.text("Phase 2: Building TF-IDF models...")
    progress_bar.progress(27)

    service_df = url_df[url_df['ptype']=='service'].reset_index(drop=True)
    blog_df    = url_df[url_df['ptype']=='blog'].reset_index(drop=True)
    keywords   = sem_df['Keyword'].fillna('').tolist()
    volumes    = sem_df['Volume'].fillna(0).tolist()
    your_pos   = sem_df[your_col].fillna('').tolist() if your_col in sem_df.columns else ['']*len(keywords)
    cleaned_kw = [clean_kw(k) for k in keywords]

    def build_sims(df):
        v = TfidfVectorizer(ngram_range=(1,3),min_df=1,sublinear_tf=True)
        v.fit(df['content'].tolist()+cleaned_kw)
        return cosine_similarity(v.transform(cleaned_kw),v.transform(df['content']))

    svc_sims  = build_sims(service_df)
    blog_sims = build_sims(blog_df)
    progress_bar.progress(45)
    status_text.text("Phase 2: Matching keywords to pages...")

    results = []
    for i,kw in enumerate(keywords):
        if i%500==0:
            progress_bar.progress(min(45+int((i/len(keywords))*30),74))
            status_text.text(f"Phase 2: Mapping {i:,}/{len(keywords):,} keywords...")

        kw_low  = kw.lower().strip()
        intent  = 'Informational' if is_informational(kw) else 'Transactional'
        kw_cats = categorize_kw(kw)
        chosen=''; source=''; cs=0.0; gs=0.0

        sims_d = blog_sims[i] if intent=='Informational' else svc_sims[i]
        pool   = blog_df      if intent=='Informational' else service_df
        for idx in np.argsort(sims_d)[::-1]:
            s = float(sims_d[idx])
            if s<threshold: break
            url = get_url(pool.iloc[idx])
            if is_compatible(kw_cats,pool.iloc[idx]['cats'],kw,url):
                chosen=url; cs=round(s,4); source='Content match'; break

        gd = gsc_dedup.get(kw_low,{})
        gu = gd.get('url','')
        guc = categorize_url(gu) if gu else frozenset()
        gv  = gu and not is_excluded(gu) and is_compatible(kw_cats,guc,kw,gu)

        if gv and chosen and gu.lower().strip()==chosen.lower().strip():
            gs=0.40; source='Content + GSC confirmed'
        elif gv and not chosen:
            chosen=gu; gs=0.40; source='GSC fallback'

        fs = round(cs*0.7+gs*0.3,4)

        rp = str(your_pos[i]).strip()
        try: pn=float(rp)
        except: pn=None
        if pn:
            if pn<=10:    rs='Ranking p1-10'
            elif pn<=20:  rs='Quick win p11-20'
            elif pn<=50:  rs='Weak ranking p21-50'
            else:         rs='Very weak p51-100'
        else: rs='Not ranking'

        results.append({'Keyword':kw,'Volume':int(volumes[i]) if volumes[i] else 0,'Your Position':rp or 'N/A','Landing Page':chosen,'Intent':intent,'Content Score':cs,'GSC Score':gs,'Final Score':fs,'Match Source':source,'Ranking Status':rs,'_cats':kw_cats})

    progress_bar.progress(75)
    return results

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 3 — CLAUDE RELEVANCE
# ─────────────────────────────────────────────────────────────────────────────
def phase3_relevance(mapped, api_key, biz_desc, exclusions, status_text, progress_bar):
    status_text.text("Phase 3: Preparing Claude API relevance check...")
    progress_bar.progress(77)
    candidates = [r['Keyword'] for r in mapped if r['Final Score']<0.20]
    confirmed  = {r['Keyword']:'RELEVANT' for r in mapped if r['Final Score']>=0.20}
    if not candidates: return {r['Keyword']:'RELEVANT' for r in mapped}
    rel = claude_relevance(api_key,candidates,biz_desc,exclusions,status_text,progress_bar,78,93)
    return {**confirmed,**rel}

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 4 — BUILD EXCEL
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(gsc_df_val, mapped, rel_map):
    wb  = Workbook()
    bf  = Font(size=10)
    lf  = Font(size=10,color='0563C1')
    mf  = Font(size=10,italic=True,color='888888')

    C = {'H':'1D9E75','GR':'EAF3DE','BL':'E8F0FE','YL':'FFFBEA','RD':'FDECEA','AM':'FFF3CD','PU':'F3E8FE','GY':'F5F5F5','WH':'FFFFFF','DGR':'D4EDDA'}

    # TAB 1 — GSC Mapping Quality
    ws1 = wb.active; ws1.title='GSC Mapping Quality'
    hdr(ws1,1,['Query','Mapped URL','Clicks','Impressions','Position','CTR','Content Score','Mapping Status'])
    sf = {'Confirmed':make_fill(C['GR']),'Plausible':make_fill(C['YL']),'Suspicious':make_fill(C['RD'])}
    for i,row in gsc_df_val.iterrows():
        r=i+2; fill=sf.get(row['Mapping Status'],make_fill(C['WH']))
        for col,val in enumerate([row['Query'],row['Mapped URL'],row['Clicks'],row['Impressions'],row['Position'],row['CTR'],row['Content Score'],row['Mapping Status']],1):
            c=ws1.cell(row=r,column=col,value=val); c.fill=fill; c.font=lf if col==2 else bf
    col_w(ws1,[45,65,10,12,10,8,14,14]); ws1.freeze_panes='A2'

    # TAB 2 — Keyword Mapping
    ws2 = wb.create_sheet('Keyword Mapping')
    hdr(ws2,1,['Keyword','Volume','Your Position','Landing Page','Intent','Content Score','GSC Score','Final Score','Match Source','Ranking Status'])
    for i,r_d in enumerate(mapped):
        r=i+2; url=r_d['Landing Page']; src=r_d['Match Source']; rel=rel_map.get(r_d['Keyword'],'')
        if url and '/blog/' in url:     fill=make_fill(C['BL'])
        elif src=='GSC fallback':       fill=make_fill(C['YL'])
        elif url:                       fill=make_fill(C['GR'])
        elif rel in('RELEVANT','BORDERLINE'): fill=make_fill(C['PU'])
        else:                           fill=make_fill(C['WH'])
        for col,val in enumerate([r_d['Keyword'],r_d['Volume'],r_d['Your Position'],url or '',r_d['Intent'],r_d['Content Score'],r_d['GSC Score'],r_d['Final Score'],src or '',r_d['Ranking Status']],1):
            c=ws2.cell(row=r,column=col,value=val); c.fill=fill; c.font=lf if col==4 else bf
    col_w(ws2,[50,12,14,65,15,14,12,12,22,18]); ws2.freeze_panes='A2'

    # TAB 3 — Opportunity Classification
    ws3 = wb.create_sheet('Opportunity Classification')
    hdr(ws3,1,['Keyword','Volume','Your Position','Landing Page','Intent','Final Score','Opportunity Type','Action'])
    of = {'Confirmed existing page':make_fill(C['GR']),'Quick win — optimise':make_fill(C['DGR']),'Weak ranking':make_fill(C['YL']),'Page exists not ranking':make_fill(C['BL']),'Business relevant gap':make_fill(C['PU']),'True content gap':make_fill(C['RD'])}
    ACTION = {'Confirmed existing page':'Monitor — already ranking well','Quick win — optimise':'Optimise title, meta, H1 — almost ranking','Weak ranking':'Improve page content + build internal links','Page exists not ranking':'Optimise existing page for this keyword','Business relevant gap':'','True content gap':'Evaluate — may need new page'}
    for r_d in mapped:
        rel=rel_map.get(r_d['Keyword'],''); url=r_d['Landing Page']; rs=r_d['Ranking Status']; fs=r_d['Final Score']
        if url and rs=='Ranking p1-10':              opp='Confirmed existing page'
        elif rs=='Quick win p11-20':                 opp='Quick win — optimise'
        elif rs in['Weak ranking p21-50','Very weak p51-100']: opp='Weak ranking'
        elif url and fs>=0.15:                       opp='Page exists not ranking'
        elif rel in('RELEVANT','BORDERLINE'):        opp='Business relevant gap'
        else:                                        opp='True content gap'
        action = ACTION.get(opp,'')
        if opp=='Business relevant gap':
            action = 'Create new service page' if r_d['Intent']=='Transactional' else 'Create new blog post'
        fill=of.get(opp,make_fill(C['WH']))
        rn=ws3.max_row+1
        for col,val in enumerate([r_d['Keyword'],r_d['Volume'],r_d['Your Position'],url or '',r_d['Intent'],fs,opp,action],1):
            c=ws3.cell(row=rn,column=col,value=val); c.fill=fill; c.font=lf if col==4 else bf
    col_w(ws3,[50,12,14,65,15,12,28,45]); ws3.freeze_panes='A2'

    # TAB 4 — Business Relevant Gaps
    ws4 = wb.create_sheet('Business Relevant Gaps')
    hdr(ws4,1,['Keyword','Volume','Intent','Service Topic','Relevance','Action Needed'])
    gaps = sorted([r for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in('RELEVANT','BORDERLINE')],key=lambda x:-x['Volume'])
    for r_d in gaps:
        rel=rel_map.get(r_d['Keyword'],''); fill=make_fill(C['PU'] if rel=='RELEVANT' else C['AM'])
        action='Create new service page' if r_d['Intent']=='Transactional' else 'Create new blog post'
        rn=ws4.max_row+1
        for col,val in enumerate([r_d['Keyword'],r_d['Volume'],r_d['Intent'],classify_topic(r_d['Keyword']),rel,action],1):
            c=ws4.cell(row=rn,column=col,value=val); c.fill=fill; c.font=bf
    col_w(ws4,[52,12,15,25,14,30]); ws4.freeze_panes='A2'

    # TAB 5 — Priority Roadmap
    ws5 = wb.create_sheet('Priority Roadmap')
    ws5['A1']='SEO Content Opportunity Roadmap'; ws5['A1'].font=Font(bold=True,size=14,color='0F6E56'); ws5.merge_cells('A1:H1')
    quick_wins = sum(1 for r in mapped if r['Ranking Status']=='Quick win p11-20')
    weak_rank  = sum(1 for r in mapped if r['Ranking Status'] in['Weak ranking p21-50','Very weak p51-100'])
    page_not_rank = sum(1 for r in mapped if r['Landing Page'] and r['Final Score']>=0.15 and r['Ranking Status']=='Not ranking')
    biz_gaps   = len(gaps)
    confirmed  = sum(1 for r in mapped if r['Landing Page'] and r['Ranking Status']=='Ranking p1-10')
    hdr(ws5,3,['Opportunity Type','Count','Action','Priority',''])
    summary = [('Quick Wins (p11-20)',quick_wins,'Optimise existing pages','High',C['DGR']),('Weak Rankings (p21-100)',weak_rank,'Improve content + internal links','High',C['YL']),('Page Exists Not Ranking',page_not_rank,'Optimise page for keyword','Medium',C['BL']),('Business Relevant Gaps',biz_gaps,'Create new pages / blog posts','Medium',C['PU']),('Already Ranking Well',confirmed,'Monitor — no action needed','Low',C['GY'])]
    for i,(opp,cnt,act,pri,color) in enumerate(summary):
        r=i+4; fill=make_fill(color)
        for col,val in enumerate([opp,cnt,act,pri,''],1):
            c=ws5.cell(row=r,column=col,value=val); c.fill=fill; c.font=Font(size=10,bold=(col==4))

    # Detail section
    dr=len(summary)+7
    ws5.cell(row=dr,column=1,value='Detailed Action List — sorted by priority then volume').font=Font(bold=True,size=12,color='0F6E56')
    ws5.merge_cells(start_row=dr,start_column=1,end_row=dr,end_column=8); dr+=1
    hdr(ws5,dr,['Keyword','Volume','Your Position','Intent','Final Score','Opportunity','Action','Priority']); dr+=1

    PORDER = {'Quick win — optimise':1,'Weak ranking':2,'Page exists not ranking':3,'Business relevant gap':4,'True content gap':5,'Confirmed existing page':6}
    OF5 = {'Quick win — optimise':make_fill(C['DGR']),'Weak ranking':make_fill(C['YL']),'Page exists not ranking':make_fill(C['BL']),'Business relevant gap':make_fill(C['PU']),'True content gap':make_fill(C['RD']),'Confirmed existing page':make_fill(C['GY'])}
    ACT5 = {'Confirmed existing page':'Monitor','Quick win — optimise':'Optimise title + meta + H1','Weak ranking':'Improve content + internal links','Page exists not ranking':'Optimise existing page for this keyword','True content gap':'Evaluate for new page'}

    all_items = []
    for r_d in mapped:
        rel=rel_map.get(r_d['Keyword'],''); url=r_d['Landing Page']; rs=r_d['Ranking Status']; fs=r_d['Final Score']
        if url and rs=='Ranking p1-10':             opp='Confirmed existing page'
        elif rs=='Quick win p11-20':                opp='Quick win — optimise'
        elif rs in['Weak ranking p21-50','Very weak p51-100']: opp='Weak ranking'
        elif url and fs>=0.15:                      opp='Page exists not ranking'
        elif rel in('RELEVANT','BORDERLINE'):       opp='Business relevant gap'
        else:                                       opp='True content gap'
        action=ACT5.get(opp,'Create new service page' if r_d['Intent']=='Transactional' else 'Create new blog post')
        pri={1:'High',2:'High',3:'Medium',4:'Medium',5:'Low',6:'Low'}.get(PORDER.get(opp,5),'Low')
        all_items.append({**r_d,'opp':opp,'action':action,'pri':pri,'po':PORDER.get(opp,5)})

    all_items.sort(key=lambda x:(x['po'],-x['Volume']))
    for item in all_items:
        fill=OF5.get(item['opp'],make_fill(C['WH']))
        for col,val in enumerate([item['Keyword'],item['Volume'],item['Your Position'],item['Intent'],item['Final Score'],item['opp'],item['action'],item['pri']],1):
            c=ws5.cell(row=dr,column=col,value=val); c.fill=fill; c.font=Font(size=10,bold=(col==8))
        dr+=1

    col_w(ws5,[50,12,14,15,12,28,42,10]); ws5.freeze_panes=f'A{len(summary)+10}'

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.title("🔍 SEO Content Mapper")
st.markdown("Complete SEO gap analysis — Screaming Frog + GSC validation + Semrush mapping + Claude API business relevance.")

with st.sidebar:
    st.header("⚙️ Settings")
    st.markdown("---")
    threshold = st.slider("Match Score Threshold",0.10,0.40,0.15,0.01)
    st.markdown("**Content Weights**")
    sw=st.slider("URL Slug",1,8,5); tw=st.slider("Page Title",1,6,3); hw=st.slider("H1 Heading",1,4,2); mw=st.slider("Meta Description",1,3,1)
    weights=(sw,tw,hw,mw)
    st.markdown("---")
    st.markdown("**Score Guide**")
    st.markdown("| Score | Quality |\n|---|---|\n| 0.50+ | Strong ✅ |\n| 0.30–0.50 | Good ✅ |\n| 0.20–0.30 | Acceptable ⚠️ |\n| 0.15–0.20 | Weak — verify |\n| Blank | Content gap |")
    st.markdown("---")
    st.markdown("**Colour Coding**")
    st.markdown("🟢 Content match  🔵 Blog  🟡 GSC fallback  🟣 Business gap  ⬜ Filtered")

# FILE 1
st.markdown('<div class="section-header">📂 File 1 — Screaming Frog Export</div>', unsafe_allow_html=True)
st.caption("Screaming Frog → Bulk Export → All. Required columns: Address, Title 1, Meta Description 1, H1-1")
sf_file = st.file_uploader("Upload Screaming Frog file (.xlsx or .csv)", type=['xlsx','csv'], key='sf')
sf_df = None
if sf_file:
    try:
        sf_df = pd.read_csv(sf_file) if sf_file.name.endswith('.csv') else pd.read_excel(sf_file)
        sf_df.columns = [c.strip() for c in sf_df.columns]
        missing = [c for c in ['Address','Title 1','Meta Description 1','H1-1'] if c not in sf_df.columns]
        if missing:
            st.markdown(f'<div class="warn-box">⚠️ Missing columns: {missing}. Found: {list(sf_df.columns[:10])}</div>', unsafe_allow_html=True); sf_df=None
        else:
            if 'Status Code' in sf_df.columns: sf_df = sf_df[sf_df['Status Code']==200].reset_index(drop=True)
            sf_df = sf_df[sf_df['Address'].notna()].reset_index(drop=True)
            st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(sf_df):,}</strong> pages</div>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)

# FILE 2
st.markdown('<div class="section-header">📂 File 2 — GSC Export</div>', unsafe_allow_html=True)
st.caption("Looker Studio export. Required columns: Query, Landing Page, Clicks, Impressions, Position, CTR")
gsc_file = st.file_uploader("Upload GSC Excel file", type=['xlsx'], key='gsc')
gsc_df = None
if gsc_file:
    try:
        gsc_xl=pd.read_excel(gsc_file,sheet_name=None); gsc_sheets=list(gsc_xl.keys())
        gsc_sheet=st.selectbox("Select GSC sheet",gsc_sheets,key='gs')
        raw=gsc_xl[gsc_sheet]; cols=list(raw.columns)
        with st.expander("Map GSC columns"):
            q_c  =st.selectbox("Query",cols,index=cols.index('Query') if 'Query' in cols else 0,key='qc')
            lp_c =st.selectbox("Landing Page",cols,index=cols.index('Landing Page') if 'Landing Page' in cols else 1,key='lpc')
            clk_c=st.selectbox("Clicks",cols,index=next((i for i,c in enumerate(cols) if 'click' in c.lower()),2),key='clkc')
            imp_c=st.selectbox("Impressions",cols,index=next((i for i,c in enumerate(cols) if 'impression' in c.lower()),3),key='impc')
            pos_c=st.selectbox("Position",cols,index=next((i for i,c in enumerate(cols) if 'position' in c.lower()),min(4,len(cols)-1)),key='posc')
            ctr_c=st.selectbox("CTR",cols,index=next((i for i,c in enumerate(cols) if 'ctr' in c.lower()),min(5,len(cols)-1)),key='ctrc')
        gsc_df=raw.rename(columns={q_c:'Query',lp_c:'Landing Page',clk_c:'Clicks',imp_c:'Impressions',pos_c:'Position',ctr_c:'CTR'})
        st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(gsc_df):,}</strong> GSC rows</div>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)

# FILE 3
st.markdown('<div class="section-header">📂 File 3 — Semrush Keyword Gap Export</div>', unsafe_allow_html=True)
st.caption("Full keyword gap export from Semrush. Required columns: Keyword, Volume, your domain's position column.")
sem_file = st.file_uploader("Upload Semrush Keyword Gap Excel file", type=['xlsx'], key='sem')
sem_df = None; your_col = None
if sem_file:
    try:
        sem_xl=pd.read_excel(sem_file,sheet_name=None); sem_sheets=list(sem_xl.keys())
        sem_sheet=st.selectbox("Select Semrush sheet",sem_sheets,key='ss')
        raw=sem_xl[sem_sheet]; cols=list(raw.columns)
        with st.expander("Map Semrush columns"):
            kw_c =st.selectbox("Keyword column",cols,index=cols.index('Keyword') if 'Keyword' in cols else 0,key='skc')
            vol_c=st.selectbox("Volume column",cols,index=cols.index('Volume') if 'Volume' in cols else 1,key='svc')
            your_col=st.selectbox("Your domain position column",cols,index=0,help="Column showing YOUR site's position e.g. cellinoplumbing.com",key='sdc')
        sem_df=raw.rename(columns={kw_c:'Keyword',vol_c:'Volume'})
        sem_df=sem_df[sem_df['Keyword'].notna()].reset_index(drop=True)
        st.markdown(f'<div class="info-box">✅ Loaded <strong>{len(sem_df):,}</strong> keywords</div>', unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)

# BUSINESS CONTEXT
st.markdown('<div class="section-header">🏢 Business Context + Claude API</div>', unsafe_allow_html=True)
col_a,col_b = st.columns(2)
with col_a:
    api_key=st.text_input("Claude API Key",type="password",placeholder="sk-ant-...",help="Required. Get from console.anthropic.com")
    biz_desc=st.text_area("Business Description *",placeholder="e.g. Cellino Plumbing provides residential and commercial plumbing, heating, cooling and electrical services in Buffalo NY. Services include furnace repair, AC installation, drain cleaning, water heater replacement and electrical panel upgrades.",height=130)
with col_b:
    exclusions=st.text_area("Services NOT offered (optional — one per line)",placeholder="e.g.\noil boiler\nseptic tank\nwell pump\nsolar panel",height=130)
    st.caption("Leave blank if unsure — Claude will use the business description to judge relevance.")

# VALIDATE & RUN
st.markdown("---")
issues=[]
if sf_df is None:         issues.append("File 1 (Screaming Frog) not uploaded or has column errors")
if gsc_df is None:        issues.append("File 2 (GSC) not uploaded or has errors")
if sem_df is None:        issues.append("File 3 (Semrush) not uploaded or has errors")
if not api_key.strip():   issues.append("Claude API key is required")
if not biz_desc.strip():  issues.append("Business description is required")

for issue in issues:
    st.markdown(f'<div class="warn-box">⚠️ {issue}</div>', unsafe_allow_html=True)

if st.button("🚀 Run Full Analysis", disabled=bool(issues), use_container_width=True, type="primary"):
    progress_bar = st.progress(0)
    status_text  = st.empty()
    try:
        gsc_val_df, url_df_proc, gsc_dedup, stop = phase1_validate_gsc(gsc_df, sf_df, weights, status_text, progress_bar)
        mapped = phase2_map_keywords(sem_df, url_df_proc, gsc_dedup, weights, threshold, your_col, status_text, progress_bar)
        rel_map = phase3_relevance(mapped, api_key, biz_desc, exclusions, status_text, progress_bar)
        progress_bar.progress(95); status_text.text("Building Excel output...")
        excel_buf = build_excel(gsc_val_df, mapped, rel_map)
        progress_bar.progress(100); status_text.text("Done!")

        st.markdown("---")
        st.subheader("📊 Analysis Complete")
        confirmed_gsc  = len(gsc_val_df[gsc_val_df['Mapping Status']=='Confirmed'])
        suspicious_gsc = len(gsc_val_df[gsc_val_df['Mapping Status']=='Suspicious'])
        mapped_kw      = sum(1 for r in mapped if r['Landing Page'])
        quick_wins     = sum(1 for r in mapped if r['Ranking Status']=='Quick win p11-20')
        biz_gaps       = sum(1 for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in('RELEVANT','BORDERLINE'))

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        for col,num,label in [(c1,len(gsc_val_df),'GSC queries'),(c2,confirmed_gsc,'GSC confirmed'),(c3,suspicious_gsc,'GSC suspicious'),(c4,quick_wins,'Quick wins p11-20'),(c5,mapped_kw,'Keywords mapped'),(c6,biz_gaps,'Business gaps')]:
            col.markdown(f'<div class="metric-card"><div class="metric-num">{num:,}</div><div class="metric-label">{label}</div></div>', unsafe_allow_html=True)

        st.markdown("")
        t1,t2,t3 = st.tabs(["GSC Validation","Keyword Mapping","Business Gaps"])
        with t1: st.dataframe(gsc_val_df.head(200),use_container_width=True,height=350)
        with t2:
            prev=pd.DataFrame([{'Keyword':r['Keyword'],'Volume':r['Volume'],'Your Position':r['Your Position'],'Landing Page':r['Landing Page'],'Final Score':r['Final Score'],'Source':r['Match Source'],'Ranking':r['Ranking Status']} for r in mapped]).sort_values('Final Score',ascending=False)
            st.dataframe(prev.head(200),use_container_width=True,height=350)
        with t3:
            gp=pd.DataFrame([{'Keyword':r['Keyword'],'Volume':r['Volume'],'Intent':r['Intent'],'Relevance':rel_map.get(r['Keyword'],''),'Topic':classify_topic(r['Keyword'])} for r in mapped if not r['Landing Page'] and rel_map.get(r['Keyword'],'') in('RELEVANT','BORDERLINE')]).sort_values('Volume',ascending=False)
            st.dataframe(gp.head(200),use_container_width=True,height=350)

        st.markdown("---")
        st.download_button("⬇️ Download Full Excel Output (5 tabs)", data=excel_buf, file_name="seo_content_mapping_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
        st.markdown('<div class="info-box"><strong>5 tabs:</strong> GSC Mapping Quality | Keyword Mapping | Opportunity Classification | Business Relevant Gaps | Priority Roadmap</div>', unsafe_allow_html=True)

    except Exception as e:
        progress_bar.progress(0)
        st.markdown(f'<div class="error-box">❌ Error: {str(e)}</div>', unsafe_allow_html=True)
        st.exception(e)
