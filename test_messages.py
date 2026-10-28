"""Tests for the message library: upload samples, exact use, figure filling, data draft, new-file update."""
import pathlib, subprocess, openpyxl
from playwright.sync_api import sync_playwright
HERE = pathlib.Path(__file__).parent; OUT = HERE/'output'
APP = (HERE.parent/'index.html').resolve().as_uri()
SRC = '/mnt/user-data/uploads/Pm_SHRI_DBT_Verification.xlsx'
NAMES = (HERE/'test_district_list.txt').read_text(encoding='utf-8').split('\n')
res = []
def check(n, c, d=''): res.append(c); print(('PASS ' if c else 'FAIL ') + n + ('' if c else f'  [{d}]'))
ws = openpyxl.load_workbook(SRC).active
rows = [list(r) for r in ws.iter_rows(min_row=5, values_only=True)]
aad = sorted([(r[12]/r[6]*100, NAMES[i]) for i, r in enumerate(rows)], key=lambda x: (-x[0], x[1]))
SETUP = """()=>{const S=__app.S;S.cfg.district.mode='list';S.cfg.district.list=%s;S.cfg.district.verifiedSig=S.fileSig;
 S.cfg.calcs=[{id:'a1',name:'Aadhaar verification (%%)',type:'percent',a:'gaurdian adhar > verify',b:'enrollment status > total enrollment',multiplier:100,decimals:1,suffix:'%%',missing:'blank'}];
 S.cfg.selected=[{key:'enrollment status > total enrollment',label:''},{key:'calc:a1',label:''}];S.cfg.ranking.by='calc:a1';S.cfg.ranking.better='higher';
 S.cfg.message.period='October 2026';S.step=5;__app.render();}""" % repr(NAMES).replace("'", '"')
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={'width': 1400, 'height': 1000})
    errs = []; pg.on('pageerror', lambda e: errs.append(str(e)))
    answers = []; pg.on('dialog', lambda d: d.accept(answers.pop(0) if answers else ''))
    pg.goto(APP); pg.evaluate('localStorage.clear()'); pg.reload()
    pg.set_input_files('#fileInput', SRC); pg.wait_for_selector('text=Confirm the data')
    pg.evaluate(SETUP); pg.wait_for_selector('text=Sample message library')
    lib = lambda: pg.evaluate('JSON.parse(localStorage.getItem("dpr.templates")||"{}")')
    # upload three samples at once
    pg.set_input_files('#sampleInput', sorted(str(f) for f in (HERE/'samples').iterdir()))
    pg.wait_for_selector('text=Monthly review - figures')
    L = lib()
    check('upload: 3 samples (.txt + .docx) saved', len(L) == 3, list(L))
    check('upload: Hindi “विषय:” line becomes the subject', L['Aadhaar reminder - exact']['subject'].startswith('विषय:'))
    check('upload: .docx paragraphs read', 'Districts needing immediate attention' in L['Monthly review - figures']['body'])
    check('upload: sample with {fields} marked fillable, plain one exact', L['Monthly review - figures']['mode'] == 'data' and L['Aadhaar reminder - exact']['mode'] == 'exact')
    # exact use
    pg.click('[data-act=useSample][data-n="Aadhaar reminder - exact"]')
    M = pg.evaluate('__app.S.cfg.message')
    check('exact: body copied word for word', M['body'] == L['Aadhaar reminder - exact']['body'] and M['mode'] == 'exact')
    pg.evaluate('()=>{__app.S.cfg.report.format="detailed";__app.preparePrint()}'); pg.emulate_media(media='print')
    pg.pdf(path=str(OUT/'msg_exact.pdf'), prefer_css_page_size=True, print_background=True); pg.emulate_media(media='screen')
    t = subprocess.run(['pdftotext', str(OUT/'msg_exact.pdf'), '-'], capture_output=True, text=True).stdout
    check('exact: text appears in the PDF', 'तीन दिवस' in t.replace('\u200c', ''), t[:300])
    # figure-filling sample
    pg.click('[data-act=useSample][data-n="Monthly review - figures"]')
    prev = pg.inner_text('[data-block=msgPreview]')
    best = aad[0]
    check('fill: subject filled with indicator and period', 'Progress on Aadhaar verification (%) – October 2026' in prev, prev[:200])
    check('fill: best district and value filled', f'{best[1]} leads with {best[0]:.1f}%' in prev, prev)
    worst5 = pg.evaluate('__app.S.D.rank.botN(__app.S.D.rank.list,5).map(e=>e.row.name)')
    check('fill: bottom list names the weakest districts', all(n in prev for n in worst5))
    # data-driven draft using plain sample's opening/closing
    pg.click('[data-act=useSample][data-n="Plain opening and closing.docx"]') if pg.query_selector('[data-n="Plain opening and closing.docx"]') else pg.click('[data-act=useSample][data-n="Plain opening and closing"]')
    pg.click('label:has(> input[name=mmode][value=data])')
    pg.select_option('select[data-bind="message.draftLang"]', 'both')
    pg.click('[data-act=draftMsg]')
    body = pg.evaluate('__app.S.cfg.message.body'); prev = pg.inner_text('[data-block=msgPreview]')
    check('draft: keeps sample opening and closing', body.startswith('समस्त जिला बेसिक शिक्षा अधिकारी,') and body.rstrip().endswith('अनुपालन आख्या प्रेषित करें।'), body[:80])
    check('draft: Hindi and English middle written', 'सर्वश्रेष्ठ प्रदर्शन' in body and 'Best-performing districts' in body)
    top5 = [n for _, n in aad[:5]]
    check('draft: top list shows correct ranks and values', f'1. {aad[0][1]} – {aad[0][0]:.1f}% (rank 1)' in prev and all(n in prev for n in top5), prev[:400])
    mean = sum(v for v, _ in aad)/75
    check('draft: below-average count correct', f'{sum(1 for v,_ in aad if v < mean)} जनपद जनपदीय औसत से नीचे' in prev)
    pg.screenshot(path=str(OUT/'07_message_library.png'), full_page=False, clip={'x':0,'y':0,'width':1400,'height':1000})
    # district-specific page
    pg.evaluate('()=>{const M=__app.S.cfg.message;M.districtTemplate="जनपद {district} का {indicator} {value} है; स्थान {rank}/{total_districts}।";__app.S.cfg.report.format="district";__app.preparePrint()}')
    pg.emulate_media(media='print'); pg.pdf(path=str(OUT/'msg_district.pdf'), prefer_css_page_size=True, print_background=True); pg.emulate_media(media='screen')
    t = subprocess.run(['pdftotext', '-f', '1', '-l', '1', str(OUT/'msg_district.pdf'), '-'], capture_output=True, text=True).stdout
    check('district pages: placeholders filled per district', f'स्थान 1/75' in t and best[1] in t, t[:300])
    # new file: figures update automatically
    pg.evaluate('__app.S.cfg.report.format="detailed"')
    answers.append('msg test'); pg.click('[data-act=saveProfileAs]')
    pg.set_input_files('#fileInput', str(HERE/'next_period.xlsx')); pg.wait_for_selector('text=Reuse your saved setup?')
    pg.select_option('select[data-compat="enrollment status > total enrollment"]', 'enrollment status > total enrolment'); pg.click('[data-act=applyCompat]')
    pg.evaluate('()=>{__app.S.cfg.district.verifiedSig=__app.S.fileSig;__app.S.step=5;__app.render()}')
    prev2 = pg.inner_text('[data-block=msgPreview]')
    new_best = sorted([((r[12]+10)/r[6]*100, NAMES[i]) for i, r in enumerate(rows)], key=lambda x: (-x[0], x[1]))[0]
    check('new file: message figures recalculated from new data', f'1. {new_best[1]} – {new_best[0]:.1f}% (rank 1)' in prev2, prev2[:300])
    check('library survives and is included in setup export', len(lib()) == 3)
    check('no JavaScript errors', not errs, errs[:2])
    b.close()
print(f'\n{sum(res)}/{len(res)} message checks passed')
