"""End-to-end test of the dashboard in headless Chromium against the reference workbook.
Every computed value is compared with an independent pandas/openpyxl calculation."""
import json, subprocess, pathlib, openpyxl
from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).parent
APP = (HERE.parent/'index.html').resolve().as_uri()
SRC = '/mnt/user-data/uploads/Pm_SHRI_DBT_Verification.xlsx'
OUT = HERE/'output'; OUT.mkdir(exist_ok=True)
NAMES = (HERE/'test_district_list.txt').read_text(encoding='utf-8').split('\n')
results = []
def check(name, cond, detail=''):
    results.append((name, bool(cond), detail)); print(('PASS ' if cond else 'FAIL ') + name + (f'  [{detail}]' if detail and not cond else ''))

# ---- independent expected values ----
ws = openpyxl.load_workbook(SRC).active
rows = [list(r) for r in ws.iter_rows(min_row=5, values_only=True)]
exp = []
for i, r in enumerate(rows):
    te, gv, it, ip = r[6], r[12], r[3], r[5]
    exp.append({'name': NAMES[i], 'aad': gv/te*100 if te else None, 'pend': ip/it*100 if it else None, 'te': te})
def comp_rank(vals, higher):
    # competition ranking on values rounded to 1 dp, ties ordered by name
    v = sorted([e for e in vals if e[1] is not None], key=lambda e: ((-e[1] if higher else e[1]), e[0]))
    out, prev, rk = [], None, 0
    for i, (n, x) in enumerate(v):
        k = round(x + 1e-12, 1)
        if k != prev: rk = i + 1
        prev = k; out.append((rk, n, x))
    return out

with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={'width': 1440, 'height': 1000})
    errors = []; pg.on('pageerror', lambda e: errors.append(str(e))); pg.on('console', lambda m: m.type == 'error' and errors.append(m.text))
    dialog_answers = []
    pg.on('dialog', lambda d: d.accept(dialog_answers.pop(0) if dialog_answers else ''))
    pg.goto(APP); pg.evaluate('localStorage.clear()'); pg.reload()
    D = lambda js: pg.evaluate('(() => { const S = window.__app.S, D = S.D; return ' + js + '; })()')

    # 1. upload and header interpretation
    pg.set_input_files('#fileInput', SRC); pg.wait_for_selector('text=Confirm the data')
    check('1a three heading rows detected', D('D.headerRows') == 3, D('D.headerRows'))
    names = D('D.cols.filter(c=>c.kind==="orig").map(c=>c.name)')
    print('   interpreted names:', names)
    expected_names = ['Total School','Total AWC','New Enrollment','Given By ICDS - Total','Given By ICDS - Verify','Given By ICDS - Pending','Total Enrollment','Balvatika 1 (3-4 Years)','Balvatika 2 (4-5 Years)','Balvatika 3 (5-6 Years)','Gaurdian Adhar - Pending','Gaurdian Adhar - Not Verify','Gaurdian Adhar - Verify']
    check('1b grouped headings become unique readable names', names == expected_names, names)
    # 2. summary row
    check('2 row 4 identified as the overall total, 75 district rows', D('D.summary.map(r=>r.src)') == [4] and D('D.dataRows.length') == 75)
    # 3. district identifiers
    check('3a no district names found -> flagged', D('D.districtReady') is False and D('D.textCols.length') == 0)
    check('3b ranking blocked without names', D('D.rank.ready') is False)
    pg.screenshot(path=str(OUT/'01_confirm_data.png'), full_page=False)
    # rename the Guardian columns through the UI (aliases)
    for key, nm in [('gaurdian adhar > pending','Guardian Aadhaar - Pending'),('gaurdian adhar > not verify','Guardian Aadhaar - Not Verify'),('gaurdian adhar > verify','Guardian Aadhaar - Verify')]:
        pg.fill(f'input[data-alias="{key}"]', nm)
    check('3c renamed headings applied', 'Guardian Aadhaar - Verify' in D('D.cols.map(c=>c.name)'))
    # paste district list
    pg.click('label:has(> input[name=dm][value=list])'); pg.wait_for_selector('textarea[data-bind="district.list"]')
    pg.fill('textarea[data-bind="district.list"]', '\n'.join(NAMES[:74]))
    check('3d wrong count rejected', D('D.districtReady') is False and '74 names' in D('D.districtErrors.join(" ")'))
    pg.fill('textarea[data-bind="district.list"]', '\n'.join(NAMES))
    check('3e correct list still needs order confirmation', D('D.districtReady') is False and D('!!D.needsVerify'))
    pg.evaluate('window.__app.render()')
    pg.check('input[data-act-change=verifyList]')
    check('3f districts ready after confirmation', D('D.districtReady') is True)
    # consistency check through UI
    pg.click('summary:has-text("Add a check")')
    pg.select_option('#chkTotal', 'enrollment status > total enrollment')
    for k in ['gaurdian adhar > pending','gaurdian adhar > not verify','gaurdian adhar > verify']: pg.check(f'.chkPart[value="{k}"]')
    pg.click('[data-act=addCheck]')
    nbad = D('D.checksOut[0].bad.length')
    check('3g Guardian parts vs Total Enrollment check reports mismatches', nbad > 0, nbad); print('   rows where Guardian Aadhaar parts != Total Enrollment:', nbad)

    # 4. select and reorder columns
    pg.click('[data-act=go][data-step="3"]')
    for k in ['enrollment status > total enrollment','enrollment status > given by icds > total','gaurdian adhar > verify']: pg.check(f'input[data-pick="{k}"]')
    sel = D('S.cfg.selected.map(s=>s.key)')
    pg.click('[data-act=mv][data-i="2"][data-d="-1"]')
    sel2 = D('S.cfg.selected.map(s=>s.key)')
    check('4 select and reorder', sel2 == [sel[0], sel[2], sel[1]], sel2)
    pg.fill('input[data-label="0"]', 'कुल नामांकन (Total Enrollment)')

    # 5. calculations via UI
    pg.click('[data-act=go][data-step="4"]')
    def add_calc(name, typ, a, b, dec=1):
        pg.fill('input[data-sbind="draft.name"]', name)
        pg.select_option('select[data-act-change=draftType]', typ)
        pg.select_option('select[data-sbind="draft.a"]', a); pg.select_option('select[data-sbind="draft.b"]', b)
        pg.fill('input[data-sbind="draft.decimals"]', str(dec))
        pg.click('[data-act=saveCalc]')
    add_calc('Aadhaar verification (%)', 'percent', 'gaurdian adhar > verify', 'enrollment status > total enrollment')
    add_calc('ICDS pending (%)', 'percent', 'enrollment status > given by icds > pending', 'enrollment status > given by icds > total')
    vals = D('D.dataRows.map(r=>[r.name, r.v["calc:"+S.cfg.calcs[0].id], r.v["calc:"+S.cfg.calcs[1].id]])')
    ok = all((abs(v[1]-e['aad'])<1e-9 if e['aad'] is not None else v[1] is None) and (abs(v[2]-e['pend'])<1e-9 if e['pend'] is not None else v[2] is None) for v,e in zip(vals,exp))
    check('5a percentages match independent calculation for all 75 rows', ok)
    zero = sum(1 for e in exp if e['pend'] is None)
    check('5b zero denominators give blanks, not numbers', D('D.calcStats["calc:"+S.cfg.calcs[1].id].zero') == zero, zero)
    pg.screenshot(path=str(OUT/'02_calculations.png'), full_page=True)
    # custom formula
    pg.fill('input[data-sbind="draft.name"]', 'Aadhaar seeded (%)')
    pg.select_option('select[data-act-change=draftType]', 'custom')
    pg.fill('#exprBox', '([Guardian Aadhaar - Verify] + [Guardian Aadhaar - Not Verify]) ÷ [Total Enrollment] × 100')
    prev = pg.inner_text('[data-block=draftPreview]')
    check('5c custom formula previews', 'Formula' in prev and 'Unknown' not in prev, prev[:120])
    pg.click('[data-act=saveCalc]')
    v3 = D('D.dataRows[0].v["calc:"+S.cfg.calcs[2].id]'); r0 = rows[0]
    check('5d custom formula correct', abs(v3 - (r0[12]+r0[11])/r0[6]*100) < 1e-9)

    # 6. ranking both directions
    pg.click('[data-act=go][data-step="5"]')
    aad = 'calc:' + D('S.cfg.calcs[0].id')
    pg.select_option('select[data-bind="ranking.by"]', aad)
    check('6a no ranking until direction is chosen', D('D.rank.ready') is False)
    pg.click('label:has(> input[name=rb][value=higher])')
    got = D('D.rank.list.map(e=>[e.rank,e.row.name])')
    want = [[r, n] for r, n, _ in comp_rank([(e['name'], e['aad']) for e in exp], True)]
    check('6b higher-is-better ranking matches', got == want, (got[:5], want[:5]))
    print('   top 3:', got[:3], ' bottom 2:', got[-2:])
    pg.click('label:has(> input[name=rb][value=lower])')
    got = D('D.rank.list.map(e=>[e.rank,e.row.name])')
    want = [[r, n] for r, n, _ in comp_rank([(e['name'], e['aad']) for e in exp], False)]
    check('6c lower-is-better ranking matches', got == want)
    pg.click('label:has(> input[name=rb][value=higher])')
    pg.select_option('select[data-bind="ranking.show"]', 'bottom'); pg.select_option('select[data-bind="ranking.n"]', '10')
    vis = D('D.rank.visible.map(e=>e.rank)')
    check('6d bottom 10 shows the 10 weakest (ties included)', len(vis) >= 10 and min(vis) >= D('D.rank.list[D.rank.list.length-10].rank'), vis)
    pg.select_option('select[data-bind="ranking.show"]', 'all')
    # tie handling: dense
    pg.select_option('select[data-bind="ranking.ties"]', 'dense')
    check('6e dense ties have no gaps', sorted(set(D('D.rank.list.map(e=>e.rank)'))) == list(range(1, len(set(D('D.rank.list.map(e=>e.rank)')))+1)))
    pg.select_option('select[data-bind="ranking.ties"]', 'competition')
    pg.select_option('select[data-bind="ranking.baseCol"]', 'enrollment status > total enrollment')
    pg.fill('input[data-bind="ranking.baseMin"]', '100')
    got = D('D.rank.list.map(e=>[e.rank,e.row.name])')
    want = [[r, n] for r, n, _ in comp_rank([(e['name'], e['aad']) for e in exp if e['te'] >= 100], True)]
    check('6f minimum base excludes tiny denominators and re-ranks', got == want and len(got) == sum(1 for e in exp if e['te'] >= 100), (got[:2], len(got)))
    pg.fill('input[data-bind="ranking.baseMin"]', '')
    pg.select_option('select[data-bind="ranking.baseCol"]', '')

    # 7. stakeholder message
    pg.fill('input[data-bind="message.subject"]', 'विषय: अभिभावक आधार सत्यापन की प्रगति')
    body = 'समस्त जनपद अपने लंबित आधार सत्यापन मामलों की समीक्षा करें तथा एक सप्ताह में प्रगति सुनिश्चित करें।\nAll districts to review pending Aadhaar verification and report progress.'
    pg.fill('textarea[data-bind="message.body"]', body)
    pg.fill('input[data-bind="message.period"]', 'October 2026')
    pg.fill('input[data-bind="message.signatory"]', 'राज्य परियोजना निदेशक'); pg.fill('input[data-bind="message.designation"]', 'State Project Director')
    check('7a message preview shows Hindi', 'लंबित' in pg.inner_text('[data-block=msgPreview]'))
    dialog_answers.append('Aadhaar review message'); pg.click('[data-act=saveTpl]')
    check('7b message template saved', 'Aadhaar review message' in pg.evaluate('Object.keys(JSON.parse(localStorage.getItem("dpr.templates")))'))
    pg.screenshot(path=str(OUT/'03_rank_message.png'), full_page=True)

    # 8/9. PDF export and verification
    pg.click('[data-act=go][data-step="6"]')
    pg.fill('input[data-bind="report.title"]', 'PM SHRI DBT verification — district review')
    pg.fill('input[data-bind="report.programme"]', 'PM SHRI · Balvatika DBT')
    pg.fill('input[data-bind="report.footer"]', 'For official use only / केवल आधिकारिक उपयोग हेतु')
    pg.screenshot(path=str(OUT/'04_preview.png'), full_page=False)
    for fmt in ['detailed', 'summary', 'topbottom', 'district']:
        pg.click(f'label:has(> input[name=fmt][value={fmt}])')
        pg.evaluate('window.__app.preparePrint()'); pg.emulate_media(media='print')
        pdf = OUT/f'report_{fmt}.pdf'
        pg.pdf(path=str(pdf), prefer_css_page_size=True, print_background=True)
        pg.emulate_media(media='screen')
        txt = subprocess.run(['pdftotext', '-layout', str(pdf), '-'], capture_output=True, text=True).stdout
        pages = int(subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout.split('Pages:')[1].split()[0])
        print(f'   {fmt}: {pages} page(s)')
        if fmt == 'detailed':
            check('8a multipage PDF with page numbers', pages >= 2 and f'Page 2 of {pages}' in txt, pages)
            check('8b table heading repeats on page 2', txt.split('\f')[1].count('Rank') >= 1)
            check('8c Hindi text present in PDF', 'लंबित' in txt.replace(' ', '') or 'लं' in txt)
            # verify the order and values printed in the PDF
            want = comp_rank([(e['name'], e['aad']) for e in exp], True)
            txt = txt[txt.find('Rank    '):] if 'Rank    ' in txt else txt[txt.find('Rank'):]
            pos = [txt.find(n + ' ') if txt.find(n + ' ') >= 0 else txt.find(n) for _, n, _ in want]
            check('9a PDF lists districts in ranking order', all(a < b for a, b in zip(pos, pos[1:])) and min(pos) >= 0)
            okv = all(f'{x:.1f}%' in txt for _, _, x in want)
            check('9b every ranked value appears in the PDF', okv)
            line = next(l for l in txt.splitlines() if want[0][1] in l)
            check('9c rank 1 row has the right value on the same line', f'{want[0][2]:.1f}%' in line and line.strip().startswith('1'), line)
        if fmt == 'district':
            check('8d one page per district', pages == 75, pages)
    check('no JavaScript errors so far', not errors, errors[:3])

    # Excel export matches
    with pg.expect_download() as dl: pg.click('[data-act=exportXlsx]')
    path = OUT/'exported_table.xlsx'; dl.value.save_as(path)
    x = openpyxl.load_workbook(path).active
    xr = [(r[0], r[1]) for r in x.iter_rows(min_row=2, values_only=True) if r[0] != '' and r[0] is not None]
    check('9d exported table ranks match', [list(t) for t in xr] == [[r, n] for r, n, _ in comp_rank([(e['name'], e['aad']) for e in exp], True)][:len(xr)] and len(xr) == 75)

    # 10. save setup, upload next period file with changed heading, reuse
    dialog_answers.append('PM SHRI DBT - Aadhaar review'); pg.click('[data-act=saveProfileAs]')
    pg.click('[data-act=go][data-step="1"]')
    pg.set_input_files('#fileInput', str(HERE/'next_period.xlsx')); pg.wait_for_selector('text=Reuse your saved setup?')
    miss = D('S.compat.missing.map(m=>m.label)')
    check('10a changed heading detected, not silently applied', miss == ['Total Enrollment'], miss)
    check('10b steps 3+ blocked until confirmed', D('!!S.compat'))
    pg.click('[data-act=applyCompat]')
    check('10c cannot apply without mapping', D('!!S.compat'))
    pg.select_option('select[data-compat="enrollment status > total enrollment"]', 'enrollment status > total enrolment')
    pg.click('[data-act=applyCompat]')
    check('10d mapping applied', not D('S.compat'))
    check('10e district list kept but order must be re-confirmed for the new file', D('D.districtReady') is False and D('!!D.needsVerify'))
    pg.check('input[data-act-change=verifyList]')
    v = D('D.dataRows[0].v[S.cfg.ranking.by]')
    check('10f calculation uses new file values', abs(v - (rows[0][12] + 10) / rows[0][6] * 100) < 1e-9, v)
    check('10g ranking, columns and message reused', D('D.rank.ready') and D('S.cfg.selected.length') >= 5 and 'लंबित' in D('S.cfg.message.body'))
    pg.screenshot(path=str(OUT/'05_reused_setup.png'), full_page=False)
    check('no JavaScript errors', not errors, errors[:3])
    b.close()

passed = sum(1 for r in results if r[1]); print(f'\n{passed}/{len(results)} checks passed')
json.dump([{'check': n, 'passed': ok} for n, ok, _ in results], open(OUT/'test_results.json', 'w'), indent=1, ensure_ascii=False)
