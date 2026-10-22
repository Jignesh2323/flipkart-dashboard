# Merges freshly collected Seller Hub settlement data (fk_blob "settle_<acc>" in the Flipkart Supabase project,
# written by the in-page collector) into settlements\*.json, daily\ and charged\, replacing the collected months
# only. For accounts in --hub (no API orders: VUSTICA, BRAINLE) the blob must hold ALL months, and hubsales.js is
# rebuilt for them from its settled-sale rows. Run build-settle.ps1 afterwards.
# Usage: python -I merge-settle.py --months 2026-09,2026-10 [--hub Acc1_VUSTICA,Acc2_BRAINLE]
import json, os, sys, urllib.request, collections

ROOT = os.path.dirname(os.path.abspath(__file__))
SB, KEY = 'https://sfccrggijuiupoipixbi.supabase.co', 'sb_publishable_PQeCkrwv-JMQTXrqvxzNHg_zsMcglij'
ACCS = ['Acc1_VUSTICA', 'Acc2_BRAINLE', 'Acc3_Kidoca', 'Acc4_TAUDI', 'Acc5_ZOTIMO', 'Acc6_Pivdo', 'Acc7_Wiska', 'Acc8_ARKZO', 'Acc9_AVEDIA']
args = dict(zip(sys.argv[1::2], sys.argv[2::2]))
months = args['--months'].split(',')
hub = [a for a in args.get('--hub', '').split(',') if a]


def blob(name):
    req = urllib.request.Request(f'{SB}/rest/v1/fk_blob?name=eq.{name}&select=data,updated_at', headers={'apikey': KEY})
    r = json.loads(urllib.request.urlopen(req).read())
    return (json.loads(r[0]['data']), r[0]['updated_at']) if r else (None, None)


def rw(path, fn, default):
    old = json.load(open(path, encoding='utf-8')) if os.path.exists(path) else default
    new = fn(old)
    with open(path, 'w', encoding='utf-8') as f: json.dump(new, f, separators=(',', ':'), ensure_ascii=False)


inm = lambda d: d[:7] in months
hubsales = {}
for acc in ACCS:
    o, at = blob('settle_' + acc)
    if not o or isinstance(o, str):
        print(f'{acc}: blob nahi mila'); continue
    if o.get('e'): print(f'{acc}: {len(o["e"])} payment-din me error: {o["e"][:5]}')
    main = o['main']
    def fmain(old):
        p = [x for x in old.get('p', []) if not inm(x[0])] + [x for x in main['p'] if inm(x[0])]
        m = {k: v for k, v in old.get('m', {}).items() if k not in months}
        m.update({k: v for k, v in main['m'].items() if k in months})
        return {'p': sorted(p), 'm': dict(sorted(m.items()))}
    rw(os.path.join(ROOT, 'settlements', acc + '.json'), fmain, {})
    rw(os.path.join(ROOT, 'settlements', 'daily', acc + '.json'),
       lambda old: dict(sorted({**{k: v for k, v in old.items() if not inm(k)}, **{k: v for k, v in o['dd'].items() if inm(k)}}.items())), {})
    mm = [m[5:7] for m in months]
    rw(os.path.join(ROOT, 'settlements', 'charged', acc + '.json'),
       lambda old: [x for x in old if x[0][:2] not in mm] + [x for x in o['cr'] if x[0][:2] in mm], [])
    if acc in hub:
        agg = collections.OrderedDict()
        for d, sku, amt, _ in o['sl']:
            k = (d, sku); a = agg.setdefault(k, [d, sku, 0, 0]); a[2] += 1; a[3] += amt
        hubsales[acc] = list(agg.values())
    last = main['p'][-1][0] if main['p'] else '-'
    print(f'{acc}: {len(main["p"])} payment din (aakhri {last}), blob {at}')

if hubsales:
    hs = os.path.join(ROOT, 'hubsales.js')
    t = open(hs, encoding='utf-8').read(); cur = json.loads(t[t.index('{'):].rstrip().rstrip(';'))
    cur.update(hubsales)
    open(hs, 'w', encoding='utf-8').write('window.FK_HUBSALES = ' + json.dumps(cur, separators=(',', ':'), ensure_ascii=False) + ';')
    print('hubsales.js:', {k: len(v) for k, v in hubsales.items()})
