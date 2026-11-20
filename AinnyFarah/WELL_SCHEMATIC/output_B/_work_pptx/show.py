import json,sys
for k in sys.argv[1:]:
    d=json.load(open(f'dry/Dulang_{k}.json'))
    print('=====',k, json.dumps(d['wellInfo']))
    for c in d['components']:
        print(f"{c['id'][:22]:<22} {c['type'][:15]:<15} {c.get('string','') or '':<12} {str(c['topMdM']):>8} {str(c['bottomMdM']):>8} ID {str(c['innerDiameter']):>6} OD {str(c['outerDiameter']):>5} {'E:'+','.join(c.get('estimatedFields',[])) if c.get('estimated') else '':<22} {c['name'][:60]!r}")
