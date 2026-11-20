import json,sys
for k in sys.argv[1:]:
    d=json.load(open(f'dry/Dulang_{k}.json'))
    print('=====',k, json.dumps(d['wellInfo'])[:600])
    for c in d['components']:
        print(f"{c['id'][:24]:<24} {c['type'][:15]:<15} {c.get('string','') or '':<12} {str(c['topMdM']):>8} {str(c['bottomMdM']):>8} ID {str(c['innerDiameter']):>6} OD {str(c['outerDiameter']):>6} {'E:'+','.join(c.get('estimatedFields',[])) if c.get('estimated') else '':<26} {c['name'][:55]!r}")
    print('float',d['floatEquipment'],'pbc',d['plugBackCement'])
