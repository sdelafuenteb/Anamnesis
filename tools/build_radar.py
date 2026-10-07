import sys, json, re, datetime, openpyxl
xlsx, tpl, out = sys.argv[1:4]
ws = openpyxl.load_workbook(xlsx)['tabla_frecuencia']
rows = list(ws.iter_rows(values_only=True))
comps = [c for c in rows[0][2:-1]]
data = []
for r in rows[1:]:
    if not r[0] or not r[1]: continue
    data.append({"cat": r[0], "label": r[1], "c": [comps[i] for i, v in enumerate(r[2:-1]) if v == '✓']})
html = open(tpl, encoding='utf8').read()
# strip skeleton
body = html.split('</head><body>', 1)[1].rsplit('</body></html>', 1)[0]
body = body.replace('<title>Radar de Repetición Retail</title>\n', '')
body = '<title>Radar de Repetición Anamnesis</title>\n' + body
body = re.sub(r'const DATA = .*?;\n\(function', lambda m: 'const DATA = ' + json.dumps({"competitors": comps, "rows": data}, ensure_ascii=False) + ';\n(function', body, count=1, flags=re.S)
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-3))).strftime('%d-%m-%Y %H:%M')
body = body.replace('Benchmark competitivo · Retail', 'Benchmark competitivo · Archivo personal de salud')
body = body.replace('Cada punta del radar es una funcionalidad. Mientras más lejos del centro, más competidores del panel la ofrecen. Sirve para distinguir lo que ya es estándar de la industria de lo que solo hace un jugador.',
 'Cada punta del radar es una funcionalidad. Mientras más lejos del centro, más competidores del panel la ofrecen. Sirve para distinguir lo que ya es estándar entre las apps de historial de salud personal de lo que solo hace un jugador.')
body = re.sub(r'Fuente: Benchmark_Retail\.xlsx, pestaña tabla_frecuencia · Actualizado [^<]*', 'Fuente: Benchmark_Anamnesis.xlsx, pestaña tabla_frecuencia · Actualizado ' + now + ' (hora Chile)', body)
body = body.replace('Datos del benchmark de competidores POS/ERP.', 'Datos del benchmark de competidores de archivo personal de salud.')
open(out, 'w', encoding='utf8').write(body)
print(len(data), 'funcionalidades;', len(comps), 'competidores;', len(body), 'bytes')
