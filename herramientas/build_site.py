# Genera la versión para publicar (GitHub Pages): imágenes como archivos, no incrustadas.
import json,os,unicodedata,shutil
base='/mnt/project-files/test-camino-de-vida'
out=base+'/sitio'; os.makedirs(out,exist_ok=True)
shutil.copytree(base+'/web/img',out+'/img',dirs_exist_ok=True)
cards={f[:-5].replace('-',' '):'img/'+f for f in os.listdir(out+'/img') if f!='portada.webp'}
data=json.load(open(base+'/contenido-test-v2.json'))
# Los textos completos no van en la web: los entrega el script de Google tras el pago.
import re
for a in data['arquetipos']:
  teaser=''
  for s in re.findall(r'[^.!?]+[.!?]+["”]?\s*', a['texto']):
    if len(teaser)>420: break
    teaser+=s
  a['teaser']=teaser; del a['texto']; del a['mensaje_extendido']
data['mensaje_duelo']=''
pagos=json.load(open(base+'/sitio-pagos.json')) if os.path.exists(base+'/sitio-pagos.json') else {'endpoint':''}
t=open(os.path.dirname(os.path.abspath(__file__))+'/template.html').read()
t=t.replace('__DATA__',json.dumps(data,ensure_ascii=False)).replace('__CARDS__',json.dumps(cards,ensure_ascii=False)).replace('__COVER__','"img/portada.webp"').replace('__PAGOS__',json.dumps(pagos))
head='<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n<meta name="description" content="Descubrí tu arquetipo del alma y la carta del tarot que te acompaña.">\n<meta property="og:title" content="Test de Alma">\n<meta property="og:image" content="img/portada.webp">\n'
t=t.replace('<meta name="claude:viewport-fit-sides" content="cover">\n','')
t=head+t.replace('<div class="wrap"','</head>\n<body>\n<div class="wrap"',1)+'\n</body>\n</html>\n'
# mover <title>/<link>/<style> dentro del head ya ocurre porque están antes del </head> insertado
open(out+'/index.html','w').write(t); print(sorted(cards)[:3], len(cards))
