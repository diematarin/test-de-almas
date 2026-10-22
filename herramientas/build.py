import json,base64,io,glob,os,unicodedata
from PIL import Image
src='/mnt/user-data/uploads/APK test/Imagenes'
out='/mnt/project-files/test-camino-de-vida/web/img'; os.makedirs(out,exist_ok=True)
def norm(s): return ''.join(c for c in unicodedata.normalize('NFD',s) if unicodedata.category(c)!='Mn').lower()
def enc(path,w,q,name):
  im=Image.open(path).convert('RGB'); h=round(im.height*w/im.width); im=im.resize((w,h),Image.LANCZOS)
  b=io.BytesIO(); im.save(b,'WEBP',quality=q); im.save(os.path.join(out,name),'WEBP',quality=q)
  return 'data:image/webp;base64,'+base64.b64encode(b.getvalue()).decode(), len(b.getvalue())
cards={};tot=0
for p in sorted(glob.glob(src+'/Tarot/*.png')):
  key=norm(os.path.basename(p)[:-4]); uri,n=enc(p,300,72,key.replace(' ','-')+'.webp'); cards[key]=uri; tot+=n
cover,n=enc(src+'/portada.jpg',900,72,'portada.webp'); tot+=n
data=json.load(open('/mnt/project-files/test-camino-de-vida/contenido-test-v2.json'))
for a in data['arquetipos']:
  k=norm(a['carta']).replace('ermitano','ermitanio')
  assert k in cards,(k,list(cards))
t=open('template.html').read()
t=t.replace('__DATA__',json.dumps(data,ensure_ascii=False)).replace('__CARDS__',json.dumps(cards)).replace('__COVER__',json.dumps(cover)).replace('__PAGOS__','{"endpoint": "", "demo": true}')
open('test-de-alma.html','w').write(t); print('imgs KB',tot//1024,'page KB',len(t.encode())//1024)
