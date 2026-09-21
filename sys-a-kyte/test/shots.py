import re, os, pathlib, threading, functools
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from playwright.sync_api import sync_playwright
SP=str(pathlib.Path(__file__).resolve().parent.parent)  # raíz del proyecto
SRC=os.environ.get('SYS_XLSX') or f'{SP}/datos/INVENTARIO_KINO_18_de_sep.xlsx'
os.makedirs(f'{SP}/test/out',exist_ok=True)
class Q(SimpleHTTPRequestHandler):
    def log_message(self,*a): pass
srv=ThreadingHTTPServer(('127.0.0.1',0),functools.partial(Q,directory=f'{SP}/app'))
port=srv.server_address[1]; threading.Thread(target=srv.serve_forever,daemon=True).start()
def go(p,w,h,name,upload,dark=False):
    b=p.chromium.launch(args=['--no-sandbox'])
    ctx=b.new_context(viewport={'width':w,'height':h},color_scheme='dark' if dark else 'light')
    ctx.route('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',lambda r:r.fulfill(path=f'{SP}/test/node_modules/xlsx/dist/xlsx.full.min.js',content_type='application/javascript'))
    ctx.route('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',lambda r:r.fulfill(path=f'{SP}/test/node_modules/jszip/dist/jszip.min.js',content_type='application/javascript'))
    ctx.route(re.compile(r'https://fonts\.(googleapis|gstatic)\.com/.*'),lambda r:r.abort())
    pg=ctx.new_page(); pg.goto(f'http://127.0.0.1:{port}/index.html'); pg.wait_for_selector('#chain .node')
    if upload:
        pg.set_input_files('#f-sys',SRC); pg.wait_for_selector('.node.sys.on',timeout=30000); pg.wait_for_selector('#lista table')
    pg.evaluate('window.scrollTo(0,0)')
    pg.screenshot(path=f'{SP}/test/out/{name}.png'); b.close()
with sync_playwright() as p:
    go(p,1200,900,'top_inicio',False)
    go(p,1200,900,'top_resultado',True)
    go(p,400,860,'mobile_resultado',True)
    go(p,1000,760,'dark_resultado',True,dark=True)
print('ok')
