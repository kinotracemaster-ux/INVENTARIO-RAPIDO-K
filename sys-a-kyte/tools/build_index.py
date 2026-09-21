#!/usr/bin/env python3
"""Arma app/index.html (versión autónoma) a partir de app/sys-a-kyte.html.

El Artifact de Claude envuelve el fragmento con su propio esqueleto al publicar.
Para abrir la app fuera de Claude (o probarla con Playwright) hay que envolverlo
con un doctype/head/body. Este script hace solo eso.

Uso:  python3 tools/build_index.py
Luego abrir app/index.html en el navegador (tabla.js debe estar al lado).
"""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
frag = (ROOT / 'app' / 'sys-a-kyte.html').read_text(encoding='utf-8')
i = frag.index('<main')
head, body = frag[:i], frag[i:]

# Reset mínimo equivalente al esqueleto que agrega la plataforma
reset = ('<style>:root{color-scheme:light dark}body{margin:0;font:14px system-ui}'
         'img{max-width:100%}[hidden]{display:none!important}</style>')
doc = ('<!doctype html><html lang="es"><head><meta charset="utf-8">'
       '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
       + reset + head + '</head><body>' + body + '</body></html>')
(ROOT / 'app' / 'index.html').write_text(doc, encoding='utf-8')
print('app/index.html listo', len(doc), 'bytes')
