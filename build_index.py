#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dashboard.src.html -> index.html 재생성 스크립트.

index.html은 gzip+base64로 압축한 대시보드 본문을 런타임에 풀어서 document.write 하는
구조입니다. 사람이 읽고 리뷰하는 원본은 dashboard.src.html 하나뿐이고,
index.html은 이 스크립트로만 만듭니다. 손으로 고치지 마세요.

    python3 build_index.py
"""
import base64
import gzip
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, 'dashboard.src.html')
OUT = os.path.join(ROOT, 'index.html')
PATTERN = re.compile(r"(const payload = ')([A-Za-z0-9+/=]+)(')")


def main():
    src = io.open(SRC, encoding='utf-8').read()
    shell = io.open(OUT, encoding='utf-8').read()
    if not PATTERN.search(shell):
        sys.exit('index.html에서 payload 상수를 찾지 못했습니다.')

    buf = io.BytesIO()
    with gzip.GzipFile(fileobj=buf, mode='wb', compresslevel=9, mtime=0) as gz:
        gz.write(src.encode('utf-8'))
    payload = base64.b64encode(buf.getvalue()).decode('ascii')

    built = PATTERN.sub(lambda m: m.group(1) + payload + m.group(3), shell, count=1)
    io.open(OUT, 'w', encoding='utf-8').write(built)
    print('src %d bytes -> payload %d chars -> %s' % (len(src), len(payload), os.path.basename(OUT)))


if __name__ == '__main__':
    main()
