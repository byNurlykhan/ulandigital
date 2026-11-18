"""
Әр жаңалыққа /n/<id>/index.html бетін жасайды.

Мақсаты: WhatsApp, Telegram, Instagram т.б. сілтемені бөліскенде
мақаланың суреті мен тақырыбы көрінсін (Open Graph). Адам сілтемені ашса,
бет оны бірден басты беттегі мақалаға (/#news-<id>) апарады.

Іске қосу:  python scripts/build_news_pages.py
GitHub Actions мұны әр 30 минут сайын автоматты түрде орындайды.
"""
import html
import io
import json
import os
import re
import shutil
import sys
import urllib.request

NEWS_API = ('https://script.google.com/macros/s/AKfycbzxsVDRrzou_ooGE5BlN88F84jFCVtX-'
            'Gc41cDuPRX-wTnGt_GvKzDJb8qB4spiaH-e/exec?action=getNews')
SITE = 'https://ulandigital.kz'
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'n')
IMG_DIR = os.path.join(OUT, 'img')
UA = {'User-Agent': 'ulandigital-news-pages/1.0'}


def fetch(url, timeout=60):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def plain(text, limit=180):
    t = re.sub(r'\[\s*(?:сурет|фото|img|видео|бейне|video)\s*:[^\]]*\]', ' ', text or '', flags=re.I)
    t = re.sub(r'^#+\s*', '', t, flags=re.M)
    t = re.sub(r'\*\*([^*\n]+)\*\*', r'\1', t)
    t = re.sub(r'\s+', ' ', t).replace('—', '–').strip()
    return t if len(t) <= limit else t[:limit].rsplit(' ', 1)[0] + '…'


def save_cover(news_id, image_url):
    """Мұқабаны 1200x630 JPEG етіп сақтайды (превьюге ең қолайлы өлшем)."""
    path = os.path.join(IMG_DIR, f'{news_id}.jpg')
    if os.path.exists(path):
        return True
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(fetch(image_url))).convert('RGB')
        tw, th = 1200, 630
        scale = max(tw / img.width, th / img.height)
        img = img.resize((round(img.width * scale), round(img.height * scale)), Image.LANCZOS)
        left, top = (img.width - tw) // 2, (img.height - th) // 2
        img.crop((left, top, left + tw, top + th)).save(path, 'JPEG', quality=82, optimize=True, progressive=True)
        return True
    except Exception as e:  # сурет жүктелмесе – жалпы превью қолданылады
        print(f'  ! {news_id}: сурет жүктелмеді ({e})', file=sys.stderr)
        return False


PAGE = """<!DOCTYPE html>
<html lang="kk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} – «Ұлан»</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{url}">
<meta property="og:type" content="article">
<meta property="og:site_name" content="«Ұлан» digital">
<meta property="og:locale" content="kk_KZ">
<meta property="og:url" content="{url}">
<meta property="og:title" content="{title}">
<meta property="og:description" content="{desc}">
<meta property="og:image" content="{image}">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="{title}">
<meta name="twitter:description" content="{desc}">
<meta name="twitter:image" content="{image}">
<meta http-equiv="refresh" content="0; url=/#news-{id}">
<script>location.replace('/#news-{id}');</script>
</head>
<body style="font-family:system-ui,sans-serif;background:#F5F3FB;color:#1C1633;padding:40px">
<p><a href="/#news-{id}">{title} →</a></p>
</body>
</html>
"""


def main():
    data = json.loads(fetch(NEWS_API).decode('utf-8'))
    news = data.get('news') or []
    if not data.get('ok') or not news:
        print('Жаңалықтар алынбады – ештеңе өзгертілмеді.')
        return
    os.makedirs(IMG_DIR, exist_ok=True)
    keep = {'img'}
    for item in news:
        nid = re.sub(r'[^A-Za-z0-9_-]', '', str(item.get('id', '')))
        if not nid:
            continue
        keep.add(nid)
        has_img = bool(item.get('imageUrl')) and save_cover(nid, item['imageUrl'])
        page = PAGE.format(
            id=nid,
            url=f'{SITE}/n/{nid}/',
            title=html.escape((item.get('title') or '«Ұлан» жаңалығы').replace('—', '–').strip()),
            desc=html.escape(plain(item.get('body'))),
            image=f'{SITE}/n/img/{nid}.jpg' if has_img else f'{SITE}/share.jpg',
        )
        os.makedirs(os.path.join(OUT, nid), exist_ok=True)
        with open(os.path.join(OUT, nid, 'index.html'), 'w', encoding='utf-8', newline='\n') as f:
            f.write(page)
    # Өшірілген жаңалықтардың беттері мен суреттерін тазалау
    for name in os.listdir(OUT):
        if name not in keep and os.path.isdir(os.path.join(OUT, name)):
            shutil.rmtree(os.path.join(OUT, name))
    for name in os.listdir(IMG_DIR):
        if name.rsplit('.', 1)[0] not in keep:
            os.remove(os.path.join(IMG_DIR, name))
    print(f'Дайын: {len(keep) - 1} жаңалық беті.')


if __name__ == '__main__':
    main()
