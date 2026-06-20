import os
import json
import re
import xml.etree.ElementTree as ET
import urllib.request
import urllib.parse
import zoneinfo
from datetime import datetime, timedelta, timezone

ANTHROPIC_API_KEY = os.environ['ANTHROPIC_API_KEY']
MS_CLIENT_ID = os.environ['MS_CLIENT_ID']
MS_REFRESH_TOKEN = os.environ['MS_REFRESH_TOKEN']
GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
WARSAW_TZ = zoneinfo.ZoneInfo('Europe/Warsaw')

BB_LAT, BB_LON = 49.8224, 19.0584
KETY_LAT, KETY_LON = 49.8825, 19.2216


def refresh_access_token():
    data = urllib.parse.urlencode({
        'grant_type': 'refresh_token',
        'client_id': MS_CLIENT_ID,
        'refresh_token': MS_REFRESH_TOKEN,
        'scope': 'Mail.Read Mail.Send Tasks.Read User.Read Calendars.Read offline_access',
    }).encode()
    req = urllib.request.Request(
        'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
        data=data,
        headers={'Content-Type': 'application/x-www-form-urlencoded'},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())['access_token']


def graph_get(token, path, params=None):
    url = path if path.startswith('http') else f'{GRAPH_BASE}{path}'
    if params:
        url += '?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(
        url,
        headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except Exception as e:
        return {'error': str(e), 'value': []}


def fetch_emails(token):
    since = (datetime.now(timezone.utc) - timedelta(days=3)).strftime('%Y-%m-%dT%H:%M:%SZ')
    result = graph_get(token, '/me/mailFolders/inbox/messages', {
        '$top': 50,
        '$select': 'subject,from,receivedDateTime,isRead,bodyPreview',
        '$filter': f"receivedDateTime ge {since}",
        '$orderby': 'receivedDateTime desc',
    })
    messages = result.get('value', [])
    if not messages:
        return 'Brak wiadomosci z ostatnich 3 dni.'
    lines = []
    for m in messages:
        status = '[NOWE] ' if not m.get('isRead') else ''
        sender = m.get('from', {}).get('emailAddress', {}).get('address', '?')
        lines.append(
            f"{status}Od: {sender}\n"
            f"Temat: {m.get('subject', '(brak tematu)')}\n"
            f"Data: {m.get('receivedDateTime', '')[:10]}\n"
            f"Podglad: {m.get('bodyPreview', '')[:400]}\n"
        )
    return '\n'.join(lines)


def fetch_todo_tasks(token):
    lists_result = graph_get(token, '/me/todo/lists')
    task_lines = []
    for lst in lists_result.get('value', []):
        list_name = lst['displayName']
        if list_name.strip().lower() != 'do zrobienia':
            continue
        list_id = lst['id']
        url = f"/me/todo/lists/{list_id}/tasks"
        params = {'$filter': "status ne 'completed'", '$top': 100}
        while url:
            result = graph_get(token, url, params if params else None)
            params = None
            for task in result.get('value', []):
                due = task.get('dueDateTime')
                due_str = f" [termin: {due['dateTime'][:10]}]" if due else ''
                task_lines.append(f"- {task['title']}{due_str}")
            url = result.get('@odata.nextLink') or None
    return '\n'.join(task_lines) if task_lines else 'Brak aktywnych zadan.'


# ── Kalendarz: polskie dni specjalne ─────────────────────────────────────────

def _easter(year):
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    from datetime import date
    return date(year, month, day)


def get_polish_special_days(d):
    """Zwraca liste (typ, nazwa) specjalnych dni dla daty d."""
    from datetime import date, timedelta
    year, month, day = d.year, d.month, d.day
    today = date(year, month, day)
    results = []

    # Stale dni wolne od pracy
    fixed = {
        (1, 1): 'Nowy Rok',
        (6, 1): 'Trzech Króli',
        (1, 5): 'Święto Pracy',
        (3, 5): 'Konstytucja 3 Maja',
        (15, 8): 'Wniebowzięcie NMP / Święto Wojska Polskiego',
        (1, 11): 'Wszystkich Świętych',
        (11, 11): 'Dzień Niepodległości',
        (25, 12): 'Boże Narodzenie (1. dzień)',
        (26, 12): 'Boże Narodzenie (2. dzień)',
    }
    if (day, month) in fixed:
        results.append(('DZIEN_WOLNY', fixed[(day, month)]))

    # Ruchome swieta
    easter = _easter(year)
    moving = {
        easter:                  ('DZIEN_WOLNY', 'Niedziela Wielkanocna'),
        easter + timedelta(1):   ('DZIEN_WOLNY', 'Poniedziałek Wielkanocny'),
        easter + timedelta(49):  ('DZIEN_WOLNY', 'Zielone Świątki'),
        easter + timedelta(60):  ('DZIEN_WOLNY', 'Boże Ciało'),
    }
    if today in moving:
        results.append(moving[today])

    # Nieoficjalne dni specjalne
    unofficial = {
        (21, 1):  'Dzień Babci',
        (22, 1):  'Dzień Dziadka',
        (14, 2):  'Walentynki',
        (8, 3):   'Dzień Kobiet',
        (21, 3):  'Pierwszy Dzień Wiosny',
        (26, 5):  'Dzień Matki',
        (1, 6):   'Dzień Dziecka',
        (23, 6):  'Dzień Ojca',
        (14, 10): 'Dzień Edukacji Narodowej',
        (1, 11):  'Zaduszki',
        (11, 11): 'Dzień Niepodległości',
    }
    if (day, month) in unofficial and ('DZIEN_WOLNY', unofficial[(day, month)]) not in results:
        results.append(('DZIEN_SPECJALNY', unofficial[(day, month)]))

    # Jutro dzien wolny — ostrzezenie wieczorne
    tomorrow = today + timedelta(1)
    for key, name in fixed.items():
        if key == (tomorrow.day, tomorrow.month):
            results.append(('JUTRO_WOLNE', name))
            break
    for move_date, (_, name) in moving.items():
        if move_date == tomorrow:
            results.append(('JUTRO_WOLNE', name))
            break

    return results


def fmt_special_days(special_days):
    """Formatuje liste dni specjalnych do stringa dla promptu."""
    if not special_days:
        return 'Brak'
    lines = []
    for typ, name in special_days:
        if typ == 'DZIEN_WOLNY':
            lines.append(f'DZIEN WOLNY OD PRACY: {name}')
        elif typ == 'DZIEN_SPECJALNY':
            lines.append(f'DZIEN SPECJALNY: {name}')
        elif typ == 'JUTRO_WOLNE':
            lines.append(f'JUTRO DZIEN WOLNY: {name}')
    return '\n'.join(lines)


# ── Kalendarz: Graph API ──────────────────────────────────────────────────────

def fetch_calendar_events(token):
    """Pobiera zdarzenia z Outlook Calendar na dzisiaj."""
    now_local = datetime.now(WARSAW_TZ)
    today_start = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    today_end   = now_local.replace(hour=23, minute=59, second=59, microsecond=0)

    url = (
        f"{GRAPH_BASE}/me/calendarView"
        f"?startDateTime={today_start.strftime('%Y-%m-%dT%H:%M:%S')}"
        f"&endDateTime={today_end.strftime('%Y-%m-%dT%H:%M:%S')}"
        f"&$select=subject,start,end,location,isAllDay,importance,bodyPreview"
        f"&$orderby=start/dateTime&$top=50"
    )
    req = urllib.request.Request(url, headers={
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json',
        'Prefer': 'outlook.timezone="Central European Standard Time"',
    })
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            data = json.loads(r.read())
    except Exception as e:
        return f'[Blad pobierania kalendarza: {e}]'

    events = data.get('value', [])
    if not events:
        return 'Brak wydarzen w kalendarzu na dzisiaj.'

    lines = []
    for e in events:
        subject = e.get('subject', '(brak tytulu)')
        is_all_day = e.get('isAllDay', False)
        importance = e.get('importance', 'normal')
        location = (e.get('location') or {}).get('displayName', '')
        preview = (e.get('bodyPreview') or '')[:120]

        if is_all_day:
            time_str = '[cały dzień]'
        else:
            try:
                s = e['start']['dateTime'][:16].replace('T', ' ')
                en = e['end']['dateTime'][11:16]
                time_str = f"{s[-5:]}-{en}"
            except Exception:
                time_str = '?'

        imp = ' [!WAŻNE]' if importance == 'high' else ''
        loc = f' @ {location}' if location else ''
        lines.append(f"{time_str}{imp}: {subject}{loc}")
        if preview and preview.strip():
            lines.append(f"  {preview.strip()}")

    return '\n'.join(lines)


def fetch_url(url, timeout=30, extra_headers=None):
    headers = {'User-Agent': 'Mozilla/5.0'}
    if extra_headers:
        headers.update(extra_headers)
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read().decode('utf-8', errors='replace')
    except Exception as e:
        return f'[Blad: {e}]'


def fetch_json(url, timeout=30, extra_headers=None):
    raw = fetch_url(url, timeout=timeout, extra_headers=extra_headers)
    try:
        return json.loads(raw)
    except Exception:
        return None


# ── wttr.in ───────────────────────────────────────────────────────────────────

def fetch_wttr(city, hour_start, hour_end):
    data = fetch_json(f'https://wttr.in/{urllib.parse.quote(city)}?format=j1')
    if not data:
        return None
    try:
        today = data['weather'][0]
        astro = today.get('astronomy', [{}])[0]
        temps, feels, precip_mm_list, precip_prob_list, winds, rain_hours = [], [], [], [], [], []
        for h in today['hourly']:
            t = int(h['time'])
            hour_h = t // 100
            if hour_start <= hour_h <= hour_end:
                temps.append(int(h['tempC']))
                feels.append(int(h['FeelsLikeC']))
                mm = float(h['precipMM'])
                prob = int(h.get('chanceofrain', 0))
                precip_mm_list.append(mm)
                precip_prob_list.append(prob)
                winds.append(int(h['windspeedKmph']))
                if prob >= 30 or mm >= 0.3:
                    rain_hours.append(f"{hour_h:02d}:00")
        if not temps:
            return None
        return {
            'source': 'wttr.in',
            'min_temp': min(temps), 'max_temp': max(temps),
            'min_feels': min(feels), 'max_feels': max(feels),
            'max_precip_mm': max(precip_mm_list),
            'total_precip_mm': round(sum(precip_mm_list), 1),
            'max_precip_prob': max(precip_prob_list),
            'max_wind': max(winds),
            'rain_hours': rain_hours,
            'sunrise': astro.get('sunrise', '?'),
            'sunset': astro.get('sunset', '?'),
        }
    except Exception:
        return None


# ── Open-Meteo ────────────────────────────────────────────────────────────────

def fetch_openmeteo(lat, lon, hour_start, hour_end):
    url = (
        f"https://api.open-meteo.com/v1/forecast"
        f"?latitude={lat}&longitude={lon}"
        f"&hourly=temperature_2m,apparent_temperature,precipitation_probability,"
        f"precipitation,windspeed_10m"
        f"&timezone=Europe%2FWarsaw&forecast_days=1"
    )
    data = fetch_json(url)
    if not data:
        return None
    try:
        hourly = data['hourly']
        times = hourly['time']
        temps, feels, precip_mm_list, precip_prob_list, winds, rain_hours = [], [], [], [], [], []
        for i, t in enumerate(times):
            hour = int(t[11:13])
            if hour_start <= hour <= hour_end:
                temps.append(round(hourly['temperature_2m'][i]))
                feels.append(round(hourly['apparent_temperature'][i]))
                mm = hourly['precipitation'][i]
                prob = hourly['precipitation_probability'][i]
                winds.append(round(hourly['windspeed_10m'][i]))
                precip_mm_list.append(mm)
                precip_prob_list.append(prob)
                if prob >= 30 or mm >= 0.3:
                    rain_hours.append(f"{hour:02d}:00")
        if not temps:
            return None
        return {
            'source': 'open-meteo.com',
            'min_temp': min(temps), 'max_temp': max(temps),
            'min_feels': min(feels), 'max_feels': max(feels),
            'max_precip_mm': max(precip_mm_list),
            'total_precip_mm': round(sum(precip_mm_list), 1),
            'max_precip_prob': max(precip_prob_list),
            'max_wind': max(winds),
            'rain_hours': rain_hours,
        }
    except Exception:
        return None


# ── yr.no ─────────────────────────────────────────────────────────────────────

def fetch_yr(lat, lon, hour_start, hour_end):
    url = f"https://api.met.no/weatherapi/locationforecast/2.0/compact?lat={lat}&lon={lon}"
    data = fetch_json(url, extra_headers={'User-Agent': 'morning-briefing/1.0 wrathtide@outlook.com'})
    if not data:
        return None
    try:
        today_local = datetime.now(WARSAW_TZ).date()
        temps, winds, precip_mm_list, precip_prob_list, rain_hours = [], [], [], [], []
        for entry in data['properties']['timeseries']:
            t_utc = datetime.strptime(entry['time'], '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc)
            t_local = t_utc.astimezone(WARSAW_TZ)
            if t_local.date() != today_local:
                continue
            h = t_local.hour
            if not (hour_start <= h <= hour_end):
                continue
            details = entry['data']['instant']['details']
            temps.append(round(details.get('air_temperature', 0)))
            winds.append(round(details.get('wind_speed', 0) * 3.6))
            next1 = entry['data'].get('next_1_hours', {}).get('details', {})
            mm = next1.get('precipitation_amount', 0) or 0
            prob = next1.get('probability_of_precipitation', 0) or 0
            precip_mm_list.append(mm)
            precip_prob_list.append(prob)
            if mm >= 0.3 or prob >= 30:
                rain_hours.append(f"{h:02d}:00")
        if not temps:
            return None
        return {
            'source': 'yr.no',
            'min_temp': min(temps), 'max_temp': max(temps),
            'max_precip_mm': max(precip_mm_list),
            'total_precip_mm': round(sum(precip_mm_list), 1),
            'max_precip_prob': max(precip_prob_list),
            'max_wind': max(winds),
            'rain_hours': rain_hours,
        }
    except Exception:
        return None


# ── IMGW ──────────────────────────────────────────────────────────────────────

def fetch_imgw_synop(station_keyword):
    data = fetch_json('https://danepubliczne.imgw.pl/api/data/synop')
    if not isinstance(data, list):
        return None
    try:
        station = next(
            (s for s in data if station_keyword.upper() in s.get('stacja', '').upper()),
            None
        )
        if not station:
            return None
        wind_ms = float(station.get('predkosc_wiatru') or 0)
        return {
            'source': 'IMGW',
            'station': station.get('stacja', '?'),
            'time': f"{station.get('data_pomiaru', '')} {station.get('godzina_pomiaru', '')}:00",
            'temp': station.get('temperatura', '?'),
            'wind_kmh': round(wind_ms * 3.6, 1),
            'precip_1h': station.get('suma_opadu', '?'),
            'pressure': station.get('cisnienie', '?'),
            'humidity': station.get('wilgotnosc_wzgledna', '?'),
        }
    except Exception:
        return None


def fetch_imgw_warnings():
    data = fetch_json('https://danepubliczne.imgw.pl/api/data/warningsmeteo')
    if not isinstance(data, list):
        return ''
    try:
        lines = []
        for w in data:
            zjawisko = w.get('nazwa_zdarzenia', '?')
            stopien  = w.get('stopien', '?')
            prob     = w.get('prawdopodobienstwo', '')
            od       = w.get('obowiazuje_od', '')[:16]
            do       = w.get('obowiazuje_do', '')[:16]
            tresc    = w.get('tresc', '')[:200]
            lines.append(f"STOPIEŃ {stopien}: {zjawisko} (prawdop.{prob}%, {od}–{do}) | {tresc}")
        return '\n'.join(lines)
    except Exception:
        return ''


def fetch_epic_free():
    url = ('https://store-site-backend-static-ipv4.ak.epicgames.com/'
           'freeGamesPromotions?locale=pl&country=PL&allowCountries=PL')
    data = fetch_json(url)
    games = (data or {}).get('data', {}).get('Catalog', {}).get('searchStore', {}).get('elements', [])
    free = []
    for g in games:
        promo = g.get('promotions') or {}
        for block in promo.get('promotionalOffers', []):
            if not isinstance(block, dict):
                continue
            for offer in block.get('promotionalOffers', []):
                if offer.get('discountSetting', {}).get('discountPercentage', -1) == 0:
                    mappings = (g.get('catalogNs') or {}).get('mappings') or [{}]
                    slug = (mappings[0] or {}).get('pageSlug', '')
                    link = (f'https://store.epicgames.com/pl/p/{slug}'
                            if slug else 'https://store.epicgames.com/pl/free-games')
                    end = offer.get('endDate', '')[:10]
                    free.append({'title': g.get('title', '?'), 'link': link, 'end': end})
    return free


def fetch_gog_free():
    url = ('https://catalog.gog.com/v1/catalog?limit=20&price=between:0,0'
           '&order=trending-desc&productType=in:game')
    data = fetch_json(url)
    products = (data or {}).get('products', [])
    result = []
    for p in products:
        title = p.get('title', '')
        if 'demo' in title.lower() or 'demo' in p.get('slug', ''):
            continue
        slug = p.get('slug', '')
        link = f'https://www.gog.com/game/{slug}' if slug else 'https://www.gog.com'
        result.append({'title': title, 'link': link})
        if len(result) >= 8:
            break
    return result


# ── Porownanie zrodel ─────────────────────────────────────────────────────────

def compare_sources(sources):
    valid = [s for s in sources if s and s.get('max_temp') is not None]
    if len(valid) < 2:
        return ''
    warnings = []
    temp_pairs = [(s['source'], s['max_temp']) for s in valid]
    temp_vals = [v for _, v in temp_pairs]
    if max(temp_vals) - min(temp_vals) >= 5:
        desc = ', '.join(f"{src}: {t}°C" for src, t in temp_pairs)
        warnings.append(f"Temperatura maks: roznica {max(temp_vals)-min(temp_vals)}°C ({desc})")
    prec_pairs = [(s['source'], s['max_precip_mm']) for s in valid
                  if s.get('max_precip_mm') is not None]
    if prec_pairs:
        p_vals = [v for _, v in prec_pairs]
        if max(p_vals) - min(p_vals) >= 3:
            desc = ', '.join(f"{src}: {p}mm" for src, p in prec_pairs)
            warnings.append(f"Opady maks: roznica {max(p_vals)-min(p_vals):.1f}mm ({desc})")
    return '\n'.join(warnings)


def fmt_source(src):
    if not src:
        return 'BLAD - brak danych'
    rain = ', '.join(src['rain_hours']) if src.get('rain_hours') else 'brak'
    return (
        f"temp {src.get('min_temp','?')}-{src.get('max_temp','?')}°C "
        f"(odcz. {src.get('min_feels','?')}-{src.get('max_feels','?')}°C), "
        f"opady max {src.get('max_precip_mm', 0):.1f}mm "
        f"(szansa {src.get('max_precip_prob', 0):.0f}%), "
        f"wiatr max {src.get('max_wind', 0)}km/h, "
        f"godz. z ryzykiem: {rain}"
    )


def fmt_imgw(src):
    if not src:
        return 'BLAD - brak danych ze stacji IMGW'
    return (
        f"Stacja {src['station']} | Pomiar: {src['time']} | "
        f"Temp: {src['temp']}°C | Wiatr: {src['wind_kmh']}km/h | "
        f"Opad 1h: {src['precip_1h']}mm | Cisnienie: {src['pressure']}hPa"
    )


# ── Ruch drogowy (OSRM — bez klucza API, czas bez korkow) ────────────────────

def fetch_traffic_osrm(lat1, lon1, lat2, lon2):
    """Szacowany czas jazdy z OSRM (Open Source Routing Machine, bez real-time traffic)."""
    url = (
        f"https://router.project-osrm.org/route/v1/driving/"
        f"{lon1},{lat1};{lon2},{lat2}?overview=false"
    )
    data = fetch_json(url, extra_headers={'User-Agent': 'morning-briefing/1.0'})
    if not data or data.get('code') != 'Ok':
        return None
    try:
        route = data['routes'][0]
        duration_min = round(route['duration'] / 60)
        distance_km = round(route['distance'] / 1000, 1)
        return {'duration_min': duration_min, 'distance_km': distance_km}
    except Exception:
        return None


# ── Newsy ─────────────────────────────────────────────────────────────────────

def fetch_news_rss(query, lang='pl', country='PL', max_items=5):
    encoded = urllib.parse.quote(query)
    url = f'https://news.google.com/rss/search?q={encoded}&hl={lang}&gl={country}&ceid={country}:{lang}'
    return fetch_rss_items(url, max_items=max_items)


def fetch_article(article_url):
    return fetch_url(f'https://r.jina.ai/{article_url}', timeout=45)[:3000]


def fetch_rss_items(url, max_items=6):
    raw = fetch_url(url, timeout=20)
    if not raw or raw.startswith('[Blad'):
        return []
    try:
        root = ET.fromstring(raw)
        items = []
        for item in root.findall('.//item'):
            title = (item.findtext('title') or '').strip()
            link = (item.findtext('link') or '').strip()
            desc = re.sub(r'<[^>]+>', '', (item.findtext('description') or ''))[:300].strip()
            if title:
                items.append({'title': title, 'link': link, 'desc': desc})
        if not items:
            for entry in root.findall('.//{http://www.w3.org/2005/Atom}entry'):
                title = (entry.findtext('{http://www.w3.org/2005/Atom}title') or '').strip()
                link_el = entry.find('{http://www.w3.org/2005/Atom}link')
                link = (link_el.get('href') or '') if link_el is not None else ''
                desc = re.sub(r'<[^>]+>', '', (
                    entry.findtext('{http://www.w3.org/2005/Atom}summary') or
                    entry.findtext('{http://www.w3.org/2005/Atom}content') or ''
                ))[:300].strip()
                if title:
                    items.append({'title': title, 'link': link, 'desc': desc})
        return items[:max_items]
    except Exception:
        return []


def fmt_rss_items(items):
    if not items:
        return '[brak danych]'
    lines = []
    for it in items:
        link_part = f' [{it["link"]}]' if it.get('link') else ''
        desc_part = f' — {it["desc"][:200]}' if it.get('desc') else ''
        lines.append(f'• {it["title"]}{link_part}{desc_part}')
    return '\n'.join(lines)


# ── Gaming watchlist ──────────────────────────────────────────────────────────

WATCHED_GAMES = [
    ('Starfield', 'Bethesda Game Studios', 'RPG open world sci-fi'),
    ('Star Citizen', 'Cloud Imperium Games', 'space sim MMO'),
    ('Elite Dangerous', 'Frontier Developments', 'space sim open world'),
    ('ARMA series', 'Bohemia Interactive', 'mil-sim tactical'),
    ('DayZ', 'Bohemia Interactive', 'survival open world'),
    ('Diablo 4', 'Blizzard Entertainment', 'action RPG'),
    ('Enshrouded', 'Keen Games', 'survival RPG open world'),
    ('Escape from Tarkov', 'Battlestate Games', 'extraction FPS survival'),
    ('Forza Horizon 6', 'Playground Games', 'racing open world'),
    ('Need for Speed series', 'EA/Criterion', 'racing'),
    ('Stoneshard', 'Ink Stains Games', 'RPG roguelike'),
    ('Wartales', 'Shiro Games', 'tactical RPG open world'),
    ('Age of Wonders 4', 'Triumph Studios', '4X strategy'),
    ('Kingdom Come Deliverance series', 'Warhorse Studios', 'RPG open world'),
    ('Heroes of Might and Magic 3/7/8', 'Ubisoft/community', 'turn-based strategy'),
    ('The Elder Scrolls 7', 'Bethesda Game Studios', 'RPG open world'),
    ("Assassin's Creed series", 'Ubisoft', 'action adventure open world'),
    ('Cyberpunk 2077 series', 'CD Projekt Red', 'RPG open world cyberpunk'),
    ('Days Gone series', 'Sony Bend Studio', 'action survival open world'),
    ('Fallout series', 'Bethesda/Obsidian', 'RPG open world post-apo'),
    ('GTA series', 'Rockstar Games', 'action open world'),
    ('Humankind', 'Amplitude Studios', '4X strategy'),
    ('Stellaris', 'Paradox Interactive', '4X grand strategy sci-fi'),
    ('Mafia series', 'Hangar 13/2K', 'action open world crime'),
    ('SnowRunner series', 'Saber Interactive', 'off-road simulation'),
    ('Outward', 'Nine Dots Studio', 'RPG survival co-op'),
    ("No Man's Sky", 'Hello Games', 'space exploration survival'),
    ('RAGE series', 'id Software', 'FPS open world post-apo'),
    ('Ready or Not', 'VOID Interactive', 'tactical FPS'),
    ('Tomb Raider series', 'Crystal Dynamics', 'action adventure'),
    ('Civilization series', 'Firaxis Games', '4X turn-based strategy'),
    ('Sniper Elite series', 'Rebellion', 'tactical FPS stealth'),
    ('Space Engineers series', 'Keen Software House', 'sandbox engineering'),
    ('State of Decay series', 'Undead Labs', 'survival open world zombie'),
    ('The Forest / Sons of the Forest', 'Endnight Games', 'survival horror'),
    ('The Long Dark', 'Hinterland Studio', 'survival'),
    ('The Precinct', 'Fallen Tree Games', 'action open world'),
    ('Wiedźmin / Witcher series', 'CD Projekt Red', 'RPG open world fantasy'),
    ('Zero Sievert', 'CABO Studio', 'extraction shooter'),
    ('Doom series', 'id Software', 'FPS'),
    ('Dying Light series', 'Techland', 'action survival open world'),
    ('Jagged Alliance series', 'Haemimont/THQ Nordic', 'tactical RPG'),
    ('Red Dead Redemption series', 'Rockstar Games', 'action open world western'),
    ('Ghost Recon series', 'Ubisoft', 'tactical FPS open world'),
    ('Far Cry series', 'Ubisoft', 'FPS open world'),
    ('The Division series', 'Ubisoft Massive', 'online RPG shooter'),
    ('The Crew series', 'Ivory Tower/Ubisoft', 'racing open world'),
    ('Manor Lords', 'Slavic Magic', 'city builder medieval strategy'),
    ('Mad Max', 'Avalanche Studios', 'action open world post-apo'),
    ('Medieval Dynasty / Sengoku Dynasty', 'Render Cube', 'survival RPG'),
    ('Northgard', 'Shiro Games', 'RTS strategy Norse'),
    ('Gothic Remake / Risen series', 'Alkimia Interactive', 'RPG open world'),
    ('Stalker 2 / Stalker series', 'GSC Game World', 'FPS survival open world'),
    ('Wasteland series', 'inXile Entertainment', 'RPG post-apo'),
]

_GAME_QUERY_1 = (
    'Starfield OR "Star Citizen" OR "Elite Dangerous" OR DayZ OR Enshrouded '
    'OR "Escape from Tarkov" OR Stoneshard OR Wartales OR "Kingdom Come" '
    'OR "Cyberpunk 2077" OR "Days Gone" OR Stalker2 OR "Manor Lords"'
)
_GAME_QUERY_2 = (
    '"Assassin\'s Creed" OR Witcher OR Wiedźmin OR Fallout OR GTA OR Stellaris '
    'OR "Dying Light" OR "Ghost Recon" OR "Far Cry" OR "The Division" OR Northgard '
    'OR "Medieval Dynasty" OR "Space Engineers" OR SnowRunner OR "No Man\'s Sky"'
)
_GAME_QUERY_3 = (
    'ARMA OR "Diablo 4" OR "Age of Wonders" OR "Heroes of Might" OR "Elder Scrolls" '
    'OR "Red Dead Redemption" OR "Tomb Raider" OR Civilization OR "Ready or Not" '
    'OR "Doom Eternal" OR "Jagged Alliance" OR "State of Decay" OR "The Crew" '
    'OR Humankind OR Mafia OR Wasteland OR Northgard OR "Gothic Remake"'
)


def fetch_watched_games_news():
    """Pobiera newsy o obserwowanych grach z Google News (3 zapytania)."""
    items = []
    for q in [_GAME_QUERY_1, _GAME_QUERY_2, _GAME_QUERY_3]:
        items += fetch_news_rss(q, max_items=15)
    seen = set()
    unique = []
    for it in items:
        key = it['title'][:60]
        if key not in seen:
            seen.add(key)
            unique.append(it)
    return unique


def select_and_summarize_gaming(gaming_items):
    """Claude wybiera 15 najlepszych newsów gamingowych i pisze streszczenia."""
    if not gaming_items:
        return []
    lines = []
    for i, it in enumerate(gaming_items):
        lines.append(f'{i+1}. {it["title"]}')

    prompt = (
        'Jesteś redaktorem serwisu gamingowego. Przejrzyj poniższe artykuły i wybierz '
        'DOKŁADNIE 15 (lub wszystkie jeśli jest mniej) najciekawszych dla gracza PC.\n'
        'Priorytet: nowe gry/zapowiedzi, DLC, patche, eventy, duże aktualizacje, recenzje.\n'
        'Pomiń: clickbait, powtórzenia tego samego tematu (zostaw najlepszy), plotki bez źródła.\n\n'
        'Format (zachowaj oryginalny numer + 1 zdanie streszczenie po polsku max 15 słów):\n'
        '3. Zdanie streszczające.\n'
        '7. Zdanie streszczające.\n\n'
        'Artykuły:\n' + '\n'.join(lines)
    )
    result = call_claude(prompt, max_tokens=2000, timeout=90)

    selected = []
    for line in result.split('\n'):
        m = re.match(r'^(\d+)\.\s+(.+)', line.strip())
        if m:
            idx = int(m.group(1)) - 1
            summary = m.group(2).strip()
            if 0 <= idx < len(gaming_items):
                selected.append((gaming_items[idx], summary))
    return selected


def select_watched_games_updates(items):
    """Claude filtruje i grupuje newsy o obserwowanych grach po grach."""
    if not items:
        return {}
    game_names = ', '.join(g[0] for g in WATCHED_GAMES)
    lines = [f'{i+1}. {it["title"]}' for i, it in enumerate(items)]

    prompt = (
        'Filtruj poniższe newsy. Zostaw TYLKO te które dotyczą co najmniej jednej z gier:\n'
        f'{game_names}\n\n'
        'Uwzględniaj: patche, DLC, eventy, darmowe przedmioty, zapowiedzi, bety, '
        'aktualizacje, nowe gry z serii, zamknięcia studiów. Pomiń niezwiązane z tymi grami.\n\n'
        'Format — artykuły pogrupowane po grach:\n'
        '=== Nazwa gry ===\n'
        'N. Jedno zdanie po polsku (co nowego/co się zmieniło).\n'
        '  Jeśli artykuł dotyczy patcha/aktualizacji — dodaj konkretne zmiany:\n'
        '  • Zmiana 1 (np. "Dodano zimowe mapy Namalsk")\n'
        '  • Zmiana 2 (np. "Poprawiono synchronizację ekwipunku")\n'
        '  Jeśli NIE jest to patch — tylko zdanie, BEZ punktów.\n\n'
        'Artykuły:\n' + '\n'.join(lines)
    )
    result = call_claude(prompt, max_tokens=3000, timeout=90)

    grouped = {}
    current_game = None
    current_article = None

    for line in result.split('\n'):
        m_game = re.match(r'^===\s*(.+?)\s*===', line.strip())
        if m_game:
            current_game = m_game.group(1).strip()
            if current_game not in grouped:
                grouped[current_game] = []
            current_article = None
            continue

        m_art = re.match(r'^(\d+)\.\s+(.+)', line.strip())
        if m_art and current_game is not None:
            idx = int(m_art.group(1)) - 1
            if 0 <= idx < len(items):
                current_article = {
                    'item': items[idx],
                    'summary': m_art.group(2).strip(),
                    'bullets': []
                }
                grouped[current_game].append(current_article)
            continue

        m_bullet = re.match(r'^\s*[•\-\*]\s+(.+)', line)
        if m_bullet and current_article is not None:
            current_article['bullets'].append(m_bullet.group(1).strip())

    return grouped


def build_watched_html(grouped):
    """Buduje HTML sekcji obserwowanych gier pogrupowanych po grach."""
    if not grouped:
        return (
            '<table width="100%" style="border-collapse:collapse;margin-bottom:8px">'
            '<tr><td style="background:#f3e5f5;padding:8px 16px;font-weight:bold;font-size:14px">'
            '🎯 Obserwowane gry — nowości, patche, DLC, eventy</td></tr>'
            '<tr><td style="padding:8px 16px;color:#888">Brak nowych informacji o obserwowanych grach.</td></tr>'
            '</table>'
        )

    rows = []
    for game_name, articles in grouped.items():
        if not articles:
            continue
        rows.append(
            f'<p style="margin:10px 0 2px 0;font-weight:bold;font-size:14px;color:#6a1b9a">'
            f'🎮 {game_name}</p>'
        )
        for art in articles:
            item = art['item']
            link = item.get('link', '')
            title = item['title']
            title_html = (
                f'<a href="{link}" style="color:#1a73e8;text-decoration:none">{title}</a>'
                if link else title
            )
            article_html = (
                f'<p style="margin:2px 0 2px 16px;padding:4px 0;border-bottom:1px solid #f0f0f0">'
                f'◆ {title_html}<br>'
                f'<span style="color:#555;font-size:13px">{art["summary"]}</span>'
            )
            for b in art['bullets']:
                article_html += (
                    f'<br><span style="color:#555;font-size:13px;margin-left:12px">• {b}</span>'
                )
            article_html += '</p>'
            rows.append(article_html)

    inner = '\n'.join(rows)
    return (
        '<table width="100%" style="border-collapse:collapse;margin-bottom:8px">'
        '<tr><td style="background:#f3e5f5;padding:8px 16px;font-weight:bold;font-size:14px">'
        '🎯 Obserwowane gry — nowości, patche, DLC, eventy</td></tr>'
        f'<tr><td style="padding:4px 16px">{inner}</td></tr>'
        '</table>'
    )


def build_free_games_html(epic_items, gog_items):
    """Buduje HTML sekcji darmowych gier."""
    rows = []
    if epic_items:
        rows.append('<b style="color:#2e7d32">Epic Games Store:</b>')
        for g in epic_items:
            end_txt = f' <span style="color:#888;font-size:12px">(do {g["end"]})</span>' if g.get('end') else ''
            rows.append(
                f'<p style="margin:3px 0">🎮 <a href="{g["link"]}" '
                f'style="color:#1a73e8;text-decoration:none;font-weight:bold">{g["title"]}</a>{end_txt}</p>'
            )
    else:
        rows.append('<p style="color:#888">Epic: brak darmowych gier w tym tygodniu</p>')

    if gog_items:
        rows.append('<b style="color:#2e7d32;display:block;margin-top:8px">GOG — stale bezpłatne:</b>')
        for g in gog_items:
            rows.append(
                f'<p style="margin:3px 0">🎮 <a href="{g["link"]}" '
                f'style="color:#1a73e8;text-decoration:none;font-weight:bold">{g["title"]}</a></p>'
            )

    inner = '\n'.join(rows)
    return (
        '<table width="100%" style="border-collapse:collapse;margin-bottom:8px;'
        'border:2px solid #4caf50">'
        '<tr><td style="background:#f1f8e9;padding:8px 16px;font-weight:bold;font-size:14px">'
        '🎁 Darmowe gry — teraz!</td></tr>'
        f'<tr><td style="padding:8px 16px">{inner}</td></tr>'
        '</table>'
    )


# ── Claude ────────────────────────────────────────────────────────────────────

def call_claude(prompt, max_tokens=8192, timeout=300):
    data = json.dumps({
        'model': 'claude-sonnet-4-6',
        'max_tokens': max_tokens,
        'messages': [{'role': 'user', 'content': prompt}],
    }).encode('utf-8')
    req = urllib.request.Request(
        'https://api.anthropic.com/v1/messages',
        data=data,
        headers={
            'x-api-key': ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())['content'][0]['text']


def select_and_summarize_news(world_items, poland_items, local_items):
    """Claude przegląda WSZYSTKIE artykuły, wybiera najważniejsze i pisze streszczenia.
    Zwraca listy (item, summary) dla wybranych artykułów."""
    sections = []
    if world_items:
        sections.append('=== SWIAT ===')
        for i, it in enumerate(world_items):
            sections.append(f'{i+1}. {it["title"]}')
    if poland_items:
        sections.append('=== POLSKA ===')
        for i, it in enumerate(poland_items):
            sections.append(f'{i+1}. {it["title"]}')
    if local_items:
        sections.append('=== LOKALNE ===')
        for i, it in enumerate(local_items):
            sections.append(f'{i+1}. {it["title"]}')

    prompt = (
        'Jesteś redaktorem prasówki dla polskiego czytelnika mieszkającego w Kętach (małe miasto, woj. małopolskie).\n'
        'Przejrzyj poniższe artykuły i wybierz NAJWAŻNIEJSZE — te które naprawdę warto przeczytać.\n\n'
        'ZASADY SELEKCJI:\n'
        '- SWIAT i POLSKA: wybierz DOKŁADNIE 10 artykułów (lub wszystkie jeśli jest ich mniej niż 10).\n'
        '  Przy wojnie/epidemii/katastrofie/ataku na kraj — wybierz do 15.\n'
        '  Priorytet: polityka, gospodarka, bezpieczeństwo, tech/AI, ważne wydarzenia.\n'
        '  Pomiń tylko: oczywisty clickbait i powielające się tematy (zostaw najlepszy z grupy).\n'
        '  NIE pomijaj artykułów tylko dlatego że temat wydaje się mniej ważny — cel to 10.\n'
        '- LOKALNE: wybierz DOKŁADNIE 10 (lub wszystkie dostępne).\n'
        '  Pomiń tylko: sport, treści rodzinne/dla dzieci.\n\n'
        'Format odpowiedzi — tylko wybrane artykuły z ORYGINALNYM numerem i streszczeniem po polsku (max 15 słów):\n'
        '=== SWIAT ===\n'
        '3. Zdanie streszczające.\n'
        '7. Zdanie streszczające.\n'
        '=== POLSKA ===\n'
        '2. Zdanie streszczające.\n'
        '=== LOKALNE ===\n'
        '1. Zdanie streszczające.\n\n'
        'Artykuły do przejrzenia:\n'
        + '\n'.join(sections)
    )

    result = call_claude(prompt, max_tokens=3000, timeout=90)

    world_sel, poland_sel, local_sel = [], [], []
    current = None
    for line in result.split('\n'):
        line = line.strip()
        if '=== SWIAT' in line:    current = 'W'
        elif '=== POLSKA' in line: current = 'P'
        elif '=== LOKALNE' in line: current = 'L'
        else:
            m = re.match(r'^(\d+)\.\s+(.+)', line)
            if m:
                idx = int(m.group(1)) - 1
                summary = m.group(2).strip()
                if current == 'W' and 0 <= idx < len(world_items):
                    world_sel.append((world_items[idx], summary))
                elif current == 'P' and 0 <= idx < len(poland_items):
                    poland_sel.append((poland_items[idx], summary))
                elif current == 'L' and 0 <= idx < len(local_items):
                    local_sel.append((local_items[idx], summary))

    return world_sel, poland_sel, local_sel


def build_news_section_html(header, selected, bg='#e8f5e9', icon='🗞️'):
    """Buduje HTML dla jednej podsekcji newsów. selected = lista (item, summary)."""
    if not selected:
        return ''
    rows = []
    for it, summ in selected:
        link = it.get('link', '')
        title = it['title']
        title_html = (
            f'<a href="{link}" style="color:#1a73e8;text-decoration:none;font-weight:bold">{title}</a>'
            if link else f'<b>{title}</b>'
        )
        rows.append(
            f'<p style="margin:4px 0;padding:6px 0;border-bottom:1px solid #f0f0f0">'
            f'◆ {title_html}<br>'
            f'<span style="color:#555;font-size:13px">{summ}</span></p>'
        )
    inner = '\n'.join(rows)
    return (
        f'<table width="100%" style="border-collapse:collapse;margin-bottom:8px">'
        f'<tr><td style="background:{bg};padding:8px 16px;font-weight:bold;font-size:14px">'
        f'{icon} {header}</td></tr>'
        f'<tr><td style="padding:4px 16px">{inner}</td></tr>'
        f'</table>'
    )


# ── Email ─────────────────────────────────────────────────────────────────────

def send_email_graph(token, subject, html_body):
    message = {
        'message': {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': 'wrathtide@outlook.com'}}],
        }
    }
    data = json.dumps(message).encode('utf-8')
    req = urllib.request.Request(
        f'{GRAPH_BASE}/me/sendMail',
        data=data,
        headers={
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json',
        },
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        pass  # 202 Accepted


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    today = datetime.now(WARSAW_TZ).strftime('%A, %d.%m.%Y')

    print('Odswiezam token Microsoft...')
    access_token = refresh_access_token()

    print('Pobieram email, zadania i kalendarz...')
    emails = fetch_emails(access_token)
    todo = fetch_todo_tasks(access_token)
    calendar_events = fetch_calendar_events(access_token)
    special_days = get_polish_special_days(datetime.now(WARSAW_TZ).date())
    special_days_str = fmt_special_days(special_days)

    now_local = datetime.now(WARSAW_TZ)
    is_weekend = now_local.weekday() >= 5
    is_holiday = any(typ == 'DZIEN_WOLNY' for typ, name in special_days)
    is_free_day = is_weekend or is_holiday

    print('Pobieram pogode...')
    imgw_alerts = fetch_imgw_warnings()

    if not is_free_day:
        bb_wttr      = fetch_wttr('Bielsko-Biala,Poland', 6, 18)
        bb_openmeteo = fetch_openmeteo(BB_LAT, BB_LON, 7, 18)
        bb_yr        = fetch_yr(BB_LAT, BB_LON, 7, 18)
        bb_imgw      = fetch_imgw_synop('BIELSKO')
        bb_warnings  = compare_sources([bb_wttr, bb_openmeteo, bb_yr])
        sunrise = bb_wttr.get('sunrise', '?') if bb_wttr else '?'
        sunset  = bb_wttr.get('sunset',  '?') if bb_wttr else '?'
    else:
        bb_wttr = bb_openmeteo = bb_yr = bb_imgw = None
        bb_warnings = ''
        sunrise = sunset = '?'

    kety_h_start = 6  if is_free_day else 16
    kety_h_end   = 23
    kety_wttr      = fetch_wttr('Kety,Poland', kety_h_start, kety_h_end)
    kety_openmeteo = fetch_openmeteo(KETY_LAT, KETY_LON, kety_h_start, kety_h_end)
    kety_yr        = fetch_yr(KETY_LAT, KETY_LON, kety_h_start, kety_h_end)
    kety_warnings  = compare_sources([kety_wttr, kety_openmeteo, kety_yr])
    kety_label     = 'cały dzień (06:00-23:00)' if is_free_day else 'dom (16:00-23:00)'

    print('Pobieram ruch drogowy...')
    if not is_free_day:
        traffic_kety_bb = fetch_traffic_osrm(KETY_LAT, KETY_LON, BB_LAT, BB_LON)
        traffic_bb_kety = fetch_traffic_osrm(BB_LAT, BB_LON, KETY_LAT, KETY_LON)
    else:
        traffic_kety_bb = None
        traffic_bb_kety = None

    print('Pobieram newsy...')
    world_tvn24      = fetch_rss_items('https://tvn24.pl/swiat.xml', max_items=15)
    world_gnews      = fetch_news_rss('world economy technology AI cybersecurity politics', lang='en', country='US', max_items=15)
    poland_oko       = fetch_rss_items('https://oko.press/feed/', max_items=15)
    poland_tvn24     = fetch_rss_items('https://tvn24.pl/najwazniejsze.xml', max_items=15)
    rcb_alerts       = fetch_rss_items('https://www.rcb.gov.pl/feed/', max_items=5)
    news_alerts_kety = fetch_news_rss('Kęty Oświęcim straż pożarna policja sanepid ostrzeżenie awaria skażenie')
    kety_kety_pl     = fetch_rss_items('https://kety.pl/rss/aktualnosci.xml', max_items=15)
    kety_mamnewsa    = fetch_news_rss('site:mamnewsa.pl', max_items=15)
    kety_24kety      = fetch_rss_items('https://24kety.pl/feed/', max_items=15)
    gaming_gryonline   = fetch_rss_items('https://www.gry-online.pl/rss/news.xml', max_items=20)
    gaming_gram        = fetch_rss_items('https://www.gram.pl/rss/content.xml', max_items=20)
    gaming_ign         = fetch_rss_items('https://pl.ign.com/feed.xml', max_items=20)
    epic_free          = fetch_epic_free()
    gog_free_list      = fetch_gog_free()
    watched_news_raw   = fetch_watched_games_news()

    # Budowanie sekcji newsów — mały Claude call na streszczenia, HTML w Pythonie
    print('Generuje streszczenia newsow...')
    world_items  = world_tvn24 + world_gnews
    poland_items = poland_oko + poland_tvn24
    local_items  = kety_kety_pl + kety_mamnewsa + kety_24kety
    world_sel, poland_sel, local_sel = select_and_summarize_news(world_items, poland_items, local_items)
    news_html = (
        build_news_section_html('Świat — polityka, gospodarka, tech/AI', world_sel,
                                bg='#e3f2fd', icon='🌍') +
        build_news_section_html('Polska', poland_sel,
                                bg='#e8f5e9', icon='🇵🇱') +
        build_news_section_html('Kęty i okolice', local_sel,
                                bg='#fff8e1', icon='🏙️')
    )

    print('Generuje sekcje gamingowa...')
    gaming_all      = gaming_gryonline + gaming_gram + gaming_ign
    gaming_sel      = select_and_summarize_gaming(gaming_all)
    watched_sel     = select_watched_games_updates(watched_news_raw)
    free_games_html = build_free_games_html(epic_free, gog_free_list)
    gaming_news_html = build_news_section_html(
        'Newsy — Gry-Online, Gram.pl, IGN.pl', gaming_sel, bg='#fce4ec', icon='🎮'
    )
    watched_html    = build_watched_html(watched_sel)
    gaming_full_html = free_games_html + gaming_news_html + watched_html

    # Warunkowe bloki danych i instrukcji pogodowych
    if not is_free_day:
        bb_weather_data = (
            f"== POGODA BIELSKO-BIALA (07-18) ==\n"
            f"[wttr.in]     {fmt_source(bb_wttr)}\n"
            f"[open-meteo]  {fmt_source(bb_openmeteo)}\n"
            f"[yr.no]       {fmt_source(bb_yr)}\n"
            f"[IMGW stacja] {fmt_imgw(bb_imgw)}\n"
            f"IMGW_ALERTS: {imgw_alerts if imgw_alerts else 'brak'}\n"
            f"BB_WARNINGS: {bb_warnings if bb_warnings else 'brak rozbieznosci'}"
        )
        bb_weather_instruction = (
            f"   a) Naglowek 'Bielsko-Biala — praca (07:00-18:00)', tlo #e8f4fd\n"
            f"      Wschod slonca: {sunrise} | Zachod: {sunset}\n"
            f"      Narracyjne 2-3 zdania z wnioskow ze wszystkich zrodel\n"
            f"      Jesli IMGW_ALERTS niepuste: wiersz tlo #ffebee, ikona ⚠️ + tresc\n"
            f"      Jesli BB_WARNINGS niepuste: wiersz tlo #fff3cd, ikona ⚠️ + tresc\n\n"
            f"   b) Naglowek 'Kety — {kety_label}', tlo #e8fde8\n"
            f"      Narracyjne 2-3 zdania\n"
            f"      Jesli KETY_WARNINGS niepuste: wiersz tlo #fff3cd, ikona ⚠️ + tresc"
        )
        umbrella_rule = "bazuj na danych z obu miast"
    else:
        bb_weather_data = (
            f"== POGODA BIELSKO-BIALA == POMINIĘTO (IS_FREE_DAY=TAK)\n"
            f"IMGW_ALERTS: {imgw_alerts if imgw_alerts else 'brak'}"
        )
        bb_weather_instruction = (
            f"   Tylko jedna sekcja: Naglowek 'Kety — {kety_label}', tlo #e8fde8\n"
            f"   Narracyjne 2-3 zdania. Jesli KETY_WARNINGS niepuste: wiersz tlo #fff3cd, ikona ⚠️\n"
            f"   NIE dodawaj sekcji Bielsko-Biala!"
        )
        umbrella_rule = "bazuj wylacznie na danych z Ket"

    prompt = f"""Jestes asystentem Michala tworzacym jego poranny raport emailowy. Dzis: {today}.

KRYTYCZNE ZASADY TECHNICZNE — email bedzie wyswietlany w Outlook.com:
- Uzyj WYLACZNIE tabel HTML do layoutu (NIE div+flexbox, NIE CSS grid)
- Wszystkie style jako INLINE (style="...") — zadnych <style> w <head>
- Zadnego JavaScript
- Zadnych zewnetrznych fontow, animacji, pseudoelementow CSS
- Szerokosci jako liczby bez jednostki w atrybucie width (np. width="600")
- Kolory jako hex (#ffffff), nie rgba()
- Pisz po polsku
- Linki: <a href="URL" style="color:#1a73e8;text-decoration:none">Tytuł</a>

STRUKTURA (w tej kolejnosci):

1. NAGLOWEK: tabela width="600", tlo #1a73e8, bialy tekst, emoji slonce, "Poranny Raport", data i dzien tygodnia

2. SEKCJA POGODA — IS_FREE_DAY: {'TAK' if is_free_day else 'NIE'}:
   NIE rob tabelki godzin. Krotkie podsumowanie narracyjne (max 2-3 zdania na miasto).
   Skupiaj sie WYLACZNIE na: ryzyku opadow (kiedy, ile), zakresie temperatury, silnym wietrze (>30km/h).
   Jesli min_temp < 0: wiersz ostrzezenia tlo #ffe0e0 "Uwaga: mróz!"

{bb_weather_instruction}

   PODSUMOWANIE (zawsze, osobny wiersz tlo #eeeeee font-weight bold):
   ☂️ Parasol: [tak/nie — krotkie uzasadnienie]  |  🧥 Kurtka: [tak/nie — uzasadnienie]
   Reguly parasola: max_precip_prob > 30% LUB total_precip_mm > 0.5mm -> TAK
   Reguly kurtki: < 10°C -> ciezka kurtka; 10-17°C -> kurtka; 18-23°C -> lekka bluza; > 24°C -> nie potrzeba
   ({umbrella_rule})

3. SEKCJA KALENDARZ (naglowek tlo #e8f0fe, ikona 📅 "Kalendarz i ważne dni"):
   a) Jesli SPECIAL_DAYS zawiera DZIEN_WOLNY: prominentny wiersz tlo #ffebee, pogrubiony, ikona 🎉
      Jesli JUTRO_WOLNE: wiersz tlo #fff3cd, ikona ⏰ "Jutro dzień wolny: [nazwa]"
      Jesli DZIEN_SPECJALNY: wiersz tlo #e8f5e9, ikona 🎂/💐/👨‍👩‍👧 zaleznie od dnia
   b) Lista wydarzen z kalendarza:
      - [godzina] Tytuł | Miejsce; cały dzień: ikona 📌; WAZNE: ikona 🔴
   c) Jesli brak wydarzen: "Wolny dzień — brak spotkań"
   Cala sekcja: jesli DZIEN_WOLNY to border-left: 4px solid #fbc02d

4. DOJAZD DO PRACY (naglowek tlo #e8eaf6, ikona 🚗):
   IS_FREE_DAY: {'TAK' if is_free_day else 'NIE'}
   Jesli IS_FREE_DAY=TAK: "🏖️ Dzień wolny — brak dojazdu do pracy"
   Jesli IS_FREE_DAY=NIE:
     Czas bazowy bez korkow (OSRM, estymacja bez real-time traffic).
     - 🚗 Rano (wyjazd Kęty 7:45): ok. X min | Y km → przyjazd ok. [oblicz]
     - 🏠 Powrót (wyjazd BB 16:25): ok. X min | Y km → przyjazd ok. [oblicz]
     Jesli brak danych: "Brak danych o trasie"

5. ALERTY KRYTYCZNE — KETY I OKOLICE (naglowek tlo #ffebee, ikona 🚨 "Alerty i Ostrzeżenia"):
   Pokazuj sekcje ZAWSZE. Jesli brak alertow: zielony wiersz "✅ Brak aktywnych alertów".
   Zrodla: IMGW, RCB, lokalne sluby Kety (dane z sekcji ALERTY w danych).
   Kategorie alertow do pokazania (TYLKO te):
   🌩️ Pogoda — silny wiatr, burze, oblodzenie, upaly, intensywne opady (IMGW)
   🔴 RCB — oficjalne alerty rzadowe (Alert RCB SMS)
   💧 Skazenie wody pitnej
   🍽️ Zagrozenia zywnosciowe / sanepid
   🏥 Alerty zdrowotne, epidemiczne
   🚒 PSP Kety — pozary, zdarzenia niebezpieczne w okolicy
   👮 Policja Kety — zagrozenia bezpieczenstwa, poszukiwania, akcje
   Format: kazdy alert w osobnym wierszu tlo #ffe0e0, ikona tematyczna + opis + zrodlo jako klikalne
   <a href="URL">krotka_nazwa</a> (np. "IMGW", "RCB", "mamnewsa.pl") — NIE pokazuj pelnego URL w tresci.
   POMIŃ: ogolna polityka, gospodarka, sport, kultura bez bezposredniego zagrozenia.

6. SKRZYNKA ODBIORCZA (naglowek tlo #fff8e1, ikona 📬):
   Wyswietl TYLKO emaile spelniajace co najmniej jeden z kryteriow:
   a) Wymagaja reakcji uzytkownika w ciagu 3 dni (odpowiedz, potwierdzenie, platnosc, decyzja, termin)
   b) Alert bezpieczenstwa (weryfikacja logowania, zmiana hasla, podejrzana aktywnosc, phishing, 2FA)
   Dla kazdego: temat pogrubiony ([NOWE] jesli nieprzeczytany), nadawca+data, 1 zdanie dlaczego akcja.
   Obramowanie 2px solid #e53935 jesli deadline <= 2 dni lub alert bezpieczenstwa.
   Obramowanie 1px solid #fbc02d jesli deadline 2-3 dni.
   Jesli zadna nie kwalifikuje sie: "Brak pilnych wiadomosci — skrzynka spokojna ✅"

7. ZADANIA TO DO (naglowek tlo #f3e5f5):
   Lista, terminy pogrubione czerwono

8. WIADOMOSCI (naglowek tlo #e8f5e9, ikona 🗞️ "Wiadomości"):
   Sekcja newsow jest generowana automatycznie poza tym promptem.
   Wstaw w tym miejscu DOKLADNIE ten znacznik HTML i nic wiecej:
   <!-- NEWS_PLACEHOLDER -->
   (nie generuj zadnych artykulow — system wstawi je automatycznie)

9. GAMING I DARMOWE GRY:
   Sekcja gamingowa jest generowana automatycznie poza tym promptem.
   Wstaw w tym miejscu DOKLADNIE ten znacznik HTML i nic wiecej:
   <!-- GAMING_PLACEHOLDER -->
   (nie generuj zadnych tresci gamingowych — system wstawi je automatycznie)

---
DANE:

{bb_weather_data}

== POGODA KETY ({kety_label}) ==
[wttr.in]     {fmt_source(kety_wttr)}
[open-meteo]  {fmt_source(kety_openmeteo)}
[yr.no]       {fmt_source(kety_yr)}
KETY_WARNINGS: {kety_warnings if kety_warnings else 'brak rozbieznosci'}

== KALENDARZ I WAZNE DNI ==
SPECIAL_DAYS:
{special_days_str}

CALENDAR_EVENTS:
{calendar_events}

== DOJAZD (OSRM, bez real-time traffic) ==
IS_FREE_DAY: {'TAK' if is_free_day else 'NIE'}
Wyjazd rano: Kety 7:45 -> Bielsko-Biala
Wyjazd popol.: Bielsko-Biala 16:25 -> Kety
Kety -> Bielsko-Biala: {f"{traffic_kety_bb['duration_min']} min | {traffic_kety_bb['distance_km']} km" if traffic_kety_bb else ('WOLNY DZIEN' if is_free_day else 'BLAD')}
Bielsko-Biala -> Kety: {f"{traffic_bb_kety['duration_min']} min | {traffic_bb_kety['distance_km']} km" if traffic_bb_kety else ('WOLNY DZIEN' if is_free_day else 'BLAD')}

== SKRZYNKA (ostatnie 3 dni, tylko inbox) ==
{emails}

== ZADANIA TO DO ==
{todo}

== ALERTY KRYTYCZNE — IMGW ==
{imgw_alerts if imgw_alerts else 'brak'}

== ALERTY KRYTYCZNE — RCB ==
{fmt_rss_items(rcb_alerts)}

== ALERTY LOKALNE — KETY/OSWIECIM (Google News) ==
{fmt_rss_items(news_alerts_kety)}

---
Zacznij od <!DOCTYPE html><html><body> i skoncz </body></html>.
Caly email: jedna zewnetrzna tabela width="600" align="center" style="background:#ffffff;border:1px solid #e0e0e0;font-family:Arial,sans-serif;font-size:14px".
Kazda sekcja: osobna tabela wewnatrz, width="100%", padding 16px.
PRIORYTET: kompletny zamkniety HTML > liczba artykulow. Jesli zbliżasz sie do limitu miejsca,
skróc streszczenia lub pomiń ostatnie artykuły w sekcji, ale ZAWSZE zakoncz </body></html>.
"""

    print('Generuje raport przez Claude Sonnet...')
    report = call_claude(prompt)

    if '<!DOCTYPE' in report:
        report = report[report.find('<!DOCTYPE'):]
    elif '<html' in report:
        report = report[report.find('<html'):]

    # Wstaw gotowy HTML newsów i gamingu w miejsce placeholderów
    report = report.replace('<!-- NEWS_PLACEHOLDER -->', news_html)
    report = report.replace('<!-- GAMING_PLACEHOLDER -->', gaming_full_html)

    print('Wysylam email...')
    send_email_graph(access_token, f'Poranny raport - {today}', report)
    print('Gotowe!')


if __name__ == '__main__':
    main()
