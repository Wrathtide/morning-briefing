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
    data = fetch_json('https://danepubliczne.imgw.pl/api/data/warnings')
    if not data:
        return ''
    try:
        relevant = ['śląskie', 'slaskie', 'małopolskie', 'malopolskie',
                    'slask', 'malopolska', 'silesia']
        items = data if isinstance(data, list) else []
        warnings = []
        for w in items:
            region = str(w.get('obszar', '') or w.get('region', '') or '').lower()
            if any(v in region for v in relevant):
                level = w.get('stopien') or w.get('level', '?')
                phenomenon = w.get('zjawisko') or w.get('phenomenon', '?')
                time_range = (w.get('czas_od_do')
                              or f"{w.get('od', '')} - {w.get('do', '')}")
                warnings.append(f"STOPIEN {level}: {phenomenon} ({time_range})")
        return '\n'.join(warnings)
    except Exception:
        return ''


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


def get_news_summaries(world_items, poland_items, local_items):
    """Małe, szybkie wołanie Claude — tylko plain-text streszczenia."""
    sections = []
    if world_items:
        sections.append(f'=== SWIAT ===')
        for i, it in enumerate(world_items):
            sections.append(f'{i+1}. {it["title"]}')
    if poland_items:
        sections.append(f'=== POLSKA ===')
        for i, it in enumerate(poland_items):
            sections.append(f'{i+1}. {it["title"]}')
    if local_items:
        sections.append(f'=== LOKALNE ===')
        for i, it in enumerate(local_items):
            sections.append(f'{i+1}. {it["title"]}')

    prompt = (
        'Napisz 1 zdanie po polsku (max 15 słów) dla każdego artykułu.\n'
        'Dla sekcji LOKALNE: jeśli artykuł dotyczy sportu, treści rodzinnych lub dla dzieci, '
        'napisz zamiast streszczenia dokładnie: SKIP\n'
        'Zachowaj format z sekcjami i numeracją.\n\n'
        + '\n'.join(sections)
    )
    result = call_claude(prompt, max_tokens=2000, timeout=60)

    world_sum, poland_sum, local_sum = [], [], []
    current = None
    for line in result.split('\n'):
        line = line.strip()
        if '=== SWIAT' in line:   current = 'W'
        elif '=== POLSKA' in line: current = 'P'
        elif '=== LOKALNE' in line: current = 'L'
        else:
            m = re.match(r'^\d+\.\s+(.+)', line)
            if m:
                s = m.group(1).strip()
                if current == 'W': world_sum.append(s)
                elif current == 'P': poland_sum.append(s)
                elif current == 'L': local_sum.append(s)

    for lst, items in [(world_sum, world_items), (poland_sum, poland_items), (local_sum, local_items)]:
        while len(lst) < len(items):
            lst.append(items[len(lst)]['title'][:100])

    return world_sum[:len(world_items)], poland_sum[:len(poland_items)], local_sum[:len(local_items)]


def build_news_section_html(header, items, summaries, bg='#e8f5e9', icon='🗞️', skip_marked=False):
    """Buduje HTML dla jednej podsekcji newsów w Pythonie."""
    rows = []
    for it, summ in zip(items, summaries):
        if skip_marked and summ.strip().upper() == 'SKIP':
            continue
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
    if not rows:
        return ''
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
    gaming_gryonline = fetch_rss_items('https://www.gry-online.pl/rss/news.xml', max_items=5)
    gaming_lowcygier = fetch_rss_items('https://lowcygier.pl/feed/', max_items=5)
    epic_games       = fetch_article('https://store.epicgames.com/en-US/free-games')
    gog_free         = fetch_article('https://www.gog.com/en/games?features=free')

    # Budowanie sekcji newsów — mały Claude call na streszczenia, HTML w Pythonie
    print('Generuje streszczenia newsow...')
    world_items  = (world_tvn24 + world_gnews)[:10]
    poland_items = (poland_oko + poland_tvn24)[:10]
    local_items  = (kety_kety_pl + kety_mamnewsa + kety_24kety)[:10]
    world_sum, poland_sum, local_sum = get_news_summaries(world_items, poland_items, local_items)
    news_html = (
        build_news_section_html('Świat — polityka, gospodarka, tech/AI', world_items, world_sum,
                                bg='#e3f2fd', icon='🌍') +
        build_news_section_html('Polska', poland_items, poland_sum,
                                bg='#e8f5e9', icon='🇵🇱') +
        build_news_section_html('Kęty i okolice', local_items, local_sum,
                                bg='#fff8e1', icon='🏙️', skip_marked=True)
    )

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

9. GAMING I DARMOWE GRY (naglowek tlo #fce4ec):
   a) DARMOWE GRY (ramka 2px solid #4caf50, tlo #f1f8e9, NA GORZE!):
      Z danych EPIC i GOG wyodrebnij konkretne gry. Kazda: nazwa jako <a href="URL">Gra</a>.
      Podaj termin jesli widoczny. Jesli brak konkretnych gier: "Brak darmowych gier w tej chwili"
   b) GRY-ONLINE.PL — NEWSY:
      Kazdy artykul: tytuł jako <a href="URL" style="color:#1a73e8;text-decoration:none">Tytuł</a>
      Pod spodem 2-3 punkty (•) streszczajace na podstawie tytulu i opisu.
   c) ŁOWCY GIER:
      Kazdy: <a href="URL" style="color:#1a73e8;text-decoration:none">Tytuł</a> — 1 zdanie opisu

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

== GAMING — GRY-ONLINE.PL ==
{fmt_rss_items(gaming_gryonline)}

== GAMING — ŁOWCY GIER ==
{fmt_rss_items(gaming_lowcygier)}

== DARMOWE GRY — EPIC GAMES ==
{epic_games}

== DARMOWE GRY — GOG ==
{gog_free}

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

    # Wstaw gotowy HTML newsów w miejsce placeholdera
    report = report.replace('<!-- NEWS_PLACEHOLDER -->', news_html)

    print('Wysylam email...')
    send_email_graph(access_token, f'Poranny raport - {today}', report)
    print('Gotowe!')


if __name__ == '__main__':
    main()
