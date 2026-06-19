import os
import json
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
        list_id = lst['id']
        # Paginate through all tasks (Graph API max $top=100 per page)
        url = f"/me/todo/lists/{list_id}/tasks"
        params = {'$filter': "status ne 'completed'", '$top': 100}
        while url:
            result = graph_get(token, url, params if params else None)
            params = None  # only on first request
            for task in result.get('value', []):
                due = task.get('dueDateTime')
                due_str = f" [termin: {due['dateTime'][:10]}]" if due else ''
                body = (task.get('body') or {}).get('content', '') or ''
                body_preview = body[:100].strip() if body.strip() else ''
                note = f" | {body_preview}" if body_preview else ''
                task_lines.append(f"- [{list_name}] {task['title']}{due_str}{note}")
            next_link = result.get('@odata.nextLink', '')
            url = next_link if next_link else None
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


# ── Newsy ─────────────────────────────────────────────────────────────────────

def fetch_news_rss(query, lang='pl', country='PL'):
    encoded = urllib.parse.quote(query)
    url = f'https://news.google.com/rss/search?q={encoded}&hl={lang}&gl={country}&ceid={country}:{lang}'
    return fetch_url(url)[:4000]


def fetch_article(article_url):
    return fetch_url(f'https://r.jina.ai/{article_url}', timeout=45)[:3000]


# ── Claude ────────────────────────────────────────────────────────────────────

def call_claude(prompt):
    data = json.dumps({
        'model': 'claude-sonnet-4-6',
        'max_tokens': 4096,
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
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read())['content'][0]['text']


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

    print('Pobieram pogode (4 zrodla)...')
    bb_wttr      = fetch_wttr('Bielsko-Biala,Poland', 6, 18)
    bb_openmeteo = fetch_openmeteo(BB_LAT, BB_LON, 7, 18)
    bb_yr        = fetch_yr(BB_LAT, BB_LON, 7, 18)
    bb_imgw      = fetch_imgw_synop('BIELSKO')
    bb_warnings  = compare_sources([bb_wttr, bb_openmeteo, bb_yr])
    imgw_alerts  = fetch_imgw_warnings()

    kety_wttr      = fetch_wttr('Kety,Poland', 15, 23)
    kety_openmeteo = fetch_openmeteo(KETY_LAT, KETY_LON, 16, 23)
    kety_yr        = fetch_yr(KETY_LAT, KETY_LON, 16, 23)
    kety_warnings  = compare_sources([kety_wttr, kety_openmeteo, kety_yr])

    sunrise = bb_wttr.get('sunrise', '?') if bb_wttr else '?'
    sunset  = bb_wttr.get('sunset',  '?') if bb_wttr else '?'

    print('Pobieram newsy i gry...')
    news_world  = fetch_news_rss('world news today', lang='en', country='US')
    news_poland = fetch_news_rss('Polska wiadomosci dzis')
    news_local  = fetch_news_rss('Kety Bielsko-Biala')
    news_gaming = fetch_news_rss('free games Epic GOG Steam Amazon Prime Gaming', lang='en', country='US')
    epic_games  = fetch_article('https://store.epicgames.com/en-US/free-games')
    gog_free    = fetch_article('https://www.gog.com/en/games?features=free')

    prompt = f"""Jestes asystentem Michala tworzacym jego poranny raport emailowy. Dzis: {today}.

KRYTYCZNE ZASADY TECHNICZNE — email bedzie wyswietlany w Outlook.com:
- Uzyj WYLACZNIE tabel HTML do layoutu (NIE div+flexbox, NIE CSS grid)
- Wszystkie style jako INLINE (style="...") — zadnych <style> w <head>
- Zadnego JavaScript
- Zadnych zewnetrznych fontow, animacji, pseudoelementow CSS
- Szerokosci jako liczby bez jednostki w atrybucie width (np. width="600")
- Kolory jako hex (#ffffff), nie rgba()
- Pisz po polsku

STRUKTURA (w tej kolejnosci):

1. NAGLOWEK: tabela width="600", tlo #1a73e8, bialy tekst, emoji slonce, "Poranny Raport", data i dzien tygodnia

2. SEKCJA POGODA:
   NIE rob tabelki godzin. Krotkie podsumowanie narracyjne (max 2-3 zdania na miasto).
   Skupiaj sie WYLACZNIE na: ryzyku opadow (kiedy, ile), zakresie temperatury (czy zimno/cieplo/upal/mroz), silnym wietrze (>30km/h).
   Jesli min_temp < 0: wiersz ostrzezenia tlo #ffe0e0 "Uwaga: mróz!"

   a) Naglowek "Bielsko-Biala — praca (07:00-18:00)", tlo #e8f4fd
      Wschod slonca: {sunrise} | Zachod: {sunset}
      Narracyjne 2-3 zdania z wnioskow ze wszystkich zrodel
      Jesli IMGW_ALERTS niepuste: wiersz tlo #ffebee, ikona ⚠️ + tresc alertu
      Jesli BB_WARNINGS niepuste: wiersz tlo #fff3cd, ikona ⚠️ + tresc

   b) Naglowek "Kety — dom (16:00-23:00)", tlo #e8fde8
      Narracyjne 2-3 zdania
      Jesli KETY_WARNINGS niepuste: wiersz tlo #fff3cd, ikona ⚠️ + tresc

   c) PODSUMOWANIE (osobny wiersz, tlo #eeeeee, font-weight bold):
      Format DOKLADNIE taki (nie zmieniaj struktury):
      ☂️ Parasol: [tak/nie — krotkie uzasadnienie]  |  🧥 Kurtka: [tak/nie — uzasadnienie, co wieczorem w Ketach]

      Reguly parasola: jesli max_precip_prob > 30% LUB total_precip_mm > 0.5mm w ktorymkolwiek miescie -> TAK
      Reguly kurtki: min_temp < 10°C -> ciezka kurtka; 10-17°C -> kurtka; 18-23°C -> lekka bluza; > 24°C -> nie potrzeba

3. SEKCJA KALENDARZ (naglowek tlo #e8f0fe, ikona 📅 "Kalendarz i ważne dni"):
   a) Jesli SPECIAL_DAYS zawiera DZIEN_WOLNY: prominentny wiersz tlo #ffebee, pogrubiony, ikona 🎉
      Jesli JUTRO_WOLNE: wiersz tlo #fff3cd, ikona ⏰ "Jutro dzień wolny: [nazwa]"
      Jesli DZIEN_SPECJALNY: wiersz tlo #e8f5e9, ikona 🎂/💐/👨‍👩‍👧 zaleznie od dnia
   b) Lista wydarzen z kalendarza (jesli sa):
      - Kazde wydarzenie: [godzina] Tytuł | Miejsce (jesli jest)
      - Caly dzien: ikona 📌 zamiast godziny
      - WAZNE: ikona 🔴 przy tytule
      - Krotki podglad opisu jesli nie jest pusty
   c) Jesli CALENDAR_EVENTS == "Brak wydarzen": jeden wiersz "Wolny dzień — brak spotkań"
   Cala sekcja: jezeli jest DZIEN_WOLNY to dodaj subtelny zolty ramki do calej sekcji (border-left: 4px solid #fbc02d)

4. SKRZYNKA ODBIORCZA (naglowek tlo #fff8e1, ikona 📬):
   Przeanalizuj WSZYSTKIE emaile z sekcji SKRZYNKA i wyswietl TYLKO te, ktore spelniaja co najmniej jeden z kryteriow:
   a) Wymagaja REAKCJI uzytkownika w ciagu 3 dni (odpowiedz, potwierdzenie, platnosc, decyzja, termin, spotkanie do zaakceptowania)
   b) Alert bezpieczenstwa (weryfikacja logowania, zmiana hasla, podejrzana aktywnosc, phishing warning, 2FA, konto zablokowane)

   Dla kazdego zakwalifikowanego emaila:
   - Temat pogrubiony jako naglowek, [NOWE] jesli nieprzeczytany
   - Nadawca + data (krotko)
   - 1 zdanie: DLACZEGO wymaga akcji / jaki rodzaj alertu
   - Obramowanie 2px solid #e53935 jesli deadline <= 2 dni lub alert bezpieczenstwa
   - Obramowanie 1px solid #fbc02d jesli deadline 2-3 dni

   Jesli ZADNA wiadomosc nie kwalifikuje sie: jeden wiersz szary "Brak pilnych wiadomosci — skrzynka spokojna ✅"
   NIE pokazuj zwyklych newsletterow, reklam, powiadomien serwisowych, potwierdzen zamowien bez akcji.

4. ZADANIA TO DO (naglowek tlo #f3e5f5):
   Lista, terminy pogrubione czerwono

5. WIADOMOSCI (naglowek tlo #e8f5e9):
   Podsekcje: Swiat | Polska | Lokalne. Kazdy temat: tytul + 2 zdania.

6. GAMING I DARMOWE GRY (naglowek tlo #fce4ec):
   Darmowe gry: ramka 2px solid #4caf50, pelna informacja (co, gdzie, do kiedy)
   Pozostale gaming newsy: lista

---
DANE:

== POGODA BIELSKO-BIALA (07-18) ==
[wttr.in]     {fmt_source(bb_wttr)}
[open-meteo]  {fmt_source(bb_openmeteo)}
[yr.no]       {fmt_source(bb_yr)}
[IMGW stacja] {fmt_imgw(bb_imgw)}
IMGW_ALERTS: {imgw_alerts if imgw_alerts else 'brak'}
BB_WARNINGS: {bb_warnings if bb_warnings else 'brak rozbieznosci'}

== POGODA KETY (16-23) ==
[wttr.in]     {fmt_source(kety_wttr)}
[open-meteo]  {fmt_source(kety_openmeteo)}
[yr.no]       {fmt_source(kety_yr)}
KETY_WARNINGS: {kety_warnings if kety_warnings else 'brak rozbieznosci'}

== KALENDARZ I WAZNE DNI ==
SPECIAL_DAYS:
{special_days_str}

CALENDAR_EVENTS:
{calendar_events}

== SKRZYNKA (ostatnie 3 dni) ==
{emails}

== ZADANIA TO DO ==
{todo}

== WIADOMOSCI SWIATOWE ==
{news_world}

== WIADOMOSCI POLSKA ==
{news_poland}

== WIADOMOSCI LOKALNE ==
{news_local[:2000]}

== GAMING / DARMOWE GRY ==
{news_gaming}

EPIC GAMES DARMOWE:
{epic_games}

GOG DARMOWE:
{gog_free}

---
Zacznij od <!DOCTYPE html><html><body> i skoncz </body></html>.
Caly email: jedna zewnetrzna tabela width="600" align="center" style="background:#ffffff;border:1px solid #e0e0e0;font-family:Arial,sans-serif;font-size:14px".
Kazda sekcja: osobna tabela wewnatrz, width="100%", padding 16px.
"""

    print('Generuje raport przez Claude Sonnet...')
    report = call_claude(prompt)

    if '<!DOCTYPE' in report:
        report = report[report.find('<!DOCTYPE'):]
    elif '<html' in report:
        report = report[report.find('<html'):]

    print('Wysylam email...')
    send_email_graph(access_token, f'Poranny raport - {today}', report)
    print('Gotowe!')


if __name__ == '__main__':
    main()
