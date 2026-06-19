import os
import json
import urllib.request
import urllib.parse
from datetime import datetime, timedelta, timezone

ANTHROPIC_API_KEY = os.environ['ANTHROPIC_API_KEY']
MS_CLIENT_ID = os.environ['MS_CLIENT_ID']
MS_REFRESH_TOKEN = os.environ['MS_REFRESH_TOKEN']
GRAPH_BASE = 'https://graph.microsoft.com/v1.0'


def refresh_access_token():
    data = urllib.parse.urlencode({
        'grant_type': 'refresh_token',
        'client_id': MS_CLIENT_ID,
        'refresh_token': MS_REFRESH_TOKEN,
        'scope': 'Mail.Read Mail.Send Tasks.Read User.Read offline_access',
    }).encode()
    req = urllib.request.Request(
        'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
        data=data,
        headers={'Content-Type': 'application/x-www-form-urlencoded'},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())['access_token']


def graph_get(token, path, params=None):
    url = f'{GRAPH_BASE}{path}'
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
    result = graph_get(token, '/me/messages', {
        '$top': 20,
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
            f"Podglad: {m.get('bodyPreview', '')[:200]}\n"
        )
    return '\n'.join(lines)


def fetch_todo_tasks(token):
    lists_result = graph_get(token, '/me/todo/lists')
    task_lines = []
    for lst in lists_result.get('value', []):
        tasks_result = graph_get(token, f"/me/todo/lists/{lst['id']}/tasks", {
            '$filter': "status ne 'completed'",
            '$top': 30,
        })
        for task in tasks_result.get('value', []):
            due = task.get('dueDateTime')
            due_str = f" [termin: {due['dateTime'][:10]}]" if due else ''
            task_lines.append(f"- [{lst['displayName']}] {task['title']}{due_str}")
    return '\n'.join(task_lines) if task_lines else 'Brak aktywnych zadan.'


def fetch_url(url, timeout=30):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read().decode('utf-8', errors='replace')
    except Exception as e:
        return f'[Blad: {e}]'


def fetch_weather_json(city):
    raw = fetch_url(f'https://wttr.in/{urllib.parse.quote(city)}?format=j1')
    try:
        return json.loads(raw)
    except Exception:
        return None


def parse_weather(data, hour_times):
    """Wyciaga dane pogodowe dla wybranych godzin (np. [600, 900, 1200, 1500, 1800])."""
    if not data:
        return '[Brak danych pogodowych]'
    try:
        today = data['weather'][0]
        astro = today.get('astronomy', [{}])[0]
        lines = [
            f"Min: {today['mintempC']}C / Max: {today['maxtempC']}C",
            f"Wschod slonca: {astro.get('sunrise','?')} | Zachod: {astro.get('sunset','?')}",
            '',
        ]
        for h in today['hourly']:
            t = int(h['time'])
            if t in hour_times:
                hour_h = t // 100
                desc = h['weatherDesc'][0]['value'] if h.get('weatherDesc') else ''
                chance_rain = h.get('chanceofrain', '0')
                lines.append(
                    f"{hour_h:02d}:00 | {h['tempC']}C (odczuwalnie {h['FeelsLikeC']}C) | "
                    f"{desc} | opady: {h['precipMM']}mm ({chance_rain}% szans) | "
                    f"wiatr: {h['windspeedKmph']} km/h {h.get('winddir16Point','')}"
                )
        return '\n'.join(lines)
    except Exception as e:
        return f'[Blad parsowania pogody: {e}]'


def fetch_news_rss(query, lang='pl', country='PL'):
    encoded = urllib.parse.quote(query)
    url = f'https://news.google.com/rss/search?q={encoded}&hl={lang}&gl={country}&ceid={country}:{lang}'
    return fetch_url(url)[:4000]


def fetch_article(article_url):
    return fetch_url(f'https://r.jina.ai/{article_url}', timeout=45)[:3000]


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


def main():
    today = datetime.now().strftime('%A, %d.%m.%Y')

    print('Odswiezam token Microsoft...')
    access_token = refresh_access_token()

    print('Pobieram dane...')
    emails = fetch_emails(access_token)
    todo = fetch_todo_tasks(access_token)

    # Pogoda — godzinowa
    # Bielsko-Biala: godziny pracy/dnia 07-18 -> wttr.in hourly: 600 900 1200 1500 1800
    # Kety: popoludnie/wieczor 16-23 -> wttr.in hourly: 1500 1800 2100
    weather_bb_json = fetch_weather_json('Bielsko-Biala,Poland')
    weather_kety_json = fetch_weather_json('Kety,Poland')
    weather_bb = parse_weather(weather_bb_json, [600, 900, 1200, 1500, 1800])
    weather_kety = parse_weather(weather_kety_json, [1500, 1800, 2100])

    news_world = fetch_news_rss('world news today', lang='en', country='US')
    news_poland = fetch_news_rss('Polska wiadomosci dzis')
    news_local = fetch_news_rss('Kety Bielsko-Biala')
    news_gaming = fetch_news_rss('free games Epic GOG Steam Amazon Prime Gaming', lang='en', country='US')
    epic_games = fetch_article('https://store.epicgames.com/en-US/free-games')
    gog_free = fetch_article('https://www.gog.com/en/games?features=free')

    prompt = f"""Jestes asystentem Michala tworzacym jego poranny raport. Dzis: {today}.

Przygotuj kompletna strone HTML. Pisz po polsku. Sekcje w tej KOLEJNOSCI:

1. POGODA (pierwsza i najwazniejsza sekcja)
2. SKRZYNKA ODBIORCZA
3. ZADANIA TO DO
4. WIADOMOSCI (swiatowe, polskie, lokalne w jednej sekcji z podsekcjami)
5. GAMING I DARMOWE GRY

---
ZASADY SEKCJI POGODA:
Pokaz dwie karty obok siebie (lub jedna pod druga na mobile):
- Karta "Bielsko-Biala - praca (07:00-18:00)": tabela godzin z temperatura, opisem, opadami, wiatrem
- Karta "Kety - dom (16:00-23:00)": tabela godzin z temperatura, opisem, opadami, wiatrem
Na gorze kazdej karty: min/max dnia, wschod/zachod slonca.
Jesli sa opady > 0 lub duzy wiatr (>30km/h) — zaznacz czerwonym/pomaranczowym kolorem.
Jedno zdanie komentarza: czy brac parasol/kurtke.

ZASADY POZOSTALYCH SEKCJI:
- Skrzynka: NOWE wiadomosci pogrubione, zaznacz pilne/deadline czerwona ramka
- Zadania To Do: pogrupowane po liscie, terminy pogrubione
- Wiadomosci swiatowe: 3-4 tematy, 2 zdania kazdy
- Wiadomosci Polska: 3-4 tematy, 2 zdania kazdy
- Wiadomosci lokalne: wszystko co jest, krotko
- Darmowe gry: PELNA INFO — tytul, opis, platforma, do kiedy bezplatna; wyroznic wizualnie
- Gaming news: tytul + jedno zdanie

---
DANE POGODOWE:

BIELSKO-BIALA (07:00-18:00):
{weather_bb}

KETY (16:00-23:00):
{weather_kety}

---
SKRZYNKA (ostatnie 3 dni):
{emails}

ZADANIA TO DO:
{todo}

WIADOMOSCI SWIATOWE:
{news_world}

WIADOMOSCI POLSKA:
{news_poland}

WIADOMOSCI LOKALNE:
{news_local[:2000]}

GAMING / DARMOWE GRY:
{news_gaming}

EPIC GAMES DARMOWE:
{epic_games}

GOG DARMOWE:
{gog_free}

---
WYMAGANIA HTML:
- Zacznij od <!DOCTYPE html>, skoncz na </html>
- Naglowek strony: "Poranny raport - {today}"
- font-family: system-ui, -apple-system, sans-serif
- Tlo strony: #f0f2f5
- max-width: 680px, margin: 0 auto, padding: 16px
- Karty sekcji: background white, border-radius: 12px, box-shadow: 0 2px 8px rgba(0,0,0,0.08), padding: 20px, margin-bottom: 16px
- Naglowki sekcji: emoji + nazwa, font-size: 18px, font-weight: 700, margin-bottom: 12px, color: #1a1a2e
- Tabela pogody: width:100%, border-collapse:collapse, kazda komorka padding:6px 8px, naprzemienne tlo wierszy
- Responsive: na mobile (max-width:480px) karty pogody jedna pod druga
- Kolory alertow: opady>0.5mm lub wiatr>30km/h -> komorka background #fff3cd; opady>3mm -> #ffe0e0
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
