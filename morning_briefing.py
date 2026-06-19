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
        return 'Brak wiadomości z ostatnich 3 dni.'
    lines = []
    for m in messages:
        status = '[NOWE] ' if not m.get('isRead') else ''
        sender = m.get('from', {}).get('emailAddress', {}).get('address', '?')
        lines.append(
            f"{status}Od: {sender}\n"
            f"Temat: {m.get('subject', '(brak tematu)')}\n"
            f"Data: {m.get('receivedDateTime', '')[:10]}\n"
            f"Podgląd: {m.get('bodyPreview', '')[:200]}\n"
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
            task_lines.append(f"• [{lst['displayName']}] {task['title']}{due_str}")
    return '\n'.join(task_lines) if task_lines else 'Brak aktywnych zadań.'


def fetch_url(url, timeout=30):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read().decode('utf-8', errors='replace')
    except Exception as e:
        return f'[Błąd: {e}]'


def fetch_weather(city):
    return fetch_url(f'https://wttr.in/{urllib.parse.quote(city)}?format=j1')


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
        pass  # 202 Accepted, brak body


def main():
    today = datetime.now().strftime('%A, %d.%m.%Y')

    print('Odswiezam token Microsoft...')
    access_token = refresh_access_token()

    print('Pobieram dane...')
    emails = fetch_emails(access_token)
    todo = fetch_todo_tasks(access_token)
    weather_kety = fetch_weather('Kety,Poland')
    weather_bb = fetch_weather('Bielsko-Biala,Poland')
    news_world = fetch_news_rss('world news today', lang='en', country='US')
    news_poland = fetch_news_rss('Polska wiadomosci dzis')
    news_local = fetch_news_rss('Kety Bielsko-Biala')
    news_gaming = fetch_news_rss('free games Epic GOG Steam Amazon Prime Gaming', lang='en', country='US')
    epic_games = fetch_article('https://store.epicgames.com/en-US/free-games')
    gog_free = fetch_article('https://www.gog.com/en/games?features=free')

    prompt = f"""Jestes asystentem Michala tworzacym jego poranny raport. Dzis: {today}.

Na podstawie ponizszych danych przygotuj strone HTML — czytelna, estetyczna, z sekcjami.
Pisz po polsku. Uzywaj emoji jako ikon sekcji. Nie powtarzaj danych — wyciagaj z nich esencje.

ZASADY FORMATOWANIA:
- Skrzynka odbiorcza: wypisz NOWE wiadomosci, zaznacz jesli cos wyglada na deadline/pilne
- Zadania To Do: lista zadan z terminami, posortowana priorytetowo
- Pogoda: temperatura min/max, odczuwalna, opady, wiatr. Jedno zdanie podsumowania.
- Wiadomosci swiatowe: 3-4 najwazniejsze tematy, po 2 zdania kazdy
- Wiadomosci Polska: 3-4 tematy, po 2 zdania
- Wiadomosci lokalne Kety/Bielsko: wszystko co jest, krotko
- Darmowe gry: PELNA INFO — tytul, opis, platforma, do kiedy bezplatna
- Gaming news: tylko tytuly i jedno zdanie opisu

DANE ZRODLOWE:

\U0001f4e7 SKRZYNKA ODBIORCZA (ostatnie 3 dni):
{emails}

✅ ZADANIA TO DO:
{todo}

\U0001f324 POGODA KETY:
{weather_kety[:1800]}

\U0001f324 POGODA BIELSKO-BIALA:
{weather_bb[:1800]}

\U0001f30d WIADOMOSCI SWIATOWE:
{news_world}

\U0001f1f5\U0001f1f1 WIADOMOSCI POLSKA:
{news_poland}

\U0001f4cd WIADOMOSCI LOKALNE:
{news_local[:2000]}

\U0001f3ae GAMING / DARMOWE GRY (RSS):
{news_gaming}

\U0001f381 EPIC GAMES DARMOWE GRY:
{epic_games}

\U0001f381 GOG DARMOWE GRY:
{gog_free}

Napisz kompletna strone HTML zaczynajac od <!DOCTYPE html> i konczac na </html>.
Dodaj na gorze naglowek "Poranny raport - {today}" oraz aktualna date.
Uzywaj emoji w naglowkach sekcji.
Uzywaj czcionki bezszeryfowej (font-family: system-ui, sans-serif), jasnego tla (#f9f9f9),
kart z bialym tlem i cieniem dla sekcji, czytelnych odstepow.
Strona ma byc wygodna do czytania na telefonie (max-width: 720px, margin: auto).
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
