import os
import json
import smtplib
import urllib.request
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

ANTHROPIC_API_KEY = os.environ['ANTHROPIC_API_KEY']
EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
EMAIL_ADDRESS = 'kwasniak.michal@outlook.com'


def fetch_url(url, timeout=30):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read().decode('utf-8', errors='replace')
    except Exception as e:
        return f'[Błąd pobierania: {e}]'


def fetch_weather(city):
    url = f'https://wttr.in/{urllib.parse.quote(city)}?format=j1'
    return fetch_url(url)


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
        'messages': [{'role': 'user', 'content': prompt}]
    }).encode('utf-8')

    req = urllib.request.Request(
        'https://api.anthropic.com/v1/messages',
        data=data,
        headers={
            'x-api-key': ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json'
        }
    )
    with urllib.request.urlopen(req, timeout=120) as r:
        result = json.loads(r.read().decode('utf-8'))
        return result['content'][0]['text']


def send_email(subject, body_html):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_ADDRESS
    msg.attach(MIMEText(body_html, 'html', 'utf-8'))

    with smtplib.SMTP('smtp-mail.outlook.com', 587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg.as_string())


def main():
    today = datetime.now().strftime('%A, %d.%m.%Y')
    print('Pobieram dane...')

    weather_kety = fetch_weather('Kety,Poland')
    weather_bb = fetch_weather('Bielsko-Biala,Poland')

    news_world = fetch_news_rss('world news today', lang='en', country='US')
    news_poland = fetch_news_rss('Polska wiadomości dziś')
    news_local = fetch_news_rss('Kęty Bielsko-Biała')
    news_gaming = fetch_news_rss('free games Epic GOG Steam Amazon Prime Gaming', lang='en', country='US')

    epic_games = fetch_article('https://store.epicgames.com/en-US/free-games')
    gog_free = fetch_article('https://www.gog.com/en/games?features=free')

    prompt = f"""Jesteś asystentem Michała tworzącym jego poranny raport. Dziś: {today}.

Na podstawie poniższych danych przygotuj JEDEN e-mail HTML — czytelny, estetyczny, z sekcjami.
Pisz po polsku. Używaj emoji jako ikon sekcji. Nie powtarzaj danych — wyciągaj z nich esencję.

ZASADY FORMATOWANIA:
- Pogoda: temperatura min/max, odczuwalna, opady, wiatr. Jedno zdanie podsumowania.
- Wiadomości światowe: 3-4 najważniejsze tematy, po 2 zdania każdy
- Wiadomości Polska: 3-4 tematy, po 2 zdania
- Wiadomości lokalne Kęty/Bielsko: wszystko co jest, krótko
- Darmowe gry: PEŁNY ARTYKUŁ — tytuł, opis, platforma, do kiedy bezpłatna
- Gaming news: tylko tytuły i jedno zdanie opisu

DANE ŹRÓDŁOWE:

🌤 POGODA KĘTY:
{weather_kety[:1800]}

🌤 POGODA BIELSKO-BIAŁA:
{weather_bb[:1800]}

🌍 WIADOMOŚCI ŚWIATOWE:
{news_world}

🇵🇱 WIADOMOŚCI POLSKA:
{news_poland}

📍 WIADOMOŚCI LOKALNE:
{news_local[:2000]}

🎮 GAMING / DARMOWE GRY (RSS):
{news_gaming}

🎁 EPIC GAMES DARMOWE GRY:
{epic_games}

🎁 GOG DARMOWE GRY:
{gog_free}

Napisz teraz kompletny e-mail HTML zaczynając od <!DOCTYPE html> i kończąc na </html>.
Dodaj na górze pogrubiony nagłówek "☀️ Poranny raport — {today}".
Użyj czcionki bezszeryfowej, czytelnych odstępów, jasnego tła.
"""

    print('Generuję raport przez Claude Sonnet...')
    email_body = call_claude(prompt)

    if '<!DOCTYPE' in email_body:
        start = email_body.find('<!DOCTYPE')
        email_body = email_body[start:]
    elif '<html' in email_body:
        start = email_body.find('<html')
        email_body = email_body[start:]

    print('Wysyłam e-mail...')
    send_email(f'☀️ Poranny raport — {today}', email_body)
    print('Gotowe!')


if __name__ == '__main__':
    main()
