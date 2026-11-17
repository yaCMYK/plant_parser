"""Conservative URL routing and content isolation, including Greeninfo cards."""
import re
from urllib.parse import urlparse
from bs4 import BeautifulSoup

MEDIA = re.compile(r'\.(?:jpe?g|png|gif|webp|svg|pdf|zip|mp4|css|js|ico)(?:$|/)', re.I)

def greeninfo(url):
    return (urlparse(url).hostname or '').lower() in {'greeninfo.ru', 'www.greeninfo.ru'}

def card_url(url):
    path = urlparse(url).path
    return bool(re.fullmatch(r'/[a-z0-9_/.-]+/[a-z]+_[a-z0-9_-]+\.html', path, re.I)) and not path.endswith('_art.html')

def html_link(url):
    return urlparse(url).scheme in {'http', 'https'} and not MEDIA.search(urlparse(url).path)

def allowed_link(url):
    if not html_link(url):
        return False
    if greeninfo(url):
        return card_url(url)
    return not re.search(r'/(?:cart|checkout|login|account|privacy|terms|about|contact|shipping)(?:[/.]|$)', urlparse(url).path, re.I)

def content_soup(soup):
    # Work on a copy: link discovery still needs the complete page.
    clone = BeautifulSoup(str(soup), 'lxml')
    for node in clone.select('script, style, nav, header, footer, aside, noscript, [role=navigation], .breadcrumb, .breadcrumbs, .menu, .sidebar, .related, .comments'):
        node.decompose()
    for selector in ('[itemtype*="Product"]', '.plant-description', '.plant-card', '#plant', 'article', 'main', '#content', '.entry-content', '.content'):
        nodes = clone.select(selector)
        if nodes:
            node = max(nodes, key=lambda n: len(n.get_text(' ', strip=True)))
            result = BeautifulSoup(str(node), 'lxml')
            if not result.find('h1') and clone.find('h1'):
                result.body.insert(0, BeautifulSoup(str(clone.find('h1')), 'lxml').find('h1'))
            return result
    return clone
