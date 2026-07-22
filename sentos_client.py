import base64
import json
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


class SentosError(Exception):
    pass


class SentosClient:
    def __init__(self, base_url, username, password, timeout=25):
        self.base_url = base_url.strip().rstrip('/')
        if not self.base_url.endswith('/api'):
            self.base_url += '/api'
        token = base64.b64encode(f'{username}:{password}'.encode('utf-8')).decode('ascii')
        self.headers = {'Authorization': f'Basic {token}', 'Accept': 'application/json', 'Content-Type': 'application/json'}
        self.timeout = timeout

    def _request(self, method, path, params=None, body=None):
        url = self.base_url + path
        if params:
            url += '?' + urlencode(params)
        data = json.dumps(body).encode('utf-8') if body is not None else None
        try:
            with urlopen(Request(url, data=data, headers=self.headers, method=method), timeout=self.timeout) as res:
                raw = res.read().decode('utf-8', errors='replace')
                return json.loads(raw) if raw else {}
        except HTTPError as e:
            raw = e.read().decode('utf-8', errors='replace')
            try: detail = json.loads(raw)
            except Exception: detail = raw
            if e.code == 401: raise SentosError('Kullanıcı adı veya parola kabul edilmedi (401).')
            if e.code == 403: raise SentosError('Bu API işlemi için yetki yok (403).')
            raise SentosError(f'Sentos HTTP {e.code}: {detail}')
        except URLError as e:
            raise SentosError(f'Sentos adresine ulaşılamadı: {e.reason}')
        except TimeoutError:
            raise SentosError('Sentos bağlantısı zaman aşımına uğradı.')

    def products(self, size=100, page=1, sku=None):
        params = {'size': size, 'page': page}
        if sku: params['sku'] = sku
        return self._request('GET', '/products', params=params)


def product_list(payload):
    if isinstance(payload, list): return payload
    if not isinstance(payload, dict): return []
    if 'id' in payload and 'sku' in payload: return [payload]
    for key in ('data', 'items', 'products', 'results'):
        value = payload.get(key)
        if isinstance(value, list): return value
        if isinstance(value, dict):
            for nested in ('data', 'items', 'products', 'results'):
                if isinstance(value.get(nested), list): return value[nested]
    return []
