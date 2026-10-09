import json
import threading
import unittest
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from pathlib import Path
from webapp import make_server

class BrowserTests(unittest.TestCase):
    def setUp(self):
        self.server, self.token = make_server()
        self.url = f'http://127.0.0.1:{self.server.server_port}'
        self.thread = threading.Thread(target=self.server.serve_forever)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def post(self, endpoint, data=None, token=None, origin=None):
        headers = {'Content-Type': 'application/json', 'X-Session-Token': token or self.token}
        if origin: headers['Origin'] = origin
        req = Request(self.url + '/api/' + endpoint, json.dumps(data or {}).encode(), headers)
        return json.loads(urlopen(req).read())

    def test_page_and_status(self):
        with urlopen(self.url) as response:
            page = response.read().decode('utf-8')
        self.assertIn(self.token, page)
        self.assertNotIn('__TOKEN__', page)
        self.assertIn('lang="en"', page)
        self.assertIn('Create narrated PPTX', page)
        self.assertFalse(self.post('status')['busy'])

    def test_cross_origin_and_token_rejected(self):
        for kwargs in ({'token': 'invalid'}, {'origin': 'https://example.com'}):
            with self.assertRaises(HTTPError) as e:
                self.post('status', **kwargs)
            self.assertEqual(e.exception.code, 403)

    def test_docx_preview(self):
        sample = Path(__file__).parent.parent / 'Ornek_Almanca_Metin.docx'
        if not sample.exists(): self.skipTest('Sample not present')
        result = self.post('preview', {'doc': str(sample)})
        self.assertEqual(set(result['slides']), {'1', '2', '3'})

    def test_invalid_start(self):
        with self.assertRaises(HTTPError) as e:
            self.post('start', dict.fromkeys(('ppt', 'doc', 'model', 'target'), 'missing'))
        self.assertEqual(e.exception.code, 400)
        self.assertFalse(self.post('status')['busy'])

if __name__ == '__main__': unittest.main()
