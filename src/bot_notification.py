import os
import requests
from requests.adapters import HTTPAdapter


class TelegramNotifier:
    def __init__(self, chat_id: str, session: requests.Session, token: str = None, retries: int = 5):
        self.token: str = os.getenv('TOKEN') if not token else token
        self.api_params = {'chat_id': chat_id, 'parse_mode': 'Markdown'}
        self.retries = retries
        self.session = session
        self.session.mount('http://', HTTPAdapter(max_retries=self.retries))

    def send_message(self, message: str, is_document: bool = False) -> requests.models.Response:
        message_type = 'sendDocument' if is_document else 'sendMessage'
        api_url = f'https://api.telegram.org/bot{self.token}/{message_type}'
        args = {'url': api_url, 'params': self.api_params, 'json': {'text': message}}
        if is_document:
            args['files'] = {'document': message}
        return self.session.post(**args)
