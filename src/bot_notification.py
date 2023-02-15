import requests
from requests.adapters import HTTPAdapter


class TelegramNotifier:
    def __init__(self, token: str, chat_id: str, session: requests.Session, retries: int = 5):
        self.token: str = token
        self.api_params = {'chat_id': chat_id, 'parse_mode': 'Markdown'}
        self.retries = retries
        self.session = session
        self.session.mount('http://', HTTPAdapter(max_retries=self.retries))

    def send_message(self, message: str, is_document: bool = False) -> requests.models.Response:
        message_type = 'sendDocument' if is_document else 'sendMessage'
        api_url = f'https://api.telegram.org/bot{self.token}/{message_type}'
        args = {'url': api_url, 'params': self.api_params, 'json': {'text': message}}
        if is_document:
            args['files'] = {'document': open(file=message, mode='rb')}
        return self.session.post(**args)


if __name__ == '__main__':
    import dotenv
    from data_structures import Notifiers
    import os

    dotenv.load_dotenv()
    with requests.Session() as session:
        notifiers = Notifiers(
            log=TelegramNotifier(token=os.getenv('TOKEN_LOG'), chat_id=os.getenv(f'CHAT_ID_LOG'), session=session),
            alert=TelegramNotifier(token=os.getenv('TOKEN_ALERT'), chat_id=os.getenv(f'CHAT_ID_ALERT'), session=session)
        )

        # notifiers.log.send_message(message='Log test message from bot')
        # notifiers.log.send_message(message=r'C:\Users\robot.ad\Desktop\2207_09.xlsb', is_document=True)

        notifiers.alert.send_message(message='Alert test message from bot')
        notifiers.alert.send_message(message=r'C:\Users\robot.ad\Desktop\2207_09.xlsb', is_document=True)
