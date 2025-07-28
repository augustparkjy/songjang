
Product = {
    'egg': {
        'coupang': [8643526485],
        'alwayz': ['66c5f527b550787272789fa2'],
        'toss': [26662045, 9300515]
    },
    'rice': {
        # 6491027744e70ba232b4fbac : 곡물 팝
        'alwayz': ['6446a65f36b5f4ba2e85b08d', '6491027744e70ba232b4fbac']
    }
}

Shipping = {
    'coupang': '쿠팡',
    'alwayz': '올웨이즈',
    'naver': '네이버',
    'toss': '토스'
}

Format = {
    'product_id': {
        'coupang': '노출상품ID',
        'alwayz': '상품아이디',
        'toss': '상품ID'
    },
    'product_name': { 
        'coupang': '등록상품명',
        'alwayz': '상품명',
        'toss': '상품명'
    },
    'product_code': {
        'alwayz': '판매자 상품코드'
    },
    'options': {
        'coupang': '등록옵션명',
        'alwayz': '옵션',
        'toss': '옵션'
    },
    'receiver': {
        'coupang': '수취인이름',
        'alwayz': '수령인',
        'toss': '수령인명'
    },
    'count': {
        'coupang': '구매수(수량)',
        'alwayz': '수량',
        'toss': '수량'
    },
    'contact': {
        'coupang': '수취인전화번호',
        'alwayz': '수령인 연락처',
        'toss': '수령인 연락처'
    },
    'address': {
        'coupang': '수취인 주소',
        'alwayz': '주소',
        'toss': '주소'
    },
    'message': {
        'coupang': '배송메세지',
        'alwayz': '수령 방법',
        'toss': '요청사항'
    },
}

class Item:  
    def __init__(self, 
                receiver: str, 
                contact: str, 
                address: str, 
                quantity: int, 
                note: str, 
                shipping_number: str, 
                receiver_manager: str = '', 
                receiver_phone: str = '', 
                zip_code: str = '', 
                item_name: str = '',
                freight: str = '',
                payment_condition: str = '',
                tracking_number: str = '',
                empty: str = None):
        self.receiver = receiver
        self.contact = contact
        self.receiver_manager = receiver_manager
        self.receiver_phone = receiver_phone
        self.zip_code = zip_code
        self.address = address
        self.quantity = quantity
        self.item_name = item_name
        self.freight = freight
        self.payment_condition = payment_condition
        self.shipping_number = shipping_number
        self.note = note
        self.tracking_number = tracking_number
        self.empty = None
