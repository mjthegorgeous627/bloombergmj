"""카카오 전송 테스트 - 나에게 보내기"""
from kakao_handler import send_kakao_message

test_msg = """키보드 교체
배송 키보드5 10063421
회수 일반키보드 10045196
CHULMIN KANG / 02-2004-9908
SHINYOUNG SECURITIES CO LTD 34-8 YOIDO-DONG"""

result = send_kakao_message(test_msg)
print("성공" if result else "실패")
