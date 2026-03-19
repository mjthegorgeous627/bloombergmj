"""Zebra 라벨 출력 테스트."""
from zebra_handler import print_label

# 실제 데이터로 교체해서 테스트
print_label(
    qr_data="https://bsp.btogo.com/example/order/12345",
    name="HYEONJEONG LEE",
    company="SHINHAN ASSET MANAGEMENT",
)
