"""
카카오 토큰 최초 발급 스크립트. 1회만 실행하면 됨.

실행 전 준비:
  1. https://developers.kakao.com 접속 → 로그인
  2. [내 애플리케이션] → [애플리케이션 추가]
     - 앱 이름: 아무거나 (예: bgsap)
  3. [앱 설정] → [앱 키] → REST API 키 복사
  4. [제품 설정] → [카카오 로그인] → 활성화 ON
  5. [제품 설정] → [카카오 로그인] → [Redirect URI]
     → http://localhost 추가
  6. [제품 설정] → [카카오 로그인] → [동의항목]
     → "카카오톡 메시지 전송" 체크 (선택 동의)

실행:
  python kakao_setup.py
"""

import json
import os
import webbrowser
import requests

TOKENS_FILE = os.path.join(os.path.dirname(__file__), "kakao_tokens.json")
AUTH_URL    = "https://kauth.kakao.com/oauth/authorize"
TOKEN_URL   = "https://kauth.kakao.com/oauth/token"
REDIRECT    = "http://localhost"


def main():
    print("=" * 50)
    print("  카카오 토큰 발급 설정")
    print("=" * 50)

    rest_api_key = input("\n[1] REST API 키를 입력하세요: ").strip()
    client_secret = input("[1-1] 클라이언트 시크릿 코드 (없으면 그냥 Enter): ").strip()

    # 브라우저로 카카오 로그인 페이지 열기
    auth_link = (
        f"{AUTH_URL}"
        f"?client_id={rest_api_key}"
        f"&redirect_uri={REDIRECT}"
        f"&response_type=code"
        f"&scope=talk_message"
    )
    print(f"\n[2] 아래 링크를 브라우저(카카오 로그인된 창)에 복사해서 여세요:")
    print(f"\n    {auth_link}\n")
    print(f"    로그인 후 주소가 'http://localhost/?code=...' 형태로 바뀜")
    print(f"    그 전체 주소를 복사하세요.")
    input("    준비되면 Enter...")

    redirected_url = input("\n[3] 이동된 전체 주소를 붙여넣으세요: ").strip()

    # code 파라미터 추출
    if "code=" not in redirected_url:
        print("❌ 주소에 code= 가 없습니다. 다시 시도하세요.")
        return

    code = redirected_url.split("code=")[-1].split("&")[0]

    # 토큰 발급
    data = {
        "grant_type":   "authorization_code",
        "client_id":    rest_api_key,
        "redirect_uri": REDIRECT,
        "code":         code,
    }
    if client_secret:
        data["client_secret"] = client_secret
    resp = requests.post(TOKEN_URL, data=data, timeout=10)

    if resp.status_code != 200:
        print(f"❌ 토큰 발급 실패: {resp.status_code} {resp.text}")
        return

    result = resp.json()
    tokens = {
        "rest_api_key":   rest_api_key,
        "client_secret":  client_secret,
        "access_token":   result["access_token"],
        "refresh_token":  result["refresh_token"],
    }

    with open(TOKENS_FILE, "w", encoding="utf-8") as f:
        json.dump(tokens, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 토큰 저장 완료: {TOKENS_FILE}")
    print("이제 카카오 전송 기능이 자동으로 작동합니다.")


if __name__ == "__main__":
    main()
