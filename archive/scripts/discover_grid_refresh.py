"""
NWBC 8.00 ALV 그리드 툴바 버튼 ID 탐색 (비대화형).

refresh_sap_list()의 grid.pressToolbarButton("&REFRESH")가 NWBC 8.00에서
COM 예외를 던지는 문제를 진단하기 위해, VL06O/VL10G/ZRMA_Q 세션이 이미
떠 있는 상태에서 실행하여 각 ALV 그리드의 실제 툴바 버튼 ID/텍스트/툴팁을
출력한다.

sap_handler.get_scripting_engine()의 캐시 패턴을 그대로 사용해 새 COM
연결을 반복 호출하지 않는다 (반복 GetObject 호출이 엔진 객체를 깨뜨리는
문제 회피).

사용법:
  python discover_grid_refresh.py
"""

import sys

from sap_handler import get_sap_session, GRID_PATH
from vl10g_handler import VL10G_GRID_PATH
from zrma_handler import ZRMA_GRID_PATH

TARGETS = [
    (0, GRID_PATH, "VL06O"),
    (1, VL10G_GRID_PATH, "VL10G"),
    (2, ZRMA_GRID_PATH, "ZRMA_Q RLKR"),
    (3, ZRMA_GRID_PATH, "ZRMA_Q Q2"),
]


def dump_grid_toolbar(grid, label):
    print(f"\n=== {label} ===")
    try:
        count = grid.GetToolbarButtonCount()
    except Exception as e:
        print(f"  [FAIL] GetToolbarButtonCount(): {e}")
        return

    print(f"  버튼 수: {count}")
    for i in range(count):
        try:
            btn_id = grid.GetToolbarButtonId(i)
        except Exception as e:
            btn_id = f"(id 읽기 실패: {e})"
        try:
            btn_text = grid.GetToolbarButtonText(i)
        except Exception:
            btn_text = ""
        try:
            btn_tip = grid.GetToolbarButtonTooltip(i)
        except Exception:
            btn_tip = ""
        try:
            btn_type = grid.GetToolbarButtonType(i)
        except Exception:
            btn_type = ""
        print(f"  [{i}] id='{btn_id}' text='{btn_text}' tooltip='{btn_tip}' type={btn_type}")

    # 원래 코드가 시도하던 ID가 지금도 존재하는지 명시적으로 확인
    try:
        grid.pressToolbarButton("&REFRESH")
        print("  [OK] pressToolbarButton('&REFRESH') 성공 (문제가 이미 해결됐거나 재현 안 됨)")
    except Exception as e:
        print(f"  [FAIL] pressToolbarButton('&REFRESH'): {e}")


def main():
    for session_idx, grid_path, label in TARGETS:
        try:
            session = get_sap_session(session_idx, retries=1)
        except Exception as e:
            print(f"\n=== {label} (세션{session_idx}) ===")
            print(f"  [FAIL] 세션 연결 실패: {e}")
            continue

        try:
            grid = session.findById(grid_path)
        except Exception as e:
            print(f"\n=== {label} (세션{session_idx}) ===")
            print(f"  [FAIL] 그리드 findById 실패 (path={grid_path}): {e}")
            continue

        dump_grid_toolbar(grid, f"{label} (세션{session_idx}, path={grid_path})")

    print("\n완료. 위 목록에서 새로고침 관련 id(예: '&REFRESH', 'REFRESH', 'Reload' 등)를 "
          "찾아 sap_handler.refresh_sap_list()의 pressToolbarButton() 인자를 갱신하세요.")


if __name__ == "__main__":
    sys.exit(main() or 0)
