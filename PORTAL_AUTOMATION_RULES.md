# Portal Automation Rules

Last updated: 2026-05-27

## Default handling

- 기본값은 serial number가 있는 배송이다.
- Serial Number가 있으면 반드시 serial 등록 화면에서 실제 serial row가 남는지 확인한 뒤 Run ShipERP를 누른다.
- QR/ZPL 인쇄는 항상 Packing Post 이후에 실행한다.

## No-serial / material-only exception

- Serial 없는 케이스는 예외다.
- 사용자가 `serial 없음`, `material only`, `material만 있음`, `no serial`이라고 명시하거나, 이미지/엑셀에 Material/Qty만 있고 S/N이 비어 있는 경우에만 material-only로 처리한다.
- material-only 흐름은 Pick Qty 입력 -> Run ShipERP -> Packing Post -> QR/ZPL 인쇄 순서다.
- 런처에서는 `Material only / no serial` 체크박스를 켠 경우에만 `--material-only`를 넘긴다.

## Excel lookup understanding

- 배송장 엑셀은 날짜별 시트 구조다.
- 현재 조회는 전체 날짜 시트를 매번 도는 것이 아니라 오늘 날짜 시트만 대상으로 한다.
- 그래도 매번 workbook을 여는 비용이 있으므로, 속도 개선은 오늘 시트를 한 번 읽어 order cache를 만들고 포털 처리 때 cache에서 가져오는 방향이 맞다.
- 캐시에서도 기본은 `serial_required=true`로 두고, serial 칸이 비어 있으며 material/qty만 있는 경우에만 `material_only=true`로 표시한다.
