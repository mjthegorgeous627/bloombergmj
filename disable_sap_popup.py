"""
SAP Logon 스크립팅 보안 팝업 비활성화.
'스크립트가 GUI에 연결되고 있습니다' 팝업이 뜨지 않도록 설정.
한 번만 실행하면 됩니다.
"""
import winreg

KEYS_TO_TRY = [
    (r"Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting", "NotifyUser", "0"),
    (r"Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting", "EnableNotification", "0"),
    (r"Software\SAP\General\Scripting", "NotifyUser", "0"),
]

for reg_path, value_name, value_data in KEYS_TO_TRY:
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, value_data)
        winreg.CloseKey(key)
        print(f"설정 완료: {reg_path}\\{value_name} = {value_data}")
    except Exception as e:
        print(f"건너뜀: {reg_path}\\{value_name} → {e}")

print("\n완료! SAP Logon을 재시작한 후 startup.py를 실행하세요.")
