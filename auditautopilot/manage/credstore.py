"""윈도우 '자격 증명 관리자'에 메일 비밀번호 저장 (표준 라이브러리 ctypes, 윈도우 전용).

다른 운영체제이거나 저장에 실패하면 아무것도 하지 않고, 비밀번호는 프로그램이 켜져 있는 동안만 기억된다.
저장된 값은 '제어판 > 자격 증명 관리자 > Windows 자격 증명'에서 'auditautopilot:메일주소'로 보이고 지울 수 있다.
"""

from __future__ import annotations

import sys

PREFIX = "auditautopilot:"


def _api():
    import ctypes
    from ctypes import wintypes

    class FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", wintypes.DWORD), ("dwHighDateTime", wintypes.DWORD)]

    class CREDENTIAL(ctypes.Structure):
        _fields_ = [
            ("Flags", wintypes.DWORD), ("Type", wintypes.DWORD), ("TargetName", wintypes.LPWSTR),
            ("Comment", wintypes.LPWSTR), ("LastWritten", FILETIME), ("CredentialBlobSize", wintypes.DWORD),
            ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)), ("Persist", wintypes.DWORD),
            ("AttributeCount", wintypes.DWORD), ("Attributes", ctypes.c_void_p),
            ("TargetAlias", wintypes.LPWSTR), ("UserName", wintypes.LPWSTR),
        ]

    advapi = ctypes.WinDLL("advapi32", use_last_error=True)
    advapi.CredWriteW.argtypes = [ctypes.POINTER(CREDENTIAL), wintypes.DWORD]
    advapi.CredWriteW.restype = wintypes.BOOL
    advapi.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                 ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
    advapi.CredReadW.restype = wintypes.BOOL
    advapi.CredFree.argtypes = [ctypes.c_void_p]
    advapi.CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
    advapi.CredDeleteW.restype = wintypes.BOOL
    return ctypes, CREDENTIAL, advapi


CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2


def save(user: str, password: str) -> str:
    if sys.platform != "win32":
        return "윈도우가 아니어서 이 프로그램이 켜져 있는 동안만 기억합니다."
    try:
        ctypes, CREDENTIAL, advapi = _api()
        blob = password.encode("utf-16-le")
        buf = (ctypes.c_ubyte * len(blob)).from_buffer_copy(blob)
        cred = CREDENTIAL()
        cred.Type = CRED_TYPE_GENERIC
        cred.TargetName = PREFIX + user
        cred.UserName = user
        cred.CredentialBlobSize = len(blob)
        cred.CredentialBlob = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        cred.Persist = CRED_PERSIST_LOCAL_MACHINE
        if not advapi.CredWriteW(ctypes.byref(cred), 0):
            return f"자격 증명 관리자 저장 실패(오류 {ctypes.get_last_error()}). 켜져 있는 동안만 기억합니다."
        return "윈도우 자격 증명 관리자에 저장했습니다. 다음에 켤 때 다시 넣지 않아도 됩니다."
    except Exception as exc:
        return f"자격 증명 관리자를 쓸 수 없습니다({exc}). 켜져 있는 동안만 기억합니다."


def load(user: str) -> str | None:
    if sys.platform != "win32" or not user:
        return None
    try:
        ctypes, CREDENTIAL, advapi = _api()
        ptr = ctypes.POINTER(CREDENTIAL)()
        if not advapi.CredReadW(PREFIX + user, CRED_TYPE_GENERIC, 0, ctypes.byref(ptr)):
            return None
        try:
            c = ptr.contents
            return ctypes.string_at(c.CredentialBlob, c.CredentialBlobSize).decode("utf-16-le")
        finally:
            advapi.CredFree(ptr)
    except Exception:
        return None


def delete(user: str) -> None:
    if sys.platform != "win32":
        return
    try:
        _, _, advapi = _api()
        advapi.CredDeleteW(PREFIX + user, CRED_TYPE_GENERIC, 0)
    except Exception:
        pass
