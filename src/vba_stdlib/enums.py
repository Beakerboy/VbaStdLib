from enum import Enum
from vba_types import VBALong


class FormShowConstants(Enum):
    # 6.1.1.1 FormShowConstants
    vbmodal = 1
    vbmodeless = 0


class VbAppWinStyle(Enum):
    # 6.1.1.2 VbAppWinStyle
    vbhide = 0
    vbmaximizedfocus = 3
    vbminimuzedfocus = 2
    vbminimizednofocus = 6
    vbnormalfocus = 1
    vbnormalnofocus = 4


# 6.1.1.3


class VbCallType(Enum):
    # 6.1.1.4 VbCallType
    vbget = 2
    vblet = 4
    vbmethod = 1
    vbset = 8


# 6.1.1.5


# 6.1.1.6


class VbDayOfWeek(Enum):
    # 6.1.1.7 VbDayOfWeek
    vbsunday = 1
    vbmonday = 2
    vbtuesday = 3
    vbwednesday = 4
    vbthursday = 5
    vbfriday = 6
    vbsaturday = 7
    vbusesystemdayofweek = 0


# 6.1.1.8


# 6.1.1.9


# 6.1.1.10


class VbMsgBoxResult(Enum):
    # 6.1.1.11 VbMsgBoxResult
    vbabort = VBALong(3)
    vbcancel = VBALong(2)
    vbignore = VBALong(5)
    vbno = VBALong(7)
    vbok = VBALong(1)
    vbretry = VBALong(4)
    vbyes = VBALong(6)


class VbMsgBoxStyle(Enum):
    # 6.1.1.12 VbMsgBoxStyle
    vbabortretryignore = VBALong(2)
    vbapplicationmodal = VBALong(0)
    vbcritical = VBALong(16)
    vbdefaultbutton1 = VBALong(0)
    vbdefaultbutton2 = VBALong(256)
    vbdefaultbutton3 = VBALong(512)
    vbdefaultbutton4 = VBALong(768)
    vbexclamation = VBALong(48)
    vbdinformation = VBALong(64)
    vbmsgboxhelpbutton = VBALong(16384)
    vbmsgboxright = VBALong(524288)
    vbmsgboxrtlreading = VBALong(1048576)
    vbokonly = VBALong(0)
    vbquestion = VBALong(32)
    vbretrycancel = VBALong(5)
    vbsystemmodal = VBALong(4096)
    vbyesno = VBALong(4)
    vbyesnocencel = VBALong(3)


# 6.1.1.13


# 6.1.1.14


# 6.1.1.15


class VbVarType(Enum):
    # 6.1.1.16 VbVarType
    vbarray = 8192
    vbboolean = 11
    vbbyte = 17
    vbcurrency = 6
    vbdataobject = 13
    vbdate = 7
    vbdecimal = 14
    vbdouble = 5
    vbempty = 0
    vberror = 10
    vbinteger = 2
    vblong = 3
    # defined only on implementations
    # that support a LongLong value type
    vblonglong = 20
    vbnull = 1
    vbobject = 9
    vbsingle = 4
    vbstring = 8
    vbuserdefinedtype = 36
    vbvariant = 12
