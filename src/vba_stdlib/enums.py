from enum import Enum

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
    vbabort = 3
    vbcancel = 2
    vbignore = 5
    vbno = 7
    vboK = 1
    vbretry = 4
    vbyes = 6


class VbMsgBoxStyle(Enum):
    # 6.1.1.12 VbMsgBoxStyle
    vbokonly = 0


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
