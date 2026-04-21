from enum import Enum

class FormShowConstants(Enum):
    # 6.1.1.1 FormShowConstants
    vbModal = 1
    vb_modeless = 0
  
class VbAppWinStyle(Enum):
    # 6.1.1.2 VbAppWinStyle
    vbHide = 0
    vbMaximizedFocus = 3
    vbMinimuzedFocus = 2
    vbMinimizedNoFocus = 6
    vbNormalFocus = 1
    vbNormalNoFocus = 4

class VbCallType(Enum):
    # 6.1.1.4 VbCallType
    vbGet = 2
    vbLet = 4
    vbMethod = 1
    vbSet = 8

class VbDayOfWeek(Enum):
    # 6.1.1.7 VbDayOfWeek
    vbsunday = 1
    vbMonday = 2
    vbTuesday = 3
    vbWednesday = 4
    vbThursday = 5
    vbFriday = 6
    vbSaturday = 7
    vbUseSystemDayOfWeek = 0

class VbMsgBoxResult(Enum):
    # 6.1.1.11 VbMsgBoxResult
    vbAbort = 3
    vbCancel = 2
    vbIgnore = 5
    vbNo = 7
    vbOK = 1
    vbRetry = 4
    vbYes = 6

class VBMsgBoxStyle(Enum):
    # 6.1.1.12 VbMsgBoxStyle
    vbOKOnly = 0

class VbVarType(Enum):
    # 6.1.1.16 VbVarType
    vbArray = 8192
    vbBoolean = 11
    vbByte = 17
    vbCurrency = 6
    vbDataObject = 13
    vbDate = 7
    vbDecimal = 14
    vbDouble = 5
    vbEmpty = 0
    vbError = 10
    vbInteger = 2
    vbLong = 3
    # defined only on implementations
    # that support a LongLong value type
    vbLongLong = 20
    vbNull = 1
    vbObject = 9
    vbSingle = 4
    vbString = 8
    vbUserDefinedType = 36
    vbVariant = 12
