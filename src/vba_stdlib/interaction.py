from typing import Any
from vba_stdlib.enums import VbCallType, VbMsgBoxResult


class Interaction:

    @staticmethod
    def CallByName(obj: object, proc_name: str, call_type: VbCallType, args: list[Any]) -> Any:
        pass

    @staticmethod
    def Choose(index, choice) -> Any:
        pass

    @staticmethod
    def CreateObject() -> Object:
        pass

    @staticmethod
    def MsgBox(prompt: Any, buttons, title, help_file, context) -> VbMsgBoxResult:
        print(prompt)
        return VbMsgBoxResult.vbOK

    @staticmethod
    def DoEvents() -> None:
        pass

    @staticmethod
    def Environ() -> None:
        pass
