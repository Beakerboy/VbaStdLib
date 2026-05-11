from typing import Any
from vba_stdlib.enums import VbCallType, VbMsgBoxResult, VbMsgBoxStyle


class Interaction:

    @staticmethod
    def callbyname(obj: object, proc_name: str, call_type: VbCallType, args: list[Any]) -> Any:
        pass

    @staticmethod
    def choose(index, choice) -> Any:
        pass

    @staticmethod
    def createobject() -> object:
        pass

    @staticmethod
    def msgbox (prompt: Any,
               buttons: VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly,
               title: str = "",
               help_file: str = "",
               context: int | None = None
              ) -> VbMsgBoxResult:
        print(str(prompt))
        return VbMsgBoxResult.vbOK

    @staticmethod
    def DoEvents() -> None:
        pass

    @staticmethod
    def Environ() -> None:
        pass
