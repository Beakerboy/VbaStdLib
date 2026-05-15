from typing import Any
from vba_stdlib.enums import VbCallType, VbMsgBoxResult, VbMsgBoxStyle


class Interaction:

    @staticmethod
    def callbyname(obj: object,
                   proc_name: str,
                   call_type: VbCallType,
                   args: list[Any]) -> Any:
        pass

    @staticmethod
    def choose(index: Any, choice: Any) -> Any:
        pass

    @staticmethod
    def createobject() -> object:
        pass

    @staticmethod
    def msgbox(prompt: Any,
               buttons: VbMsgBoxStyle = VbMsgBoxStyle.vbokonly,
               title: str = "",
               help_file: str = "",
               context: int | None = None) -> VbMsgBoxResult:
        print(str(prompt))
        return VbMsgBoxResult.vbok

    @staticmethod
    def doevents() -> None:
        pass

    @staticmethod
    def environ() -> None:
        pass
