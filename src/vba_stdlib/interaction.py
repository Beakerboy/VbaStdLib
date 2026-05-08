from typing import Any
from vba_stdlib.enums import VbCallType, VbMsgBoxResult, VbMsgBoxStyle


class Interaction:

    @staticmethod
    def library() -> dict:
        return {
            "name": "interaction",
            "type": FunctionType.MODULE,
            "functions": {
                "msgbox": {
                    "name": "msgbox",
                    "type": FunctionType.FUNCTION,
                    "handle": getattr(Interaction, "msgbox"),
                }
            }
        }

    @staticmethod
    def CallByName(obj: object, proc_name: str, call_type: VbCallType, args: list[Any]) -> Any:
        pass

    @staticmethod
    def Choose(index, choice) -> Any:
        pass

    @staticmethod
    def CreateObject() -> object:
        pass

    @staticmethod
    def MsgBox (prompt: Any,
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
