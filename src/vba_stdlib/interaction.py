from typing import Any
from vba_stdlib.enums import VbMsgBoxResult, VbMsgBoxStyle


class Interaction:

    @staticmethod
    def msgbox(prompt: Any,
               buttons: VbMsgBoxStyle = VbMsgBoxStyle.vbokonly,
               title: str = "",
               help_file: str = "",
               context: int | None = None) -> VbMsgBoxResult:
        print(str(prompt))
        return VbMsgBoxResult.vbok
