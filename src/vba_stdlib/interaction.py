from typing import Any, TypeVar
from vba_stdlib.enums import VbMsgBoxResult, VbMsgBoxStyle
from vba_types import VBATypeBase, VBAString, VBALong


T = TypeVar('T', bound='Interaction')


class Interaction:

    @staticmethod
    def msgbox(prompt: VBATypeBase,
               buttons: VbMsgBoxStyle = VbMsgBoxStyle.vbokonly,
               title: VBATypeBase | None = None,
               help_file: VBAString | None = None,
               context: VBALong | None = None) -> VbMsgBoxResult:

        if title is None:
            title = "Microsoft Excel"
        else:
            title = str(title)
        if help_file is not None and context is None:
            # raise some sort of error
            pass
        number_type = VBALong(buttons.value.value & 7)
        
        if number_type == VbMsgBoxStyle.vbokonly.value:
            button_text = "OK"
        elif number_type == VbMsgBoxStyle.vbokcancel.value:
            button_text = "OK    Cancel"
        elif number_type == VbMsgBoxStyle.vbabortretryignore.value:
            button_text = "Abort    Retry    Ignore"
        elif number_type == VbMsgBoxStyle.vbyesnocancel.value:
            button_text = "Yes    No    Cancel"
        elif number_type == VbMsgBoxStyle.vbyesno.value:
            button_text = "Yes    No"
        elif number_type == VbMsgBoxStyle.vbretrycancel.value:
            button_text = "Retry    Cancel"
        else:
            # what kind of exception should this be?
            raise Exception("unknown button type")

        icon = VBALong(buttons.value.value & 112)
        window_icon = "X"
        if icon == VBALong(0):
            window_icon = ""
        elif icon == VbMsgBoxStyle.vbcritical.value:
            window_icon = "X"
        elif icon == VbMsgBoxStyle.vbquestion.value:
            window_icon = "?"
        elif icon == VbMsgBoxStyle.vbexclamation.value:
            window_icon = "!"
        elif icon == VbMsgBoxStyle.vbinformation.value:
            window_icon = "i"
        else:
            # what kind of exception should this be?
            raise Exception("unknown icon type")

        default = VBALong(buttons.value.value & 768)
        if default == VbMsgBoxStyle.vbdefaultbutton1.value:
            default_button = 1
        elif default == VbMsgBoxStyle.vbdefaultbutton2.value:
            default_button = 2
        elif default == VbMsgBoxStyle.vbdefaultbutton3.value:
            default_button = 3
        else:
            # default == VbMsgBoxStyle.vbdefaultbutton4:
            # are there ever 4 buttons? Is "Help" a fourth?
            # Help Button does not return. Another button must be pressed
            default_button = 4
        print(f"{title}\n{icon}\n{prompt}\n{button_text}")
        # input = input()
        # What to do if we receive an option not presented?
        # F1 is help
        # Esc is cancel
        return VbMsgBoxResult.vbok
