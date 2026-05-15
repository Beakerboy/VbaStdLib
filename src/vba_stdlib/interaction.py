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
        number_type = buttons.value.value & 7
        
        if number_type == VbMsgBoxStyle.vbokonly.value.value:
            button_list = [VbMsgBoxStyle.vbok]
            button_text = "OK"
        elif number_type == VbMsgBoxStyle.vbokcancel.value.value:
            button_list = [VbMsgBoxStyle.vbok, VbMsgBoxStyle.vbcancel]
            button_text = "OK    Cancel"
        elif number_type == VbMsgBoxStyle.vbabortretryignore.value.value:
            button_text = "Abort    Retry    Ignore"
        elif number_type == VbMsgBoxStyle.vbyesnocancel.value.value:
            button_text = "Yes    No    Cancel"
        elif number_type == VbMsgBoxStyle.vbyesno.value.value:
            button_text = "Yes    No"
        elif number_type == VbMsgBoxStyle.vbretrycancel.value.value:
            button_text = "Retry    Cancel"
        else:
            # what kind of exception should this be?
            raise Exception("unknown button type")

        icon = buttons.value.value & 112
        window_icon = "X"
        if icon == 0:
            window_icon = ""
        elif icon == VbMsgBoxStyle.vbcritical.value.value:
            window_icon = "X"
        elif icon == VbMsgBoxStyle.vbquestion.value.value:
            window_icon = "?"
        elif icon == VbMsgBoxStyle.vbexclamation.value.value:
            window_icon = "!"
        elif icon == VbMsgBoxStyle.vbinformation.value.value:
            window_icon = "i"
        else:
            # what kind of exception should this be?
            raise Exception("unknown icon type")

        default = buttons.value.value & 768
        if default == VbMsgBoxStyle.vbdefaultbutton1.value.value:
            default_button = 1
        elif default == VbMsgBoxStyle.vbdefaultbutton2.value.value:
            default_button = 2
        elif default == VbMsgBoxStyle.vbdefaultbutton3.value.value:
            default_button = 3
        else:
            # default == VbMsgBoxStyle.vbdefaultbutton4:
            # are there ever 4 buttons? Is "Help" a fourth?
            # Help Button does not return. Another button must be pressed
            default_button = 4
        if default_button > len(button_list):
            default_button = 1
        print(f"{title}\n{icon}\n{prompt}\n{button_text}")
        user_supplied = input()
        tabs = user_supplied.count("\t")
        pos = (default_button + tabs - 1) % len(button_list)
        # What to do if we receive an option not presented?
        # F1 is help
        # Esc is cancel
        return button_list[pos]
