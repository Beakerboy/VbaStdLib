from typing import Any
from vba_stdlib.enums import VbMsgBoxResult, VbMsgBoxStyle
from vba_types import VBAString, VBALong


class Interaction:

    @staticmethod
    def msgbox(prompt: Any,
               buttons: VbMsgBoxStyle = VbMsgBoxStyle.vbokonly,
               title: VBAString | None = None,
               help_file: VBAString | None = None,
               context: VBALong | None = None) -> VbMsgBoxResult:
        if title is None:
            # title = current_project_name
            pass
        if help_file is not None and context is None:
            # raise some sort of error
            pass
        number_type = buttons & 7
        
        if number_type == VbMsgBoxStyle.vbokonly:
            button_text = "OK"
        elif number_type == VbMsgBoxStyle.vbokcancel:
            button_text = "OK    Cancel"
        elif number_type == VbMsgBoxStyle.vbabortretryignore:
            button_text = "Abort    Retry    Ignore"
        elif number_type == VbMsgBoxStyle.vbyesnocancel:
            button_text = "Yes    No    Cancel"
        elif number_type == VbMsgBoxStyle.vbyesno:
            button_text = "Yes    No"
        elif number_type == VbMsgBoxStyle.vbretrycancel:
            button_text = "Retry    Cancel"
        else:
            # what kind of exception should this be?
            raise Exception("unknown button type")

        icon = button & 112
        if icon == VbMsgBoxStyle.vbcritical:
            window_icon = "X"
        elif icon == VbMsgBoxStyle.vbquestion:
            window_icon = "?"
        elif icon == VbMsgBoxStyle.vbexclamation:
            window_icon = "!"
        elif icon == VbMsgBoxStyle.vbinformation:
            window_icon = "i"
        else:
            # what kind of exception should this be?
            raise Exception("unknown icon type")

        default = button & 768
        if default == VbMsgBoxStyle.vbdefaultbutton1:
            default_button = 1
        elif default == VbMsgBoxStyle.vbdefaultbutton2:
            default_button = 2
        elif default == VbMsgBoxStyle.vbdefaultbutton3:
            default_button = 3
        else:
            # default == VbMsgBoxStyle.vbdefaultbutton4:
            # are there ever 4 buttons? Is "Help" a fourth?
            # Help Button does not return. Another button must be pressed
            default_button = 4
        print(str(prompt))
        # input = input()
        # What to do if we receive an option not presented?
        # F1 is help
        # Esc is cancel
        return VbMsgBoxResult.vbok
