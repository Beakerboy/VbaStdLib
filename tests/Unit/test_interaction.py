from vba_stdlib.interaction import Interaction
from vba_stdlib.enums import VbMsgBoxStyle, VbMsgBoxResult
from pytest_mock import MockerFixture

def test_msgbox(mocker: MockerFixture) -> None:
    mock_print = mocker.patch('builtins.print')
    mock_input = mocker.patch('builtins.input', return_value="")
    result = Interaction.msgbox("hello")
    assert result == VbMsgBoxResult.vbok
    expected = "Microsoft Excel\n\nhello\nOK"
    mock_print.assert_called_with(expected)

def test_msgbox2(mocker: MockerFixture) -> None:
    mock_print = mocker.patch('builtins.print')
    mock_input = mocker.patch('builtins.input', return_value="\t")
    result = Interaction.msgbox("hello", VbMsgBoxStyle.vbokcancel)
    assert result == VbMsgBoxResult.vbcancel
    expected = "Microsoft Excel\n\nhello\nOK    Cancel"
    mock_print.assert_called_with(expected)
