from vba_stdlib.interaction import Interaction
from pytest_mock import MockerFixture

def test_msgbox(mocker: MockerFixture) -> None:
    mock_print = mocker.patch('builtins.print')
    result = Interaction.msgbox("hello")
    assert int(result) == 1
    expected = "Microsoft Excel\n\n\hello\nOK    Cancel"
    mock_print.assert_called_with(expected)
