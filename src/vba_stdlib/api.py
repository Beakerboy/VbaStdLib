from vba_stdlib.interaction import Interaction


api = {
    "name": "vba",
    "modules": {
        "name": "interaction",
        "functions": {
            "msgbox": {
                "name": "msgbox",
                "type": "function",
                "handle": getattr(Interaction, "msgbox"),
            }
        }
    }
}
