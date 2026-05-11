from vba_stdlib.interaction import Interaction


api = {
    "name": "vba",
    "type": "project",
    "modules": {
        "name": "interaction",
        "type": "module",
        "functions": {
            "msgbox": {
                "name": "msgbox",
                "type": "function",
                "handle": getattr(Interaction, "msgbox"),
            }
        }
    }
}
