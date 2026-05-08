from vba_stdlib.interaction import Interaction


api = {
    "name": "vba",
    "modules": {
        "name": "interaction",
        "functions": {
            "msgbox": {
                "name": "msgbox",
                "handle": getattr(Interaction, "msgbox"),
            }
        }
    }
}
