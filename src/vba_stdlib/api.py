from vba_stdlib.interaction import Interaction


api = {
    "name": "vba",
    "type": "project",
    "modules": {
        "interaction": {
            "name": "interaction",
            "type": "module",
            "functions": {
                "msgbox": {
                    "name": "msgbox",
                    "type": "function",
                    "project": "vba",
                    "module": "interaction",
                    "handle": getattr(Interaction, "msgbox"),
                }
            }
        }
    }
}
