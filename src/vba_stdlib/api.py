from vba_stdlib.interaction import Interaction
from vba_stdlib.math import Math

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
                    "params": [{
                            "name": "prompt",
                            "optional": False,
                            "default": ""
                    }]
                }
            }
        },
        "math": {
            "name": "math",
            "type": "module",
            "functions": {
                "abs": {
                    "name": "abs",
                    "type": "function",
                    "project": "vba",
                    "module": "math",
                    "handle": getattr(Math, "abs"),
                    "params": [{
                            "name": "number",
                            "optional": False,
                            "default": ""
                    }]
                },
                "cos": {
                    "name": "cos",
                    "type": "function",
                    "project": "vba",
                    "module": "math",
                    "handle": getattr(Math, "cos"),
                    "params": [{
                            "name": "number",
                            "optional": False,
                            "default": ""
                    }]
                },
                "sqr": {
                    "name": "sqr",
                    "type": "function",
                    "project": "vba",
                    "module": "math",
                    "handle": getattr(Math, "sqr"),
                    "params": [{
                            "name": "number",
                            "optional": False,
                            "default": ""
                    }]
                },
                 "sin": {
                    "name": "sin",
                    "type": "function",
                    "project": "vba",
                    "module": "math",
                    "handle": getattr(Math, "sin"),
                    "params": [{
                            "name": "number",
                            "optional": False,
                            "default": ""
                    }]
                },
            }
        }
    }
}
