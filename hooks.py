# hooks.py — el vigilante del paso 9: escribir solo en entregas/.
#
# Un hook es código TUYO que corre en TU proceso, en un momento exacto del
# ciclo del agente. Corre ANTES que las reglas de permiso y que el modo.
#
# TRAMPA: las claves que devuelve un hook van en camelCase, TAMBIÉN en Python
# (hookSpecificOutput, permissionDecision, permissionDecisionReason).
#
# Evento PreToolUse, matcher "Write|Edit". Se registra así:
#   hooks={"PreToolUse": [HookMatcher(matcher="Write|Edit", hooks=[HOOK])]}

import pathlib
from typing import Any

CARPETA_PERMITIDA = "entregas"   # relativa al cwd del agente (./workspace)


async def solo_en_entregas(
    input_data: dict[str, Any], tool_use_id: str | None, context: Any
) -> dict[str, Any]:
    cwd = pathlib.Path(input_data.get("cwd") or ".").resolve()
    destino = pathlib.Path(input_data.get("tool_input", {}).get("file_path", ""))
    destino = (destino if destino.is_absolute() else cwd / destino).resolve()
    if destino.is_relative_to(cwd / CARPETA_PERMITIDA):
        return {}  # dentro de entregas/: que siga
    return {
        "systemMessage": f"Hook bloqueó una escritura fuera de {CARPETA_PERMITIDA}/: {destino}",
        "hookSpecificOutput": {
            "hookEventName": input_data["hook_event_name"],
            "permissionDecision": "deny",
            # Lo lee el MODELO: dile dónde sí puede escribir.
            "permissionDecisionReason": (
                f"Solo puedes escribir dentro de {CARPETA_PERMITIDA}/. "
                f"Guarda el archivo como {CARPETA_PERMITIDA}/{destino.name}. "
                "El xlsx original nunca se modifica."
            ),
        },
    }


HOOK = solo_en_entregas
