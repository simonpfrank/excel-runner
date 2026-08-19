"""Run-preparation and run-state layer: action discovery for now. See Spec sec 5.

Session management, the scratch-copy execution model, and both validation tiers land here in
later increments (Spec sec 5.2/5.3/5.4) — this module currently covers sec 5.1 (registry).
"""

import inspect
from collections.abc import Callable
from dataclasses import dataclass
from types import ModuleType
from typing import Any, Literal

from excel_runner.core import ACTION_CAPABILITIES, ActionResult


@dataclass(frozen=True)
class ActionSpec:
    """A discovered action: its name, callable, capability, and parameter schema.

    Args:
        name: The action's name, matching an `action:` field in a workflow step.
        fn: The action function itself.
        capability: Which backend this action needs. "depends_on_param" is a named, single
            exception (PRD sec 7's `read_metadata`) — not a general mechanism.
        param_schema: `{"properties": {name: {"type": ...}}, "required": [...]}`, derived from
            `fn`'s signature (excluding `session`).
    """

    name: str
    fn: Callable[..., ActionResult]
    capability: Literal["file", "com", "depends_on_param"]
    param_schema: dict[str, Any]


def _generate_param_schema(fn: Callable[..., ActionResult]) -> dict[str, Any]:
    """Build a param schema from an action function's signature, excluding `session`."""
    signature = inspect.signature(fn)
    properties: dict[str, Any] = {}
    required: list[str] = []
    for param_name, param in signature.parameters.items():
        if param_name == "session":
            continue
        properties[param_name] = {"type": param.annotation}
        if param.default is inspect.Parameter.empty:
            required.append(param_name)
    return {"properties": properties, "required": required}


def discover_actions(module: ModuleType) -> dict[str, ActionSpec]:
    """Scan a module for capability-tagged action functions and build the registry.

    Args:
        module: The module to scan — normally `excel_runner.actions`.

    Returns:
        Mapping of action name to `ActionSpec`.
    """
    registry: dict[str, ActionSpec] = {}
    for name, fn in inspect.getmembers(module, inspect.isfunction):
        capability = ACTION_CAPABILITIES.get(name)
        if capability is None:
            continue
        registry[name] = ActionSpec(
            name=name,
            fn=fn,
            capability=capability,
            param_schema=_generate_param_schema(fn),
        )
    return registry
