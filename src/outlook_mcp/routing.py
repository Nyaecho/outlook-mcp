"""Account routing: which account serves a given tool call.

Multi-account configs route *capabilities* (mail / calendar / contacts / todo)
to accounts, so one Microsoft identity per concern — mail on the account that
receives notifications, To Do on the one that holds the task lists. The
classification reuses ``toolsets.TOOL_GROUPS`` (already drift-guarded) and maps
each group onto a capability:

- mail-centric groups (mail, drafts, attachments, folders, admin, digest) fold
  into ``mail`` — they all act on the mailbox.
- calendar / contacts / todo map one-to-one.
- ``account`` tools (whoami, auth_status, …) follow the *active* account, not
  a capability: ``capability_for`` returns None for them.
- ``delta`` tools span capabilities and split by name.
"""

from __future__ import annotations

from outlook_mcp.toolsets import TOOL_GROUPS

_GROUP_CAPABILITY: dict[str, str] = {
    "mail": "mail",
    "drafts": "mail",
    "attachments": "mail",
    "folders": "mail",
    # changes_since composes the mail/calendar/contacts deltas; routed with
    # mail, the dominant capability — split setups that route calendar away
    # from mail should call the per-capability delta tools instead.
    "digest": "mail",
    "admin": "mail",
    "calendar": "calendar",
    "contacts": "contacts",
    "todo": "todo",
}

_TOOL_CAPABILITY_OVERRIDES: dict[str, str] = {
    "outlook_list_inbox_delta": "mail",
    "outlook_list_events_delta": "calendar",
    "outlook_list_contacts_delta": "contacts",
}


def capability_for(tool_name: str | None) -> str | None:
    """Routing capability for a tool, or None when it follows the active account."""
    if tool_name in _TOOL_CAPABILITY_OVERRIDES:
        return _TOOL_CAPABILITY_OVERRIDES[tool_name]
    group = TOOL_GROUPS.get(tool_name or "")
    if group is None or group == "account":
        return None
    return _GROUP_CAPABILITY.get(group)
