"""In-memory token -> upload session. Same pattern as webapp/app.py's
_RESULTS dict: fine for a local, single-user, single-process tool; state
is lost on restart and isn't shared across worker processes."""
from __future__ import annotations

import uuid
from dataclasses import dataclass

from sync_studio.master import MasterResult
from sync_studio.parser import TabData


@dataclass
class Session:
    original_bytes: bytes
    original_filename: str
    baseline_tabs: list[TabData]
    master: MasterResult


_SESSIONS: dict[str, Session] = {}


def create_session(original_bytes: bytes, original_filename: str, baseline_tabs: list[TabData], master: MasterResult) -> str:
    token = uuid.uuid4().hex
    _SESSIONS[token] = Session(
        original_bytes=original_bytes,
        original_filename=original_filename,
        baseline_tabs=baseline_tabs,
        master=master,
    )
    return token


def get_session(token: str) -> Session | None:
    return _SESSIONS.get(token)
