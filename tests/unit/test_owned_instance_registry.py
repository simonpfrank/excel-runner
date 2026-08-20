"""Unit tests for OwnedInstanceRegistry (Spec sec 3.1, PRD sec 6.2.1).

Real xlwings against a real, locally-spawned Excel instance — no mocks (project convention).
Slower than the file-backend tests (spawning/quitting Excel takes real wall-clock time) but
that's the honest cost of not faking COM automation; skipped outright where Excel isn't
available (tests/unit/conftest.py).
"""

import os
import time

import pytest

from tests.unit.conftest import requires_excel

_TERMINATION_GRACE_SECONDS = 5.0


def _process_alive(pid: int) -> bool:
    """Check whether `pid` is still running.

    `os.kill(pid, 0)` is a POSIX idiom that doesn't work on Windows for arbitrary PIDs
    (raises `OSError: [WinError 87]` instead of a clean liveness signal) — found empirically
    running this suite on Windows for the first time. Uses `ctypes`/`OpenProcess` there
    instead, stdlib only, no new dependency for a two-line test helper.
    """
    if os.name == "nt":
        import ctypes

        handle = ctypes.windll.kernel32.OpenProcess(
            0x1000, False, pid
        )  # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    return True


def _wait_until_dead(pid: int, timeout: float = _TERMINATION_GRACE_SECONDS) -> bool:
    """Poll for a process to actually exit.

    Quitting Excel on macOS is asynchronous — `app.quit()` returns before the underlying
    process has actually terminated (confirmed empirically: ~0.5s observed locally). An
    immediate `os.kill(pid, 0)` right after quit() is a race, not a real check.
    """
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _process_alive(pid):
            return True
        time.sleep(0.1)
    return not _process_alive(pid)


@requires_excel
class TestSpawn:
    def test_spawn_returns_a_new_app_and_tracks_its_pid(self) -> None:
        from excel_runner.backends import OwnedInstanceRegistry

        registry = OwnedInstanceRegistry()
        try:
            app = registry.spawn()
            assert app.pid in registry.pids
        finally:
            registry.close_owned()

    def test_spawn_never_reuses_an_existing_instance(self) -> None:
        from excel_runner.backends import OwnedInstanceRegistry

        registry = OwnedInstanceRegistry()
        try:
            first = registry.spawn()
            second = registry.spawn()
            assert first.pid != second.pid
            assert set(registry.pids) == {first.pid, second.pid}
        finally:
            registry.close_owned()


@requires_excel
class TestCloseOwned:
    def test_quits_every_owned_instance_and_clears_tracking(self) -> None:
        from excel_runner.backends import OwnedInstanceRegistry

        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        pid = app.pid

        registry.close_owned()

        assert registry.pids == ()
        assert _wait_until_dead(pid)

    def test_close_owned_with_nothing_spawned_does_not_raise(self) -> None:
        from excel_runner.backends import OwnedInstanceRegistry

        registry = OwnedInstanceRegistry()
        registry.close_owned()  # should not raise

    def test_one_failing_close_does_not_prevent_others_from_closing(self) -> None:
        """Crash-safety requirement (PRD sec 6.3, mirrors SessionManager.close_all()): every
        owned instance must get a close attempt, even if an earlier one fails.

        Uses a fake exploding stand-in, not a real double-quit — found empirically that a
        genuinely-already-dead App's quit() is *not* a reliable way to trigger this on macOS:
        it raises "-600 Application isn't running" only when it's the only Excel process on
        the machine; with another owned instance still alive, the same call silently no-ops
        instead (confirmed it doesn't cross-target and kill the other instance either — just
        an inert no-op, not a safety hole). Not something this registry can rely on to fail
        predictably, same reasoning as SessionManager's own test.
        """
        from excel_runner.backends import OwnedInstanceRegistry

        class _ExplodingApp:
            def quit(self) -> None:
                raise RuntimeError("simulated quit failure")

        registry = OwnedInstanceRegistry()
        # The exploding entry is inserted first, so iteration reaches it before the real one —
        # this is what actually proves close_owned() doesn't stop after the first failure.
        registry._owned[-1] = _ExplodingApp()  # type: ignore[assignment]
        still_alive = registry.spawn()

        with pytest.raises(ExceptionGroup):
            registry.close_owned()

        assert _wait_until_dead(
            still_alive.pid
        )  # proves it still got a real close attempt
