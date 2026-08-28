"""Unit tests for `discover_write_intent_link_graph` (docs/recalc_and_link_refresh_plan.md
R3/R4): builds the graph `compute_link_commit_order()` consumes, from real workbook paths plus
`plan()`'s write-intent set. `scan_external_link_targets` is monkeypatched here — it's already
covered by its own real-Excel tests in test_link_discovery.py, and an absolute (R3/R4) link is
hard to reproduce portably in a fixture (Excel collapses same-drive links back to relative
form, per test_link_discovery.py's probe finding) — so this is the correct mock boundary
(project convention: mock only already-covered dependencies, never the thing under test)."""

from pathlib import Path

import pytest

from excel_runner.engine import discover_write_intent_link_graph


def _touch(path: Path) -> Path:
    path.write_text("placeholder")
    return path


class TestDiscoverWriteIntentLinkGraph:
    def test_no_links_gives_every_write_intent_workbook_an_empty_set(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr("excel_runner.engine.scan_external_link_targets", lambda path: [])
        paths = {
            "a": _touch(tmp_path / "a.xlsx"),
            "b": _touch(tmp_path / "b.xlsx"),
        }

        graph = discover_write_intent_link_graph(paths, write_intent={"a", "b"})

        assert graph == {"a": set(), "b": set()}

    def test_absolute_link_to_another_write_intent_workbook_becomes_an_edge(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        a_path = _touch(tmp_path / "a.xlsx")
        b_path = _touch(tmp_path / "b.xlsx")

        def fake_scan(path: Path) -> list[str]:
            return [str(b_path)] if path == a_path else []

        monkeypatch.setattr("excel_runner.engine.scan_external_link_targets", fake_scan)
        paths = {"a": a_path, "b": b_path}

        graph = discover_write_intent_link_graph(paths, write_intent={"a", "b"})

        assert graph == {"a": {"b"}, "b": set()}

    def test_absolute_link_to_a_read_only_workbook_is_r3_and_produces_no_edge(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """R3: target is a declared workbook, but it's not write-intent — leave untouched."""
        a_path = _touch(tmp_path / "a.xlsx")
        c_path = _touch(tmp_path / "c.xlsx")

        def fake_scan(path: Path) -> list[str]:
            return [str(c_path)] if path == a_path else []

        monkeypatch.setattr("excel_runner.engine.scan_external_link_targets", fake_scan)
        paths = {"a": a_path, "c": c_path}

        graph = discover_write_intent_link_graph(paths, write_intent={"a"})

        assert graph == {"a": set()}

    def test_same_folder_and_relative_subpath_links_are_ignored(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """R1/R2 are not R4 concerns for this graph, even if the bare filename happens to
        match another write-intent workbook's name."""
        a_path = _touch(tmp_path / "a.xlsx")
        b_path = _touch(tmp_path / "b.xlsx")

        def fake_scan(path: Path) -> list[str]:
            return ["b.xlsx", "sub/b.xlsx"] if path == a_path else []

        monkeypatch.setattr("excel_runner.engine.scan_external_link_targets", fake_scan)
        paths = {"a": a_path, "b": b_path}

        graph = discover_write_intent_link_graph(paths, write_intent={"a", "b"})

        assert graph == {"a": set(), "b": set()}

    def test_absolute_link_to_something_outside_this_run_is_ignored(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        a_path = _touch(tmp_path / "a.xlsx")
        outside_path = tmp_path / "elsewhere" / "unrelated.xlsx"

        def fake_scan(path: Path) -> list[str]:
            return [str(outside_path)] if path == a_path else []

        monkeypatch.setattr("excel_runner.engine.scan_external_link_targets", fake_scan)
        paths = {"a": a_path}

        graph = discover_write_intent_link_graph(paths, write_intent={"a"})

        assert graph == {"a": set()}

    def test_workbook_that_does_not_exist_yet_is_skipped_but_still_present_as_a_key(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """create_if_missing workbook — nothing to scan yet."""
        calls: list[Path] = []
        monkeypatch.setattr(
            "excel_runner.engine.scan_external_link_targets",
            lambda path: calls.append(path) or [],
        )
        paths = {"new": tmp_path / "new.xlsx"}

        graph = discover_write_intent_link_graph(paths, write_intent={"new"})

        assert graph == {"new": set()}
        assert calls == []
