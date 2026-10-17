"""Unit tests for project_management module using pytest."""

from __future__ import annotations

import uuid

import pytest

import project_management as pm


def _make_app() -> pm.ProjectManagementApp:
    """Create a ProjectManagementApp instance without constructing Tk widgets."""
    app = pm.ProjectManagementApp.__new__(pm.ProjectManagementApp)
    app.data = {"projects": [], "tasks": []}
    app.selected_project_id = None
    return app


def test_safe_float_parses_valid_and_empty_values() -> None:
    """safe_float should parse valid text and treat empty as 0.0."""
    assert pm.safe_float("1,234.5", "Bid Price") == 1234.5
    assert pm.safe_float("", "Bid Price") == 0.0


def test_safe_float_rejects_invalid_text() -> None:
    """safe_float should raise ValueError for non-numeric input."""
    with pytest.raises(ValueError, match="Bid Price must be a valid number"):
        pm.safe_float("abc", "Bid Price")


def test_safe_int_parses_valid_and_empty_values() -> None:
    """safe_int should parse valid text and treat empty as 0."""
    assert pm.safe_int("42", "Duration") == 42
    assert pm.safe_int("", "Duration") == 0


def test_parse_date_rejects_wrong_format() -> None:
    """parse_date should reject non-YYYY-MM-DD input."""
    with pytest.raises(ValueError, match="must be in YYYY-MM-DD format"):
        pm.parse_date("10/17/2026", "Start Date")


def test_days_onsite_and_duration_calculations() -> None:
    """Date helpers should return expected day counts."""
    assert pm.calculate_days_onsite("2026-10-17", "2026-10-18") == 1
    assert pm.calculate_duration_days("2026-10-17", "2026-10-18") == 2


def test_suggest_next_job_number_uses_numeric_max() -> None:
    """Next job-number suggestion should ignore non-numeric values and increment max numeric."""
    app = _make_app()
    app.data["projects"] = [
        {"id": "a", "job_number": "ABC"},
        {"id": "b", "job_number": "1005"},
        {"id": "c", "job_number": "1002"},
    ]
    assert app._suggest_next_job_number() == "1006"


def test_suggest_next_job_number_defaults_when_no_numeric_jobs() -> None:
    """Next job-number suggestion should default to 1000 without numeric jobs."""
    app = _make_app()
    app.data["projects"] = [{"id": "a", "job_number": "X-1"}]
    assert app._suggest_next_job_number() == "1000"


def test_add_project_rejects_duplicate_job_number(monkeypatch: pytest.MonkeyPatch) -> None:
    """add_project should disallow duplicate job numbers."""
    app = _make_app()
    app.data["projects"] = [{"id": "existing", "job_number": "1001"}]

    payload = {
        "project_type": "House Demo",
        "project_address": "123 Main St",
        "job_number": "1001",
        "project_manager": "John Doe",
        "crew_name": "Team 1",
        "duration_days": 3,
        "bid_price": 1000.0,
        "job_cost": 800.0,
        "p_and_l": 200.0,
    }
    app._project_input = lambda: payload
    app.save_data = lambda: None
    app.refresh_all_views = lambda: None
    app.clear_project_form = lambda: None
    monkeypatch.setattr(pm.messagebox, "showinfo", lambda *args, **kwargs: None)

    with pytest.raises(ValueError, match="Job Number already exists"):
        app.add_project()


def test_update_project_rejects_duplicate_job_number_from_other_project() -> None:
    """update_project should reject changing to another project's job number."""
    app = _make_app()
    app.data["projects"] = [
        {"id": "1", "job_number": "1001"},
        {"id": "2", "job_number": "1002"},
    ]
    app.selected_project_id = "1"

    app._project_input = lambda: {
        "project_type": "Interior Demolition",
        "project_address": "Address",
        "job_number": "1002",
        "project_manager": "Jane Smith",
        "crew_name": "Team 2",
        "duration_days": 5,
        "bid_price": 2000.0,
        "job_cost": 1500.0,
        "p_and_l": 500.0,
    }

    with pytest.raises(ValueError, match="already exists for another project"):
        app.update_project()


def test_add_project_success_appends_project_with_uuid(monkeypatch: pytest.MonkeyPatch) -> None:
    """add_project should append a new project and generate an ID."""
    app = _make_app()
    payload = {
        "project_type": "Building Demo",
        "project_address": "500 Demo Ave",
        "job_number": "1003",
        "project_manager": "John Doe",
        "crew_name": "Team 1",
        "duration_days": 7,
        "bid_price": 5000.0,
        "job_cost": 4200.0,
        "p_and_l": 800.0,
    }
    app._project_input = lambda: payload.copy()
    app.save_data = lambda: None
    app.refresh_all_views = lambda: None
    app.clear_project_form = lambda: None
    monkeypatch.setattr(pm.messagebox, "showinfo", lambda *args, **kwargs: None)

    app.add_project()

    assert len(app.data["projects"]) == 1
    created = app.data["projects"][0]
    assert created["job_number"] == "1003"
    uuid.UUID(str(created["id"]))


def test_collect_report_rows_aggregates_to_one_row_per_project() -> None:
    """_collect_report_rows should aggregate task schedule into a single project row."""
    app = _make_app()
    app.data["projects"] = [
        {
            "id": "p1",
            "project_type": "House Demo",
            "project_address": "101 Main St",
            "job_number": "2001",
            "project_manager": "John Doe",
            "crew_name": "Team 1",
            "duration_days": 2,
            "bid_price": 3000.0,
            "job_cost": 2500.0,
            "p_and_l": 500.0,
        }
    ]
    app.data["tasks"] = [
        {
            "id": "t1",
            "task": "Task 1",
            "description": "",
            "job_number": "2001",
            "start_date": "2026-10-17",
            "end_date": "2026-10-17",
            "status": "In Progress",
        },
        {
            "id": "t2",
            "task": "Task 2",
            "description": "",
            "job_number": "2001",
            "start_date": "2026-10-18",
            "end_date": "2026-10-18",
            "status": "Todo",
        },
    ]

    rows = app._collect_report_rows(active_only=False, job_filter="", address_filter="")

    assert len(rows) == 1
    row = rows[0]
    assert row["job_number"] == "2001"
    assert row["start_date"] == "2026-10-17"
    assert row["end_date"] == "2026-10-18"
    assert row["days_onsite"] == 2
    assert row["status"] == "In Progress"


def test_collect_report_rows_active_only_excludes_done_projects() -> None:
    """Active-only report should exclude projects whose aggregate status is Done."""
    app = _make_app()
    app.data["projects"] = [
        {
            "id": "p1",
            "project_type": "House Demo",
            "project_address": "101 Main St",
            "job_number": "2001",
            "project_manager": "John Doe",
            "crew_name": "Team 1",
            "duration_days": 1,
            "bid_price": 3000.0,
            "job_cost": 2500.0,
            "p_and_l": 500.0,
        }
    ]
    app.data["tasks"] = [
        {
            "id": "t1",
            "task": "Task 1",
            "description": "",
            "job_number": "2001",
            "start_date": "2026-10-17",
            "end_date": "2026-10-17",
            "status": "Done",
        }
    ]

    rows = app._collect_report_rows(active_only=True, job_filter="", address_filter="")
    assert rows == []
