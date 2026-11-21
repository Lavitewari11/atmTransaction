import pytest
import os
import sys

def test_tc_t_10_app_files_exist():
    """TC-T-10: Verify app.py and all pages exist."""
    pages = [
        "app.py",
        "src/ui.py",
        "pages/01_Fleet_Overview.py",
        "pages/02_Failure_Category.py",
        "pages/03_Terminal_Health.py",
        "pages/04_Temporal_Patterns.py",
        "pages/05_Network_Latency.py",
        "pages/06_Financial_Impact.py",
    ]
    for p in pages:
        assert os.path.exists(p), f"Missing file: {p}"

def test_tc_t_11_ui_theme_config():
    """TC-T-11: Verify .streamlit/config.toml defines valid primary color."""
    config_path = ".streamlit/config.toml"
    assert os.path.exists(config_path)
    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()
    assert 'primaryColor = "#2B6CB0"' in content
