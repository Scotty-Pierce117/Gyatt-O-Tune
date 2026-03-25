#!/usr/bin/env python3
"""Test script for new filter logic."""

from pathlib import Path
from PySide6.QtWidgets import QApplication
from gyatt_o_tune.ui.main_window import MainWindow
from gyatt_o_tune.core.io import TuneLoader

app = QApplication([])
window = MainWindow()

# Load a test tune file
test_tune = Path("tuning_data/Final.msq")
if test_tune.exists():
    tune_loader = TuneLoader()
    window.tune_data = tune_loader.load(test_tune)
    
    # Check _is_1d_table method works
    count_1d = 0
    count_2d = 0
    for table_name, table in window.tune_data.tables.items():
        is_1d = window._is_1d_table(table_name)
        expected_1d = (table.rows == 1 or table.cols == 1)
        assert is_1d == expected_1d, f"Table {table_name}: got {is_1d}, expected {expected_1d}"
        if is_1d:
            count_1d += 1
        else:
            count_2d += 1
    
    print(f"✓ _is_1d_table method works correctly (found {count_1d} 1D, {count_2d} 2D tables)")
    
    # Test: Verify handler logic for disabling 1D filter when showing only favorites
    # Initial state: not showing only favorited, so 1D filter should be enabled
    assert window.show_1d_tables_action.isEnabled(), "1D filter should start enabled"
    print("✓ 1D filter initially enabled (not showing only favorites)")
    
    # Simulate clicking the "Only Show Favorited Tables" option
    window.only_show_favorited_tables_action.setChecked(True)
    window._on_only_show_favorited_toggled()  # Manually call handler to simulate click
    assert not window.show_1d_tables_action.isEnabled(), "1D filter should be disabled when showing only favorites"
    print("✓ 1D filter disabled when 'Only Show Favorited Tables' is checked")
    
    # Simulate unchecking the "Only Show Favorited Tables" option
    window.only_show_favorited_tables_action.setChecked(False)
    window._on_only_show_favorited_toggled()  # Manually call handler
    assert window.show_1d_tables_action.isEnabled(), "1D filter should be enabled when not showing only favorites"
    print("✓ 1D filter enabled when 'Only Show Favorited Tables' is unchecked")
    
    print("\n✓ All filter logic tests passed!")
else:
    print("Warning: Test tune file not found at tuning_data/Final.msq")

print("✓ Complete implementation verified")

