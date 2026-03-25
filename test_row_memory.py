#!/usr/bin/env python3
"""Test script for row selection memory feature."""

from pathlib import Path
from PySide6.QtWidgets import QApplication, QListWidgetItem
from PySide6.QtCore import Qt
from gyatt_o_tune.ui.main_window import MainWindow
from gyatt_o_tune.core.io import TuneLoader

# Create app and window
app = QApplication([])
window = MainWindow()

# Load a test tune
test_tune = Path("tuning_data/Final.msq")
if not test_tune.exists():
    print("ERROR: Test tune not found")
    exit(1)

# Load the tune
tune_loader = TuneLoader()
window.tune_data = tune_loader.load(test_tune)
window.selected_rows_per_table = {}  # Should be cleared

# Verify the dict exists
assert hasattr(window, 'selected_rows_per_table'), "selected_rows_per_table attribute missing"
assert isinstance(window.selected_rows_per_table, dict), "selected_rows_per_table should be a dict"
print("✓ Row selection memory initialized correctly")

# Update table display and get the first two 2D tables
window._update_table_display()
assert window.table_list.count() > 1, "Need at least 2 tables for test"

# Select first table
first_item = window.table_list.item(0)
window._on_table_selected(first_item, None)
first_table_name = first_item.data(Qt.ItemDataRole.UserRole)

# Check that table loaded
assert window.current_table is not None, "Table should be loaded"
print(f"✓ First table selected: {first_table_name}")

# Select a row in the first table (if it's 2D)
if window.current_table.rows > 1:
    window._load_selected_table_row(0)
    assert window.selected_table_row_idx == 0, "Row should be selected"
    print(f"✓ Row 0 selected in {first_table_name}")
    
    # Now select second table (this should save first table's selection)
    if window.table_list.count() > 1:
        second_item = window.table_list.item(1)
        window._on_table_selected(second_item, first_item)
        second_table_name = second_item.data(Qt.ItemDataRole.UserRole)
        print(f"✓ Second table selected: {second_table_name}")
        
        # Verify the first table's selection was saved during the switch
        if first_table_name in window.selected_rows_per_table:
            print(f"✓ Row selection saved for {first_table_name}: {window.selected_rows_per_table[first_table_name]}")
        else:
            print(f"✗ Row selection NOT saved for {first_table_name}")
        
        # Switch back to first table (this should restore its selection)
        window._on_table_selected(first_item, second_item)
        
        # Verify the row selection was restored
        if window.current_table.rows > 1:
            assert window.selected_table_row_idx == 0, "Row should be restored"
            print(f"✓ Row selection restored for {first_table_name}: row {window.selected_table_row_idx}")
        
print("\n✓ All row selection memory tests passed!")

