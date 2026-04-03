import json
import sys
from pathlib import Path
from typing import Any

from PyQt6 import QtCore, QtWidgets
from PyQt6.QtGui import QColor

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = BASE_DIR / "word.ribbon.json"


class RibbonConfigEditor(QtWidgets.QMainWindow):
    ALLOWED_COMMAND_KINDS = {"button", "toggle", "combo", "fontCombo", "menuButton"}
    ALLOWED_ITEM_SIZES = {"small", "medium", "large"}

    def __init__(self, config_path: Path) -> None:
        super().__init__()
        self.config_path = config_path
        self.config: dict[str, Any] = {}
        self.validation_errors: list[dict[str, Any]] = []

        self.setWindowTitle("Ribbon JSON Editor")
        self.resize(1100, 700)

        self._build_ui()
        self.load_config()

    def _build_ui(self) -> None:
        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        root_layout = QtWidgets.QVBoxLayout(central)

        toolbar_layout = QtWidgets.QHBoxLayout()
        self.path_label = QtWidgets.QLabel(str(self.config_path))
        self.path_label.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
        toolbar_layout.addWidget(self.path_label, 1)

        self.reload_button = QtWidgets.QPushButton("Reload")
        self.reload_button.clicked.connect(self.load_config)
        toolbar_layout.addWidget(self.reload_button)

        self.save_button = QtWidgets.QPushButton("Save")
        self.save_button.clicked.connect(self.save_config)
        toolbar_layout.addWidget(self.save_button)

        self.validate_button = QtWidgets.QPushButton("Validate")
        self.validate_button.clicked.connect(self.validate_and_report)
        toolbar_layout.addWidget(self.validate_button)

        root_layout.addLayout(toolbar_layout)

        self.tabs = QtWidgets.QTabWidget(self)
        root_layout.addWidget(self.tabs, 1)

        self._build_quick_access_tab()
        self._build_structure_tab()
        self._build_errors_tab()

        self.statusBar().showMessage("Ready", 1500)

    def _build_quick_access_tab(self) -> None:
        page = QtWidgets.QWidget(self)
        self.quick_access_page = page
        layout = QtWidgets.QVBoxLayout(page)

        buttons_layout = QtWidgets.QHBoxLayout()
        add_button = QtWidgets.QPushButton("Add")
        add_button.clicked.connect(self.add_quick_access_row)
        buttons_layout.addWidget(add_button)

        remove_button = QtWidgets.QPushButton("Remove Selected")
        remove_button.clicked.connect(self.remove_selected_quick_access_rows)
        buttons_layout.addWidget(remove_button)
        buttons_layout.addStretch(1)

        layout.addLayout(buttons_layout)

        self.quick_access_table = QtWidgets.QTableWidget(0, 2, self)
        self.quick_access_table.setHorizontalHeaderLabels(["commandRef", "menuRef"])
        self.quick_access_table.horizontalHeader().setStretchLastSection(True)
        self.quick_access_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.quick_access_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
            | QtWidgets.QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        layout.addWidget(self.quick_access_table, 1)

        self.tabs.addTab(page, "Quick Access")

    def _build_structure_tab(self) -> None:
        page = QtWidgets.QWidget(self)
        self.structure_page = page
        layout = QtWidgets.QVBoxLayout(page)

        buttons_layout = QtWidgets.QHBoxLayout()
        add_tab_button = QtWidgets.QPushButton("Add Tab")
        add_tab_button.clicked.connect(self.add_tab_node)
        buttons_layout.addWidget(add_tab_button)

        add_group_button = QtWidgets.QPushButton("Add Group")
        add_group_button.clicked.connect(self.add_group_node)
        buttons_layout.addWidget(add_group_button)

        add_item_button = QtWidgets.QPushButton("Add Item")
        add_item_button.clicked.connect(self.add_item_node)
        buttons_layout.addWidget(add_item_button)

        remove_button = QtWidgets.QPushButton("Remove Selected")
        remove_button.clicked.connect(self.remove_selected_tree_node)
        buttons_layout.addWidget(remove_button)
        buttons_layout.addStretch(1)

        layout.addLayout(buttons_layout)

        self.structure_tree = QtWidgets.QTreeWidget(self)
        self.structure_tree.setColumnCount(8)
        self.structure_tree.setHeaderLabels(
            [
                "nodeType",
                "id",
                "labelKey",
                "order/priority",
                "commandRef",
                "size",
                "menuRef",
                "showText",
            ]
        )
        self.structure_tree.setAlternatingRowColors(True)
        self.structure_tree.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
            | QtWidgets.QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self.structure_tree.header().setStretchLastSection(False)
        self.structure_tree.header().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.structure_tree, 1)

        self.tabs.addTab(page, "Tabs/Groups/Items")

    def _build_errors_tab(self) -> None:
        page = QtWidgets.QWidget(self)
        self.errors_page = page
        layout = QtWidgets.QVBoxLayout(page)

        help_label = QtWidgets.QLabel("Double-click an error to jump to the related field.")
        layout.addWidget(help_label)

        self.errors_list = QtWidgets.QListWidget(self)
        self.errors_list.itemDoubleClicked.connect(self._open_selected_error)
        layout.addWidget(self.errors_list, 1)

        open_button = QtWidgets.QPushButton("Go To Selected")
        open_button.clicked.connect(self._open_selected_error)
        layout.addWidget(open_button)

        self.tabs.addTab(page, "Errors")

    def _set_validation_errors(self, errors: list[dict[str, Any]]) -> None:
        self.validation_errors = errors
        self.errors_list.clear()

        for index, error in enumerate(errors):
            item = QtWidgets.QListWidgetItem(error.get("message", "Unknown error"))
            item.setData(QtCore.Qt.ItemDataRole.UserRole, index)
            self.errors_list.addItem(item)

    def _open_selected_error(self, item: QtWidgets.QListWidgetItem | None = None) -> None:
        selected_item = item
        if selected_item is None:
            selected_item = self.errors_list.currentItem()
        if selected_item is None:
            return

        index = selected_item.data(QtCore.Qt.ItemDataRole.UserRole)
        if index is None or not (0 <= index < len(self.validation_errors)):
            return

        error = self.validation_errors[index]
        target = error.get("target", "none")

        if target == "quick":
            row = int(error.get("row", 0))
            column = int(error.get("column", 0))
            self.tabs.setCurrentWidget(self.quick_access_page)
            self.quick_access_table.setCurrentCell(row, column)
            table_item = self.quick_access_table.item(row, column)
            if table_item is not None:
                self.quick_access_table.scrollToItem(table_item)
            return

        if target == "tree":
            tab_index = int(error.get("tab_index", -1))
            group_index = int(error.get("group_index", -1))
            item_index = int(error.get("item_index", -1))
            column = int(error.get("column", 0))

            if tab_index < 0 or tab_index >= self.structure_tree.topLevelItemCount():
                return

            tab_item = self.structure_tree.topLevelItem(tab_index)
            current_item = tab_item

            if group_index >= 0 and group_index < tab_item.childCount():
                current_item = tab_item.child(group_index)

            if item_index >= 0 and group_index >= 0 and group_index < tab_item.childCount():
                group_item = tab_item.child(group_index)
                if item_index < group_item.childCount():
                    current_item = group_item.child(item_index)

            self.tabs.setCurrentWidget(self.structure_page)
            self.structure_tree.setCurrentItem(current_item, column)
            self.structure_tree.scrollToItem(current_item)
            return

        self.tabs.setCurrentWidget(self.errors_page)

    def _iter_tree_items(self):
        for i in range(self.structure_tree.topLevelItemCount()):
            root = self.structure_tree.topLevelItem(i)
            yield root
            yield from self._iter_child_items(root)

    def _iter_child_items(self, parent: QtWidgets.QTreeWidgetItem):
        for index in range(parent.childCount()):
            child = parent.child(index)
            yield child
            yield from self._iter_child_items(child)

    def _clear_validation_marks(self) -> None:
        for row in range(self.quick_access_table.rowCount()):
            for column in range(self.quick_access_table.columnCount()):
                item = self.quick_access_table.item(row, column)
                if item is not None:
                    item.setBackground(QColor("transparent"))

        for tree_item in self._iter_tree_items():
            for column in (4, 6):
                tree_item.setBackground(column, QColor("transparent"))

    def _mark_quick_access_cell(self, row: int, column: int) -> None:
        item = self.quick_access_table.item(row, column)
        if item is None:
            item = QtWidgets.QTableWidgetItem("")
            self.quick_access_table.setItem(row, column, item)
        item.setBackground(QColor("#ffd6d6"))

    def _collect_known_ids(self) -> tuple[set[str], set[str]]:
        command_ids = {
            command.get("id", "")
            for command in self.config.get("commands", [])
            if isinstance(command, dict) and command.get("id")
        }
        menu_ids = {
            menu.get("id", "")
            for menu in self.config.get("menus", [])
            if isinstance(menu, dict) and menu.get("id")
        }
        return command_ids, menu_ids

    def validate_ui_references(self) -> list[dict[str, Any]]:
        self._clear_validation_marks()
        errors: list[dict[str, Any]] = []

        command_ids, menu_ids = self._collect_known_ids()

        for row in range(self.quick_access_table.rowCount()):
            command_item = self.quick_access_table.item(row, 0)
            menu_item = self.quick_access_table.item(row, 1)

            command_ref = command_item.text().strip() if command_item else ""
            menu_ref = menu_item.text().strip() if menu_item else ""

            if command_ref and command_ref not in command_ids:
                self._mark_quick_access_cell(row, 0)
                errors.append(
                    {
                        "message": f"QuickAccess row {row + 1}: unknown commandRef '{command_ref}'",
                        "target": "quick",
                        "row": row,
                        "column": 0,
                    }
                )

            if menu_ref and menu_ref not in menu_ids:
                self._mark_quick_access_cell(row, 1)
                errors.append(
                    {
                        "message": f"QuickAccess row {row + 1}: unknown menuRef '{menu_ref}'",
                        "target": "quick",
                        "row": row,
                        "column": 1,
                    }
                )

        for i in range(self.structure_tree.topLevelItemCount()):
            tab_item = self.structure_tree.topLevelItem(i)
            if tab_item.text(0) != "tab":
                continue

            tab_name = tab_item.text(1).strip() or f"tab[{i + 1}]"
            for j in range(tab_item.childCount()):
                group_item = tab_item.child(j)
                if group_item.text(0) != "group":
                    continue

                group_name = group_item.text(1).strip() or f"group[{j + 1}]"
                for k in range(group_item.childCount()):
                    item_node = group_item.child(k)
                    if item_node.text(0) != "item":
                        continue

                    command_ref = item_node.text(4).strip()
                    menu_ref = item_node.text(6).strip()
                    position = f"{tab_name}/{group_name}/item[{k + 1}]"

                    if command_ref and command_ref not in command_ids:
                        item_node.setBackground(4, QColor("#ffd6d6"))
                        errors.append(
                            {
                                "message": f"{position}: unknown commandRef '{command_ref}'",
                                "target": "tree",
                                "tab_index": i,
                                "group_index": j,
                                "item_index": k,
                                "column": 4,
                            }
                        )

                    if menu_ref and menu_ref not in menu_ids:
                        item_node.setBackground(6, QColor("#ffd6d6"))
                        errors.append(
                            {
                                "message": f"{position}: unknown menuRef '{menu_ref}'",
                                "target": "tree",
                                "tab_index": i,
                                "group_index": j,
                                "item_index": k,
                                "column": 6,
                            }
                        )

        return errors

    def validate_config_structure(self) -> list[str]:
        errors: list[str] = []

        commands = self.config.get("commands", [])
        menus = self.config.get("menus", [])

        command_id_occurrences: dict[str, int] = {}
        for index, command in enumerate(commands, start=1):
            if not isinstance(command, dict):
                errors.append(f"commands[{index}] must be an object")
                continue

            command_id = str(command.get("id", "")).strip()
            kind = str(command.get("kind", "")).strip()
            execute = str(command.get("execute", "")).strip()

            if not command_id:
                errors.append(f"commands[{index}] missing required field 'id'")
            else:
                command_id_occurrences[command_id] = command_id_occurrences.get(command_id, 0) + 1

            if not kind:
                errors.append(f"commands[{index}] ({command_id or '?'}) missing required field 'kind'")
            elif kind not in self.ALLOWED_COMMAND_KINDS:
                allowed = ", ".join(sorted(self.ALLOWED_COMMAND_KINDS))
                errors.append(f"commands[{index}] ({command_id or '?'}) has invalid kind '{kind}', allowed: {allowed}")

            if not execute:
                errors.append(f"commands[{index}] ({command_id or '?'}) missing required field 'execute'")

        duplicate_commands = sorted([item for item, count in command_id_occurrences.items() if count > 1])
        for duplicate_id in duplicate_commands:
            errors.append(f"duplicate command id '{duplicate_id}'")

        menu_id_occurrences: dict[str, int] = {}
        for index, menu in enumerate(menus, start=1):
            if not isinstance(menu, dict):
                errors.append(f"menus[{index}] must be an object")
                continue

            menu_id = str(menu.get("id", "")).strip()
            if not menu_id:
                errors.append(f"menus[{index}] missing required field 'id'")
            else:
                menu_id_occurrences[menu_id] = menu_id_occurrences.get(menu_id, 0) + 1

        duplicate_menus = sorted([item for item, count in menu_id_occurrences.items() if count > 1])
        for duplicate_id in duplicate_menus:
            errors.append(f"duplicate menu id '{duplicate_id}'")

        ui_tabs = self._read_tabs()
        tab_id_occurrences: dict[str, int] = {}
        group_id_occurrences: dict[str, int] = {}
        for tab in ui_tabs:
            tab_id = tab.get("id", "")
            if not tab_id:
                errors.append("ui.tabs contains tab without id")
            else:
                tab_id_occurrences[tab_id] = tab_id_occurrences.get(tab_id, 0) + 1

            groups = tab.get("groups", [])
            if not groups:
                errors.append(f"tab '{tab_id or '?'}' has no groups")
            for group in groups:
                group_id = group.get("id", "")
                if not group_id:
                    errors.append(f"tab '{tab_id or '?'}' contains group without id")
                else:
                    group_id_occurrences[group_id] = group_id_occurrences.get(group_id, 0) + 1

                items = group.get("items", [])
                if not items:
                    errors.append(f"group '{group_id or '?'}' in tab '{tab_id or '?'}' has no items")

                for item_index, item in enumerate(items, start=1):
                    command_ref = str(item.get("commandRef", "")).strip()
                    if not command_ref:
                        errors.append(
                            f"group '{group_id or '?'}' in tab '{tab_id or '?'}' has item[{item_index}] without commandRef"
                        )

                    size = str(item.get("size", "")).strip()
                    if size and size not in self.ALLOWED_ITEM_SIZES:
                        allowed_sizes = ", ".join(sorted(self.ALLOWED_ITEM_SIZES))
                        errors.append(
                            f"group '{group_id or '?'}' item[{item_index}] has invalid size '{size}', allowed: {allowed_sizes}"
                        )

        duplicate_tabs = sorted([item for item, count in tab_id_occurrences.items() if count > 1])
        for duplicate_id in duplicate_tabs:
            errors.append(f"duplicate tab id '{duplicate_id}'")

        duplicate_groups = sorted([item for item, count in group_id_occurrences.items() if count > 1])
        for duplicate_id in duplicate_groups:
            errors.append(f"duplicate group id '{duplicate_id}'")

        return errors

    def validate_and_report(self) -> bool:
        structure_errors = [
            {
                "message": message,
                "target": "none",
            }
            for message in self.validate_config_structure()
        ]
        reference_errors = self.validate_ui_references()
        errors = structure_errors + reference_errors
        self._set_validation_errors(errors)

        if not errors:
            self.statusBar().showMessage("Validation passed", 2500)
            QtWidgets.QMessageBox.information(self, "Validation", "No errors found.")
            return True

        preview = "\n".join(error["message"] for error in errors[:15])
        suffix = ""
        if len(errors) > 15:
            suffix = f"\n... and {len(errors) - 15} more"

        QtWidgets.QMessageBox.warning(
            self,
            "Validation errors",
            f"Validation failed:\n\n{preview}{suffix}",
        )
        self.tabs.setCurrentWidget(self.errors_page)
        self.statusBar().showMessage("Validation failed", 4000)
        return False

    def load_config(self) -> None:
        try:
            with self.config_path.open("r", encoding="utf-8") as config_file:
                self.config = json.load(config_file)
        except Exception as error:
            QtWidgets.QMessageBox.critical(self, "Load Error", f"Cannot load config file:\n{error}")
            return

        self.populate_quick_access_table()
        self.populate_structure_tree()
        self._set_validation_errors([])
        self.statusBar().showMessage("Config loaded", 2000)

    def populate_quick_access_table(self) -> None:
        quick_access = self.config.get("ui", {}).get("quickAccess", [])
        self.quick_access_table.setRowCount(0)

        for item in quick_access:
            row = self.quick_access_table.rowCount()
            self.quick_access_table.insertRow(row)
            self.quick_access_table.setItem(row, 0, QtWidgets.QTableWidgetItem(item.get("commandRef", "")))
            self.quick_access_table.setItem(row, 1, QtWidgets.QTableWidgetItem(item.get("menuRef", "")))

    def populate_structure_tree(self) -> None:
        self.structure_tree.clear()

        tabs = self.config.get("ui", {}).get("tabs", [])
        for tab in tabs:
            tab_item = self._make_tree_item(
                node_type="tab",
                node_id=tab.get("id", ""),
                label_key=tab.get("labelKey", ""),
                order_or_priority=str(tab.get("order", "")),
            )
            self.structure_tree.addTopLevelItem(tab_item)

            for group in tab.get("groups", []):
                group_item = self._make_tree_item(
                    node_type="group",
                    node_id=group.get("id", ""),
                    label_key=group.get("labelKey", ""),
                    order_or_priority=str(group.get("priority", "")),
                )
                tab_item.addChild(group_item)

                for ribbon_item in group.get("items", []):
                    item_node = self._make_tree_item(
                        node_type="item",
                        node_id="",
                        label_key="",
                        order_or_priority="",
                        command_ref=ribbon_item.get("commandRef", ""),
                        size=ribbon_item.get("size", ""),
                        menu_ref=ribbon_item.get("menuRef", ""),
                        show_text=str(ribbon_item.get("showText", "")),
                    )
                    group_item.addChild(item_node)

        self.structure_tree.expandAll()

    def _make_tree_item(
        self,
        node_type: str,
        node_id: str,
        label_key: str,
        order_or_priority: str,
        command_ref: str = "",
        size: str = "",
        menu_ref: str = "",
        show_text: str = "",
    ) -> QtWidgets.QTreeWidgetItem:
        item = QtWidgets.QTreeWidgetItem(
            [node_type, node_id, label_key, order_or_priority, command_ref, size, menu_ref, show_text]
        )
        item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
        item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsSelectable)
        item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsEnabled)

        # node type stays immutable to avoid malformed hierarchy edits.
        item.setFlags(item.flags() & ~QtCore.Qt.ItemFlag.ItemNeverHasChildren)
        return item

    def add_quick_access_row(self) -> None:
        row = self.quick_access_table.rowCount()
        self.quick_access_table.insertRow(row)
        self.quick_access_table.setItem(row, 0, QtWidgets.QTableWidgetItem("cmd.newCommand"))
        self.quick_access_table.setItem(row, 1, QtWidgets.QTableWidgetItem(""))

    def remove_selected_quick_access_rows(self) -> None:
        selected_rows = sorted({index.row() for index in self.quick_access_table.selectedIndexes()}, reverse=True)
        for row in selected_rows:
            self.quick_access_table.removeRow(row)

    def add_tab_node(self) -> None:
        tab_item = self._make_tree_item("tab", "tab.new", "tab.new", "100")
        self.structure_tree.addTopLevelItem(tab_item)
        self.structure_tree.setCurrentItem(tab_item)

    def add_group_node(self) -> None:
        selected = self.structure_tree.currentItem()
        if selected is None:
            QtWidgets.QMessageBox.information(self, "Add Group", "Select a tab node first.")
            return

        if selected.text(0) == "group":
            selected = selected.parent()

        if selected is None or selected.text(0) != "tab":
            QtWidgets.QMessageBox.information(self, "Add Group", "Select a tab node first.")
            return

        group_item = self._make_tree_item("group", "grp.new", "grp.new", "10")
        selected.addChild(group_item)
        selected.setExpanded(True)
        self.structure_tree.setCurrentItem(group_item)

    def add_item_node(self) -> None:
        selected = self.structure_tree.currentItem()
        if selected is None:
            QtWidgets.QMessageBox.information(self, "Add Item", "Select a group node first.")
            return

        if selected.text(0) == "item":
            selected = selected.parent()

        if selected is None or selected.text(0) != "group":
            QtWidgets.QMessageBox.information(self, "Add Item", "Select a group node first.")
            return

        item_node = self._make_tree_item("item", "", "", "", "cmd.newCommand", "small", "", "false")
        selected.addChild(item_node)
        selected.setExpanded(True)
        self.structure_tree.setCurrentItem(item_node)

    def remove_selected_tree_node(self) -> None:
        selected = self.structure_tree.currentItem()
        if selected is None:
            return

        parent = selected.parent()
        if parent is None:
            index = self.structure_tree.indexOfTopLevelItem(selected)
            self.structure_tree.takeTopLevelItem(index)
            return

        parent.removeChild(selected)

    def _read_quick_access(self) -> list[dict[str, Any]]:
        quick_access: list[dict[str, Any]] = []
        for row in range(self.quick_access_table.rowCount()):
            command_item = self.quick_access_table.item(row, 0)
            menu_item = self.quick_access_table.item(row, 1)

            command_ref = command_item.text().strip() if command_item else ""
            menu_ref = menu_item.text().strip() if menu_item else ""
            if not command_ref:
                continue

            entry: dict[str, Any] = {"commandRef": command_ref}
            if menu_ref:
                entry["menuRef"] = menu_ref
            quick_access.append(entry)

        return quick_access

    def _to_int_or_default(self, value: str, default: int) -> int:
        try:
            return int(value)
        except ValueError:
            return default

    def _to_bool_or_none(self, value: str) -> bool | None:
        normalized = value.strip().lower()
        if normalized == "true":
            return True
        if normalized == "false":
            return False
        return None

    def _read_tabs(self) -> list[dict[str, Any]]:
        tabs: list[dict[str, Any]] = []

        for i in range(self.structure_tree.topLevelItemCount()):
            tab_item = self.structure_tree.topLevelItem(i)
            if tab_item.text(0) != "tab":
                continue

            tab: dict[str, Any] = {
                "id": tab_item.text(1).strip() or f"tab.{i}",
                "labelKey": tab_item.text(2).strip() or f"tab.{i}",
                "order": self._to_int_or_default(tab_item.text(3).strip(), 100),
                "groups": [],
            }

            for j in range(tab_item.childCount()):
                group_item = tab_item.child(j)
                if group_item.text(0) != "group":
                    continue

                group: dict[str, Any] = {
                    "id": group_item.text(1).strip() or f"grp.{j}",
                    "labelKey": group_item.text(2).strip() or f"grp.{j}",
                    "priority": self._to_int_or_default(group_item.text(3).strip(), 10),
                    "items": [],
                }

                for k in range(group_item.childCount()):
                    item_node = group_item.child(k)
                    if item_node.text(0) != "item":
                        continue

                    command_ref = item_node.text(4).strip()
                    if not command_ref:
                        continue

                    ui_item: dict[str, Any] = {"commandRef": command_ref}

                    size = item_node.text(5).strip()
                    if size:
                        ui_item["size"] = size

                    menu_ref = item_node.text(6).strip()
                    if menu_ref:
                        ui_item["menuRef"] = menu_ref

                    show_text = self._to_bool_or_none(item_node.text(7))
                    if show_text is not None:
                        ui_item["showText"] = show_text

                    group["items"].append(ui_item)

                tab["groups"].append(group)

            tabs.append(tab)

        return tabs

    def save_config(self) -> None:
        if not self.validate_and_report():
            return

        ui = self.config.setdefault("ui", {})
        ui["quickAccess"] = self._read_quick_access()
        ui["tabs"] = self._read_tabs()

        try:
            with self.config_path.open("w", encoding="utf-8") as config_file:
                json.dump(self.config, config_file, ensure_ascii=False, indent=2)
                config_file.write("\n")
        except Exception as error:
            QtWidgets.QMessageBox.critical(self, "Save Error", f"Cannot save config file:\n{error}")
            return

        self.statusBar().showMessage("Config saved", 3000)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)

    config_path = DEFAULT_CONFIG_PATH
    if len(sys.argv) > 1:
        config_path = Path(sys.argv[1]).resolve()

    editor = RibbonConfigEditor(config_path)
    editor.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
