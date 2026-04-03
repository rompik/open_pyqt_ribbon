import json
import sys
from pathlib import Path
from typing import Any, Callable

from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtGui import QIcon

from ribbex import RibbonBar, RowWise

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "word.ribbon.json"
DEFAULT_LANGUAGE = "en"


def load_config(config_path: Path) -> dict[str, Any]:
    with config_path.open("r", encoding="utf-8") as config_file:
        return json.load(config_file)


def resolve_language(argv: list[str], config: dict[str, Any]) -> str:
    available = set(config.get("resources", {}).get("labels", {}).keys())
    selected = argv[1].lower() if len(argv) > 1 else DEFAULT_LANGUAGE
    if selected in available:
        return selected
    return DEFAULT_LANGUAGE


def translate(config: dict[str, Any], language: str, key: str, fallback: str | None = None) -> str:
    labels = config.get("resources", {}).get("labels", {})
    localized = labels.get(language, {}).get(key)
    if localized is not None:
        return localized

    default_localized = labels.get(DEFAULT_LANGUAGE, {}).get(key)
    if default_localized is not None:
        return default_localized

    if fallback is not None:
        return fallback
    return key


def icon_from_key(config: dict[str, Any], key: str | None) -> QIcon:
    if not key:
        return QIcon()

    icon_path = config.get("resources", {}).get("icons", {}).get(key)
    if not icon_path:
        return QIcon()

    return QIcon(str(BASE_DIR / icon_path))


def bind_execute_action(action: QtGui.QAction, execute: str, on_change_language: Callable[[str], None]) -> None:
    if execute.startswith("ui.setLanguage:"):
        target_language = execute.split(":", maxsplit=1)[1]
        action.triggered.connect(lambda _checked=False, language=target_language: on_change_language(language))


def build_menu(
    config: dict[str, Any],
    language: str,
    menu_id: str,
    command_map: dict[str, dict[str, Any]],
    menu_map: dict[str, dict[str, Any]],
    on_change_language: Callable[[str], None],
    parent: QtWidgets.QWidget,
) -> QtWidgets.QMenu:
    menu_definition = menu_map.get(menu_id, {})
    menu = QtWidgets.QMenu(parent)

    for item in menu_definition.get("items", []):
        if "commandRef" in item:
            command = command_map.get(item["commandRef"], {})
            label = translate(config, language, command.get("labelKey", ""), fallback=item["commandRef"])
            action = QtGui.QAction(icon_from_key(config, command.get("iconKey")), label, menu)
            bind_execute_action(action, command.get("execute", ""), on_change_language)
            menu.addAction(action)
            continue

        label = translate(config, language, item.get("labelKey", ""), fallback=item.get("id", "Item"))
        action = QtGui.QAction(label, menu)
        bind_execute_action(action, item.get("execute", ""), on_change_language)
        menu.addAction(action)

    return menu


def build_ribbon(
    config: dict[str, Any],
    language: str,
    on_change_language: Callable[[str], None],
) -> RibbonBar:
    ribbonbar = RibbonBar()
    ribbonbar.setApplicationIcon(QIcon(str(BASE_DIR / "word.png")))
    ribbonbar.applicationOptionButton().setToolTip(
        translate(config, language, "app.tooltip", fallback="Microsoft Word")
    )

    commands = config.get("commands", [])
    menus = config.get("menus", [])
    command_map = {command["id"]: command for command in commands}
    menu_map = {menu["id"]: menu for menu in menus}

    for qa_item in config.get("ui", {}).get("quickAccess", []):
        command = command_map.get(qa_item.get("commandRef", ""))
        if not command:
            continue

        button = QtWidgets.QToolButton()
        button.setAutoRaise(True)
        label = translate(config, language, command.get("labelKey", ""), fallback=command["id"])
        button.setText(label)
        button.setToolTip(label)

        icon = icon_from_key(config, command.get("iconKey"))
        if not icon.isNull():
            button.setIcon(icon)

        menu_ref = qa_item.get("menuRef")
        if menu_ref:
            button.setMenu(build_menu(config, language, menu_ref, command_map, menu_map, on_change_language, button))
            button.setPopupMode(QtWidgets.QToolButton.ToolButtonPopupMode.InstantPopup)

        ribbonbar.addQuickAccessButton(button)

    tabs = sorted(config.get("ui", {}).get("tabs", []), key=lambda tab: tab.get("order", 0))
    for tab in tabs:
        category_label = translate(config, language, tab.get("labelKey", ""), fallback=tab.get("id", "Tab"))
        category = ribbonbar.addCategory(category_label)

        groups = sorted(tab.get("groups", []), key=lambda group: group.get("priority", 0))
        for group in groups:
            panel_label = translate(config, language, group.get("labelKey", ""), fallback=group.get("id", "Group"))
            panel = category.addPanel(panel_label)

            for item in group.get("items", []):
                command = command_map.get(item.get("commandRef", ""))
                if not command:
                    continue

                kind = command.get("kind")
                if kind == "fontCombo":
                    layout = command.get("layout", {})
                    panel.addSmallFontComboBox(
                        colSpan=layout.get("colSpan", 1),
                        fixedHeight=layout.get("fixedHeight", False),
                    )
                    continue

                if kind == "combo":
                    layout = command.get("layout", {})
                    mode = RowWise if layout.get("mode") == "rowWise" else RowWise
                    panel.addSmallComboBox(
                        command.get("items", []),
                        fixedHeight=layout.get("fixedHeight", False),
                        mode=mode,
                    )
                    continue

                label = translate(config, language, command.get("labelKey", ""), fallback=command["id"])
                icon = icon_from_key(config, command.get("iconKey"))
                tooltip = label
                show_text = item.get("showText", True)
                size = item.get("size", "small")

                widget = None
                if kind == "toggle":
                    widget = panel.addSmallToggleButton(label, icon=icon, showText=show_text, tooltip=tooltip)
                elif size == "large":
                    widget = panel.addLargeButton(label, icon=icon, tooltip=tooltip)
                else:
                    widget = panel.addSmallButton(label, icon=icon, showText=show_text, tooltip=tooltip)

                menu_ref = item.get("menuRef")
                if menu_ref and widget is not None:
                    for action in build_menu(
                        config,
                        language,
                        menu_ref,
                        command_map,
                        menu_map,
                        on_change_language,
                        widget,
                    ).actions():
                        widget.addAction(action)

    return ribbonbar


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setFont(QtGui.QFont("Times New Roman", 8))

    window = QtWidgets.QMainWindow()
    window.setWindowIcon(QIcon(str(BASE_DIR / "word.png")))
    window.statusBar()
    central_widget = QtWidgets.QWidget()
    window.setCentralWidget(central_widget)
    layout = QtWidgets.QVBoxLayout(central_widget)
    layout.addWidget(QtWidgets.QTextEdit(), 1)

    state = {
        "config": load_config(CONFIG_PATH),
        "language": DEFAULT_LANGUAGE,
        "ribbonbar": None,
    }
    state["language"] = resolve_language(sys.argv, state["config"])

    def apply_language(language: str) -> None:
        if language not in state["config"].get("resources", {}).get("labels", {}):
            language = DEFAULT_LANGUAGE

        state["language"] = language
        window.setWindowTitle(translate(state["config"], language, "app.title", fallback="Microsoft Word"))

        new_ribbonbar = build_ribbon(state["config"], language, apply_language)
        previous_ribbonbar = state["ribbonbar"]

        window.setMenuBar(new_ribbonbar)
        state["ribbonbar"] = new_ribbonbar

        if previous_ribbonbar is not None:
            previous_ribbonbar.deleteLater()

    def reload_config() -> None:
        try:
            new_config = load_config(CONFIG_PATH)
        except Exception as error:
            window.statusBar().showMessage(f"Failed to reload ribbon config: {error}", 4000)
            return

        state["config"] = new_config
        if state["language"] not in state["config"].get("resources", {}).get("labels", {}):
            state["language"] = DEFAULT_LANGUAGE

        apply_language(state["language"])
        window.statusBar().showMessage("Ribbon config reloaded", 2000)

    apply_language(state["language"])

    reload_timer = QtCore.QTimer(window)
    reload_timer.setSingleShot(True)
    reload_timer.timeout.connect(reload_config)

    config_watcher = QtCore.QFileSystemWatcher(window)
    config_watcher.addPath(str(CONFIG_PATH))
    config_watcher.addPath(str(BASE_DIR))

    def schedule_reload() -> None:
        if CONFIG_PATH.exists() and str(CONFIG_PATH) not in config_watcher.files():
            config_watcher.addPath(str(CONFIG_PATH))
        reload_timer.start(250)

    config_watcher.fileChanged.connect(lambda _path: schedule_reload())
    config_watcher.directoryChanged.connect(lambda _path: schedule_reload())

    window.resize(1500, 1000)
    window.show()
    sys.exit(app.exec())
