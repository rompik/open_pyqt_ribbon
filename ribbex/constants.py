from enum import IntEnum

from PyQt6 import QtGui


class RibbonCategoryStyle(IntEnum):
    """The button style of a category."""

    Normal = 0
    Context = 1


Normal = RibbonCategoryStyle.Normal
Context = RibbonCategoryStyle.Context


#: A list of context category colors
contextColors = [
    QtGui.QColor(201, 89, 156),  # Rose red
    QtGui.QColor(242, 203, 29),  # Yellow
    QtGui.QColor(255, 157, 0),  # Orange
    QtGui.QColor(14, 81, 167),  # Blue
    QtGui.QColor(228, 0, 69),  # Red
    QtGui.QColor(67, 148, 0),  # Green
]


class RibbonSpaceFindMode(IntEnum):
    """Mode to find available space in a grid layout, ColumnWise or RowWise."""

    ColumnWise = 0
    RowWise = 1


ColumnWise = RibbonSpaceFindMode.ColumnWise
RowWise = RibbonSpaceFindMode.RowWise


class RibbonStyle(IntEnum):
    Default = 0
    Debug = 1


Debug = RibbonStyle.Debug
Default = RibbonStyle.Default


class RibbonButtonStyle(IntEnum):
    """Button style, Small, Medium, or Large."""

    Small = 0
    Medium = 1
    Large = 2


Small = RibbonButtonStyle.Small
Medium = RibbonButtonStyle.Medium
Large = RibbonButtonStyle.Large
