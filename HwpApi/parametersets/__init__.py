from __future__ import annotations
from hwpapi.functions import from_hwpunit, to_hwpunit, convert_hwp_color_to_hex, convert_to_hwp_color
import pprint
import re
from typing import Any, Dict, List, Optional, Union, Callable, Type, Protocol, Literal, Iterable


__all__ = ['PARAMETERSET_REGISTRY', 'DIRECTION_MAP', 'CAP_FULL_SIZE_MAP', 'ALIGNMENT_MAP', 'VERT_ALIGN_MAP', 'VERT_REL_TO_MAP',
           'HORZ_REL_TO_MAP', 'HORZ_ALIGN_MAP', 'ALIGN_MAP', 'ALIGN_TYPE_MAP', 'FONTTYPE_MAP', 'TEXT_DIRECTION_MAP',
           'TEXT_ALIGN_MAP', 'LINE_SPACING_TYPE_MAP', 'LINE_WRAP_MAP', 'TEXT_WRAP_MAP', 'TEXT_FLOW_MAP',
           'LATIN_LINE_BREAK_MAP', 'NONLATIN_LINE_BREAK_MAP', 'SHADOW_TYPE_MAP', 'BACKGROUND_TYPE_MAP',
           'GRADATION_TYPE_MAP', 'ROTATION_SETTING_MAP', 'PIC_EFFECT_MAP', 'SEARCH_DIRECTION_MAP', 'BORDER_TEXT_MAP',
           'UNDERLINE_TYPE_MAP', 'OUTLINE_TYPE_MAP', 'STRIKEOUT_TYPE_MAP', 'USE_KERNING_MAP', 'DIAC_SYM_MARK_MAP',
           'USE_FONT_SPACE_MAP', 'HEADING_TYPE_MAP', 'NUMBERING_TYPE_MAP', 'NUMBER_FORMAT_MAP', 'PAGE_BREAK_MAP',
           'ALL_MAPPINGS', 'ParameterBackend', 'ComBackend', 'AttrBackend', 'PsetBackend', 'HParamBackend',
           'make_backend', 'wrap_parameterset', 'resolve_action_args', 'apply_staged_to_backend',
           'MissingRequiredError', 'PropertyDescriptor', 'IntProperty', 'BoolProperty', 'StringProperty',
           'ColorProperty', 'UnitProperty', 'MappedProperty', 'TypedProperty', 'NestedProperty', 'ArrayProperty',
           'HArrayWrapper', 'ListProperty', 'ParameterSetMeta', 'ParameterSet', 'update_from', 'GenericParameterSet',
           'BorderFill', 'Caption', 'DrawFillAttr', 'CharShape', 'ParaShape', 'ShapeObject', 'Table', 'FindReplace',
           'BulletShape', 'Cell', 'CtrlData', 'DrawArcType', 'DrawCoordInfo', 'DrawCtrlHyperlink', 'DrawEditDetail',
           'DrawImageAttr', 'DrawImageScissoring', 'DrawLayout', 'DrawLineAttr', 'DrawRectType', 'DrawResize',
           'DrawRotate', 'DrawScAction', 'DrawShadow', 'DrawShear', 'DrawTextart', 'InsertText', 'ListProperties',
           'NumberingShape', 'TabDef', 'ActionCrossRef', 'AutoFill', 'AutoNum', 'BookMark', 'BorderFillExt',
           'CaptureEnd', 'CellBorderFill', 'ChCompose', 'ChangeRome', 'CodeTable', 'ColDef', 'ConvertCase',
           'ConvertFullHalf', 'ConvertHiraToGata', 'ConvertJianFan', 'ConvertToHangul', 'DeleteCtrls', 'DocFilters',
           'DocFindInfo', 'DocumentInfo', 'DropCap', 'Dutmal', 'EngineProperties', 'EqEdit', 'ExchangeFootnoteEndNote',
           'FieldCtrl', 'FileConvert', 'FileInfo', 'FileOpen', 'FileSaveAs', 'FileSaveBlock', 'FileSendMail',
           'FileSetSecurity', 'FlashProperties', 'FootnoteShape', 'FtpDownload', 'FtpUpload', 'GotoE', 'GridInfo',
           'HeaderFooter', 'HyperLink', 'Idiom', 'IndexMark', 'InputDateStyle', 'InsertFieldTemplate', 'InsertFile',
           'Internet', 'KeyMacro', 'LinkDocument', 'ListParaPos', 'MailMergeGenerate', 'MakeContents', 'MarkpenShape',
           'MasterPage', 'MemoShape', 'MousePos', 'MovieProperties', 'OleCreation', 'PageBorderFill', 'PageDef',
           'PageHiding', 'PageNumCtrl', 'PageNumPos', 'Password', 'Preference', 'Presentation', 'Print', 'PrintToImage',
           'PrintWatermark', 'QCorrect', 'RevisionDef', 'SaveFootnote', 'ScriptMacro', 'SecDef', 'SectionApply',
           'ShapeCopyPaste', 'ShapeObjectCopyPaste', 'Sort', 'Style', 'StyleDelete', 'StyleTemplate', 'Sum',
           'SummaryInfo', 'TableCreation', 'TableDeleteLine', 'TableDrawPen', 'TableInsertLine', 'TableSplitCell',
           'TableStrToTbl', 'TableSwap', 'TableTblToStr', 'TableTemplate', 'TextCtrl', 'TextVertical',
           'UserQCommandFile', 'VersionInfo', 'ViewProperties', 'ViewStatus']

# Global registry for auto-wrapping ParameterSets
PARAMETERSET_REGISTRY = {}

# Value mappings (string ↔ int) for MappedProperty — defined in mappings.py
from .mappings import *  # noqa: F401,F403
from .mappings import ALL_MAPPINGS  # explicit for clarity


# ── Backends (COM ↔ Python) ─────────────────────────────────────────────
# See backends.py for implementations.
from .backends import (
    ParameterBackend, ComBackend, AttrBackend, PsetBackend, HParamBackend,
    _is_com, _looks_like_pset, make_backend,
)


def wrap_parameterset(pset_obj, pset_id=None):
    """
    Auto-wrap a pset object in the appropriate ParameterSet class.

    Three-tier wrapping strategy:
    1. If pset_id matches a known class, use that class
    2. If type name matches a known class, use that class
    3. Fall back to GenericParameterSet for unknown types

    Args:
        pset_obj: The pset COM object to wrap
        pset_id: Optional pset ID string (will auto-detect if not provided)

    Returns:
        Wrapped ParameterSet instance or original object if not a pset
    """
    from hwpapi.logging import get_logger
    logger = get_logger('parametersets.wrap')

    # Already wrapped?
    if isinstance(pset_obj, ParameterSet):
        return pset_obj

    # Not a pset?
    if not _looks_like_pset(pset_obj):
        return pset_obj

    # Try to get pset ID
    if pset_id is None:
        try:
            pset_id = pset_obj.GetIDStr() if hasattr(pset_obj, 'GetIDStr') else None
        except:
            pset_id = None

    # Tier 1: Known type by pset_id
    if pset_id and pset_id in PARAMETERSET_REGISTRY:
        cls = PARAMETERSET_REGISTRY[pset_id]
        logger.debug(f"Wrapping {pset_id} with {cls.__name__}")
        return cls(pset_obj)

    # Try by type name
    type_name = type(pset_obj).__name__
    if type_name in PARAMETERSET_REGISTRY:
        cls = PARAMETERSET_REGISTRY[type_name]
        logger.debug(f"Wrapping {type_name} with {cls.__name__}")
        return cls(pset_obj)

    # Tier 2: Generic wrapper
    logger.debug(f"Wrapping unknown pset {pset_id} with GenericParameterSet")
    return GenericParameterSet(pset_obj, pset_id=pset_id)


def resolve_action_args(app: Any, action_name: str, hnode: Any) -> tuple[str, Any]:
    """
    Resolve action arguments for HAction.GetDefault/Execute calls.

    Args:
        app: HWP application object
        action_name: Action name like "FindReplace"
        hnode: HParameterSet node like app.api.HParameterSet.HFindReplace

    Returns:
        Tuple of (action_name, arg) where arg is hnode.HSet if available, else hnode
    """
    # Try to use .HSet if available (preferred for some actions)
    if hasattr(hnode, 'HSet'):
        return (action_name, hnode.HSet)
    else:
        return (action_name, hnode)


def apply_staged_to_backend(backend: ParameterBackend, staged: dict, prefix: str = "") -> None:
    """
    Apply staged changes to any backend type, handling dotted keys recursively.
    Works for both ComBackend (map-style) and HParamBackend (tree-style).

    Args:
        backend: Any ParameterBackend implementation
        staged: Dict of staged changes (may contain nested dicts)
        prefix: Key prefix for recursion (internal use)
    """
    for key, value in staged.items():
        full_key = f"{prefix}.{key}" if prefix else key

        if isinstance(value, dict):
            # Recursively handle nested dicts
            apply_staged_to_backend(backend, value, full_key)
        elif isinstance(value, ParameterSet):
            # Handle nested ParameterSet - apply its staged changes
            if hasattr(value, '_staged') and value._staged:
                apply_staged_to_backend(backend, value._staged, full_key)
        else:
            # Apply primitive value
            try:
                backend.set(full_key, value)
            except Exception as e:
                # Re-raise with context
                raise RuntimeError(f"Failed to apply staged value '{full_key}': {e}") from e



# ── Property descriptors and HArrayWrapper ──────────────────────────────
from .properties import (
    MissingRequiredError, PropertyDescriptor, IntProperty, BoolProperty,
    StringProperty, ColorProperty, UnitProperty, MappedProperty, TypedProperty,
    NestedProperty, ArrayProperty, HArrayWrapper, ListProperty,
)

# ── ParameterSet base class, metaclass, GenericParameterSet ─────────────
from .base import (
    ParameterSetMeta, ParameterSet, GenericParameterSet,
    _pascal_to_snake, _snake_to_pascal, update_from,
)

# ── ParameterSet subclasses ─────────────────────────────────────────────
# 143 subclasses organized by HWP functional domain. Each domain module
# auto-registers its classes to PARAMETERSET_REGISTRY via ParameterSetMeta.
from .sets import *  # noqa: F401,F403


def _finalize_attr_lookups():
    """Populate snake_case/lowercase entries in each class's _attr_lookup.

    Must be called AFTER all ParameterSet subclasses are defined so that
    PARAMETERSET_REGISTRY contains the full list.
    """
    for cls in list(PARAMETERSET_REGISTRY.values()):
        if not hasattr(cls, '_all_properties'):
            continue
        lookup = cls.__dict__.get('_attr_lookup')
        if lookup is None:
            continue
        for key in cls._all_properties:
            snake = _pascal_to_snake(key)
            lookup.setdefault(snake, key)
            lookup.setdefault(snake.lower(), key)

_finalize_attr_lookups()
