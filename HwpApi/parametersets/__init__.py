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

class MissingRequiredError(ValueError):
    """Raised when required parameters are missing during apply()."""
    pass


class PropertyDescriptor:
    """
    Descriptor for parameter properties backed by a staged ParameterSet.

    Features:
    - Reads prefer staged values, then snapshot, then default.
    - Writes stage the value; nothing is sent until ParameterSet.apply().
    - Optional automatic wrapping of nested ParameterSets via `wrap=...`.
    """

    @property
    def logger(self):
        from hwpapi.logging import get_logger
        return get_logger("parametersets.PropertyDescriptor")
    def __init__(
        self,
        key: str,
        doc: str = "",
        *,
        to_python: Optional[Callable[[Any], Any]] = None,
        to_backend: Optional[Callable[[Any], Any]] = None,
        default: Any = None,
        readonly: bool = False,
        required: bool = False,
        wrap: Optional[Type["ParameterSet"]] = None,  # <-- NEW: nested ParameterSet subclass
    ):
        self.key = key
        self.doc = doc
        self.name: Optional[str] = None
        self.to_python = to_python
        self.to_backend = to_backend
        self.default = default
        self.readonly = readonly
        self.required = required
        self.wrap = wrap

    def __set_name__(self, owner, name):
        self.name = name
        reg = getattr(owner, "_property_registry", None)
        if reg is None:
            reg = {}
            setattr(owner, "_property_registry", reg)
        reg[name] = self

    def __get__(self, instance: Optional["ParameterSet"], owner):
        if instance is None:
            return self

        # Get staged/snapshot value
        val = instance._ps_get(self)

        # Automatic wrapping for nested ParameterSets
        if self.wrap is not None:
            # Serve from cache if we already wrapped this key
            cached = instance._wrapper_cache.get(self.key)
            if cached is not None:
                return cached

            # If staged/snapshot holds an existing ParameterSet, cache and return it
            if isinstance(val, ParameterSet):
                instance._wrapper_cache[self.key] = val
                return val

            # If backend returned a raw nested object (COM or dict-like), wrap it now
            if val is not None:
                # Handle both lambda functions (like lambda: CharShape) and direct class references
                if callable(self.wrap):
                    try:
                        # Try calling with no arguments first (for lambda: ClassName patterns)
                        wrapper_class = self.wrap()
                        wrapped = wrapper_class(val) if val != {} else wrapper_class()
                    except TypeError:
                        # If that fails, try calling with val as argument (for direct callable patterns)
                        wrapped = self.wrap(val)
                else:
                    # Direct class reference
                    wrapped = self.wrap(val) if val != {} else self.wrap()

                instance._wrapper_cache[self.key] = wrapped
                # Also stage the wrapper so reads stay consistent
                instance._staged[self.key] = wrapped
                return wrapped

            # If value is None, fall through to default handling below

        # Non-wrapped (or None) path - but for wrapped properties, create empty object
        if val is None:
            # For wrapped properties, return an empty instance instead of None
            if self.wrap is not None:
                # Create empty wrapped object and cache it
                if callable(self.wrap):
                    try:
                        wrapper_class = self.wrap()
                        wrapped = wrapper_class()
                    except TypeError:
                        wrapped = self.wrap()
                else:
                    wrapped = self.wrap()

                instance._wrapper_cache[self.key] = wrapped
                instance._staged[self.key] = wrapped
                return wrapped

            # For non-wrapped properties, return default if available
            if self.default is not None:
                return self.default

        # Auto-wrap pset objects using registry (Tier 1/2)
        if val is not None and _looks_like_pset(val):
            val = wrap_parameterset(val, self.key)

        return self.to_python(val) if (val is not None and self.to_python) else val

    def __set__(self, instance: "ParameterSet", value: Any):
        if self.readonly:
            raise AttributeError(f"'{self.name}' is read-only")

        # If this property is a nested ParameterSet, normalize on assignment:
        if self.wrap is not None:
            # Allow passing: ParameterSet, raw COM object, or dict-like
            if isinstance(value, ParameterSet):
                wrapped = value
            else:
                # If dict/raw object given, create a wrapper
                wrapped = self.wrap(value if value is not None else {})
            # Keep cache consistent so subsequent gets reuse same object
            instance._wrapper_cache[self.key] = wrapped
            # Stage the wrapper (not the raw) — parent apply() will unwrap
            instance._ps_set(self, wrapped)
            return

        # Primitive path: run to_backend if provided
        v = self.to_backend(value) if (value is not None and self.to_backend) else value
        instance._ps_set(self, v)

    def _get_value(self, instance):
        return instance._ps_get(self)
    def _set_value(self, instance, value):
        instance._ps_set(self, value)
    def _del_value(self, instance): return instance._ps_del(self)


class IntProperty(PropertyDescriptor):
    """Property descriptor for integer values with optional range validation."""

    def __init__(self, key: str, doc: str, min_val: Optional[int] = None, max_val: Optional[int] = None):
        super().__init__(key, doc)
        self.min_val = min_val
        self.max_val = max_val

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return int(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if not isinstance(value, (int, float)):
            raise TypeError(f"Value for '{self.key}' must be numeric")

        value = int(value)

        if self.min_val is not None and value < self.min_val:
            raise ValueError(f"Value {value} for '{self.key}' is below minimum {self.min_val}")
        if self.max_val is not None and value > self.max_val:
            raise ValueError(f"Value {value} for '{self.key}' is above maximum {self.max_val}")

        return self._set_value(instance, value)


class BoolProperty(PropertyDescriptor):
    """Property descriptor for boolean values (0 or 1)."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return bool(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if isinstance(value, bool):
            numeric_value = 1 if value else 0
        elif isinstance(value, (int, float)):
            numeric_value = 1 if value else 0
        else:
            raise TypeError(f"Value for '{self.key}' must be boolean or numeric")

        return self._set_value(instance, numeric_value)


class StringProperty(PropertyDescriptor):
    """Property descriptor for string values."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return str(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if not isinstance(value, str):
            value = str(value)

        return self._set_value(instance, value)


class ColorProperty(PropertyDescriptor):
    """Property descriptor for color values with hex conversion."""

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return convert_hwp_color_to_hex(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        numeric_value = convert_to_hwp_color(value)
        return self._set_value(instance, numeric_value)


class UnitProperty(PropertyDescriptor):
    """
    Property descriptor for unit-based values with automatic conversion.

    Supports automatic conversion between user-friendly units (mm, cm, in, pt)
    and HWP's internal HWPUNIT format.

    Attributes:
        key (str): Parameter key in HWP
        doc (str): Documentation string
        default_unit (str): Default unit when bare numbers are provided (default: "mm")

    Example:
        >>> class PageDef(ParameterSet):
        ...     width = UnitProperty("Width", "Page width", default_unit="mm")
        >>> page = PageDef(action.CreateSet())
        >>> page.width = "210mm"   # String with unit
        >>> page.width = "21cm"    # Auto-converts to HWPUNIT
        >>> page.width = 210       # Bare number, assumes mm
        >>> print(page.width)      # Returns value in mm
        210.0
    """

    def __init__(self, key: str, doc: str, default_unit: str = "mm"):
        super().__init__(key, doc)
        self.default_unit = default_unit

    def __get__(self, instance, owner):
        """Get value in default unit."""
        if instance is None:
            return self
        value = self._get_value(instance)
        if value is None:
            return None
        # Convert from HWPUNIT to default unit
        return from_hwpunit(value, self.default_unit)

    def __set__(self, instance, value):
        """
        Set value from number or string with unit.

        Args:
            value: Number or string like "210mm", "21cm", "8.27in", "12pt"
        """
        if value is None:
            return self._del_value(instance)

        # Use enhanced to_hwpunit which handles both numbers and strings
        try:
            hwp_value = to_hwpunit(value, self.default_unit)
        except (ValueError, TypeError) as e:
            raise TypeError(
                f"Value for '{self.key}' must be numeric or unit string "
                f"(e.g., '210mm', '21cm', 210). Got: {value}"
            ) from e

        return self._set_value(instance, hwp_value)


class MappedProperty(PropertyDescriptor):
    """Property descriptor for mapped values (string <-> integer)."""

    def __init__(self, key: str, mapping: Dict[str, int], doc: str):
        super().__init__(key, doc)
        self.mapping = mapping
        self.reverse_mapping = {v: k for k, v in mapping.items()}

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        if value is None:
            return None
        return self.reverse_mapping.get(value, value)

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if isinstance(value, str):
            if value not in self.mapping:
                valid_keys = list(self.mapping.keys())
                raise ValueError(f"Invalid value '{value}' for '{self.key}'. Valid options: {valid_keys}")
            numeric_value = self.mapping[value]
        elif isinstance(value, (int, float)):
            numeric_value = int(value)
            if numeric_value not in self.reverse_mapping:
                valid_values = list(self.reverse_mapping.keys())
                raise ValueError(f"Invalid numeric value '{numeric_value}' for '{self.key}'. Valid values: {valid_values}")
        else:
            raise TypeError(f"Value for '{self.key}' must be string or numeric")

        return self._set_value(instance, numeric_value)


class TypedProperty(PropertyDescriptor):
    """
    Alias for PropertyDescriptor for typed parameter sets.
    Inherit all logic from PropertyDescriptor; do not override __get__ or __set__.
    Use this for clarity when defining typed sub-ParameterSets.
    """
    def __init__(self, key: str, doc: str = "", wrap=None, **kwargs):
        """Initialize TypedProperty with support for positional wrap argument."""
        super().__init__(key, doc, wrap=wrap, **kwargs)


    # Inherits __get__ and __set__ from PropertyDescriptor


class NestedProperty(PropertyDescriptor):
    """
    Auto-creating nested ParameterSet property.

    Automatically calls backend.create_itemset(key, setid) on first access,
    wraps the result in the specified ParameterSet class, and caches the instance.

    Attributes:
        key (str): Parameter key in HWP (e.g., "FindCharShape")
        setid (str): SetID for CreateItemSet call (e.g., "CharShape")
        param_class (Type[ParameterSet]): ParameterSet class to wrap
        doc (str): Documentation string
        _cache_attr (str): Internal cache attribute name

    Example:
        >>> class FindReplace(ParameterSet):
        ...     find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape)
        >>> pset = FindReplace(action.CreateSet())
        >>> pset.find_char_shape.bold = True  # Auto-creates CharShape!
    """

    def __init__(self, key: str, setid: str, param_class: Type["ParameterSet"], doc: str = ""):
        super().__init__(key, doc)
        self.setid = setid
        self.param_class = param_class
        self._cache_attr = f"_nested_cache_{key}"

    def __get__(self, instance: "ParameterSet", owner) -> "ParameterSet":
        """
        Get nested ParameterSet, creating it if needed.

        Returns cached instance if available, otherwise:
        1. Calls backend.create_itemset(key, setid) to create COM object
        2. Wraps in param_class
        3. Caches for future access
        4. Returns wrapped instance

        Raises:
            RuntimeError: If ParameterSet not bound to backend
        """
        if instance is None:
            return self

        # Check cache first (subsequent access)
        if hasattr(instance, self._cache_attr):
            return getattr(instance, self._cache_attr)

        # Verify backend is available
        if instance._backend is None:
            raise RuntimeError(
                f"Cannot access nested property '{self.key}': "
                "ParameterSet not bound. Call .bind() or pass parameterset= to __init__"
            )

        # Create via CreateItemSet (first access)
        if hasattr(instance._backend, 'create_itemset'):
            # PsetBackend - use CreateItemSet
            nested_pset_com = instance._backend.create_itemset(self.key, self.setid)
            nested_wrapped = self.param_class(nested_pset_com)
        else:
            # Fallback for HParamBackend or other backends
            try:
                nested_com = instance._backend.get(self.key)
                nested_wrapped = self.param_class(nested_com)
            except (KeyError, AttributeError):
                # Create unbound instance as last resort
                nested_wrapped = self.param_class()

        # Cache for subsequent access
        setattr(instance, self._cache_attr, nested_wrapped)
        return nested_wrapped

    def __set__(self, instance: "ParameterSet", value: "ParameterSet"):
        """
        Allow direct assignment of ParameterSet instance.

        Args:
            instance: Parent ParameterSet
            value: ParameterSet instance to assign

        Raises:
            TypeError: If value is not the correct ParameterSet type
        """
        if not isinstance(value, self.param_class):
            raise TypeError(
                f"Value for '{self.key}' must be {self.param_class.__name__}, "
                f"got {type(value).__name__}"
            )

        # Cache the provided instance
        setattr(instance, self._cache_attr, value)

        # Stage for apply()
        instance._staged[self.key] = value

class ArrayProperty(PropertyDescriptor):
    """
    Auto-creating array property for HArray COM objects.

    Provides Pythonic list-like interface with automatic type conversion
    and optional length validation.

    Attributes:
        key (str): Parameter key in HWP
        item_type (Type): Type of array elements (int, float, str, tuple)
        doc (str): Documentation string
        min_length (Optional[int]): Minimum array length
        max_length (Optional[int]): Maximum array length
        _cache_attr (str): Internal cache attribute name

    Example:
        >>> class TabDef(ParameterSet):
        ...     tab_stops = ArrayProperty("TabStops", int, "Tab stop positions")
        >>> tab_def = TabDef(action.CreateSet())
        >>> tab_def.tab_stops = [1000, 2000, 3000]
        >>> tab_def.tab_stops.append(4000)
    """

    def __init__(self, key: str, item_type: Type, doc: str = "",
                 min_length: Optional[int] = None, max_length: Optional[int] = None):
        super().__init__(key, doc)
        self.item_type = item_type
        self.min_length = min_length
        self.max_length = max_length
        self._cache_attr = f"_array_cache_{key}"

    def __get__(self, instance: "ParameterSet", owner) -> "HArrayWrapper":
        """
        Get HArrayWrapper, creating it if needed.

        Returns:
            HArrayWrapper providing list-like interface to HArray
        """
        if instance is None:
            return self

        # Check cache first
        if hasattr(instance, self._cache_attr):
            return getattr(instance, self._cache_attr)

        # Try to get existing HArray from backend
        if instance._backend is not None:
            try:
                harray_com = instance._backend.get(self.key)
                if harray_com is not None:
                    wrapper = HArrayWrapper(harray_com, self.item_type,
                                           instance._backend, self.key)
                    setattr(instance, self._cache_attr, wrapper)
                    return wrapper
            except (KeyError, AttributeError):
                pass

        # Return empty wrapper (will create HArray on modification)
        wrapper = HArrayWrapper(None, self.item_type, instance._backend, self.key)
        setattr(instance, self._cache_attr, wrapper)
        return wrapper

    def __set__(self, instance: "ParameterSet", value: Union[List, Tuple, None]):
        """
        Set array from Python list/tuple.

        Args:
            value: Python list/tuple to set, or None to clear

        Raises:
            TypeError: If value is not list/tuple
            ValueError: If length constraints violated
        """
        if value is None:
            # Clear array
            wrapper = HArrayWrapper(None, self.item_type, instance._backend, self.key)
            setattr(instance, self._cache_attr, wrapper)
            instance._staged[self.key] = []
            return

        # Validate type
        if not isinstance(value, (list, tuple)):
            raise TypeError(
                f"Value for '{self.key}' must be list or tuple, "
                f"got {type(value).__name__}"
            )

        value_list = list(value)

        # Validate length
        if self.min_length is not None and len(value_list) < self.min_length:
            raise ValueError(
                f"Array '{self.key}' must have at least {self.min_length} items, "
                f"got {len(value_list)}"
            )
        if self.max_length is not None and len(value_list) > self.max_length:
            raise ValueError(
                f"Array '{self.key}' must have at most {self.max_length} items, "
                f"got {len(value_list)}"
            )

        # Create wrapper with initial values
        wrapper = HArrayWrapper(None, self.item_type, instance._backend, self.key,
                                initial_values=value_list)
        setattr(instance, self._cache_attr, wrapper)
        instance._staged[self.key] = value_list

class HArrayWrapper:
    """
    Pythonic wrapper around HWP's HArray COM object.

    Provides full list interface: indexing, iteration, append, insert, etc.
    Automatically syncs changes with underlying COM HArray when available.

    Attributes:
        _harray: COM HArray object (or None if not created yet)
        _item_type: Python type for array elements
        _backend: ParameterSet backend (for array creation)
        _key: Parameter key
        _local_cache: Python list holding current values
    """

    def __init__(self, harray_com: Any, item_type: Type,
                 backend: Optional[Any] = None, key: Optional[str] = None,
                 initial_values: Optional[List] = None):
        self._harray = harray_com
        self._item_type = item_type
        _backend = backend
        self._key = key
        self._local_cache = list(initial_values) if initial_values else []

        # Sync from COM if available
        if self._harray is not None:
            self._sync_from_com()

    def _sync_from_com(self):
        """Load values from COM HArray into local cache."""
        if self._harray is None:
            return
        try:
            count = self._harray.Count
            self._local_cache = [
                self._convert_from_com(self._harray.Item(i))
                for i in range(count)
            ]
        except Exception:
            pass  # Keep local cache if COM access fails

    def _sync_to_com(self):
        """Push local cache to COM HArray."""
        if self._harray is None:
            return
        try:
            # Clear existing
            while self._harray.Count > 0:
                self._harray.RemoveAt(0)
            # Add all from cache
            for value in self._local_cache:
                self._harray.Add(self._convert_to_com(value))
        except Exception:
            pass  # Keep local cache if sync fails

    def _convert_to_com(self, value: Any) -> Any:
        """Convert Python value to COM-compatible type."""
        if self._item_type == tuple:
            return value  # Tuples may need special handling
        return self._item_type(value)

    def _convert_from_com(self, value: Any) -> Any:
        """Convert COM value to Python type."""
        if self._item_type == tuple:
            return value
        return self._item_type(value)

    # List interface implementation
    def __len__(self) -> int:
        return len(self._local_cache)

    def __getitem__(self, index: int) -> Any:
        return self._local_cache[index]

    def __setitem__(self, index: int, value: Any):
        self._local_cache[index] = self._convert_to_com(value)
        self._sync_to_com()

    def __delitem__(self, index: int):
        del self._local_cache[index]
        self._sync_to_com()

    def __iter__(self):
        return iter(self._local_cache)

    def __repr__(self) -> str:
        return f"HArrayWrapper({self._local_cache})"

    def append(self, value: Any):
        """Add item to end of array."""
        self._local_cache.append(self._convert_to_com(value))
        self._sync_to_com()

    def insert(self, index: int, value: Any):
        """Insert item at index."""
        self._local_cache.insert(index, self._convert_to_com(value))
        self._sync_to_com()

    def remove(self, value: Any):
        """Remove first occurrence of value."""
        self._local_cache.remove(value)
        self._sync_to_com()

    def pop(self, index: int = -1) -> Any:
        """Remove and return item at index."""
        value = self._local_cache.pop(index)
        self._sync_to_com()
        return value

    def clear(self):
        """Remove all items."""
        self._local_cache.clear()
        self._sync_to_com()

    def to_list(self) -> List:
        """Convert to plain Python list."""
        return list(self._local_cache)

class ListProperty(PropertyDescriptor):
    """Property descriptor for list values."""

    def __init__(self, key: str, doc: str, item_type: Optional[Type] = None,
                 min_length: Optional[int] = None, max_length: Optional[int] = None):
        super().__init__(key, doc)
        self.item_type = item_type
        self.min_length = min_length
        self.max_length = max_length

    def __get__(self, instance, owner):
        if instance is None:
            return self
        value = self._get_value(instance)
        return list(value) if value is not None else None

    def __set__(self, instance, value):
        if value is None:
            return self._del_value(instance)

        if not isinstance(value, (list, tuple)):
            raise TypeError(f"Value for '{self.key}' must be a list or tuple")

        value = list(value)

        # Length validation
        if self.min_length is not None and len(value) < self.min_length:
            raise ValueError(f"List for '{self.key}' must have at least {self.min_length} items")
        if self.max_length is not None and len(value) > self.max_length:
            raise ValueError(f"List for '{self.key}' must have at most {self.max_length} items")

        # Type validation
        if self.item_type is not None:
            for i, item in enumerate(value):
                if not isinstance(item, self.item_type):
                    if self.item_type == tuple and isinstance(item, (list, tuple)) and len(item) == 2:
                        value[i] = tuple(item)
                    else:
                        try:
                            value[i] = self.item_type(item)
                        except (ValueError, TypeError):
                            raise TypeError(f"Item {i} in list for '{self.key}' must be of type {self.item_type.__name__}")

        return self._set_value(instance, value)


class ParameterSetMeta(type):
    """Metaclass for automatic property registration and ParameterSet registration."""

    def __new__(cls, name, bases, namespace):
        # Collect property descriptors
        properties = {}
        for key, value in namespace.items():
            if isinstance(value, PropertyDescriptor):
                properties[key] = value

        # Store property registry
        namespace['_property_registry'] = properties

        # Create the class
        new_class = super().__new__(cls, name, bases, namespace)

        # Collect all properties from base classes too
        all_properties = {}
        for base in reversed(new_class.__mro__):
            if hasattr(base, '_property_registry'):
                all_properties.update(base._property_registry)

        new_class._all_properties = all_properties

        # Build O(1) attribute lookup table: maps name variants to actual key.
        # Used by ParameterSet.__getattr__/__setattr__ to avoid dir() iteration.
        # Populated lazily to handle forward references to _pascal_to_snake.
        attr_lookup = {}
        for key in all_properties:
            attr_lookup[key] = key               # exact match
            attr_lookup[key.lower()] = key       # case-insensitive
        new_class._attr_lookup = attr_lookup
        # snake_case entries added after _pascal_to_snake is defined via
        # ParameterSetMeta._finalize_lookup() called at module init.

        # Auto-register all ParameterSet subclasses (NEW)
        if name not in ('ParameterSet', 'GenericParameterSet'):
            # Register by class name
            PARAMETERSET_REGISTRY[name] = new_class
            PARAMETERSET_REGISTRY[name.lower()] = new_class

            # Register by _pset_id if available
            if '_pset_id' in namespace:
                PARAMETERSET_REGISTRY[namespace['_pset_id']] = new_class

        return new_class


def _pascal_to_snake(name: str) -> str:
    """Convert PascalCase to snake_case.

    Examples:
        FaceNameHangul -> face_name_hangul
        Bold -> bold
        TextColor -> text_color
    """
    # Insert underscore before uppercase letters (except at start)
    snake = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', name)
    snake = re.sub('([a-z0-9])([A-Z])', r'\1_\2', snake)
    return snake.lower()


def _snake_to_pascal(name: str) -> str:
    """Convert snake_case to PascalCase.

    Examples:
        face_name_hangul -> FaceNameHangul
        bold -> Bold
        text_color -> TextColor
    """
    parts = name.split('_')
    return ''.join(word.capitalize() for word in parts)


class ParameterSet(metaclass=ParameterSetMeta):
    """
    Unified ParameterSet supporting both pset and HSet backends.

    - Prioritizes pset objects (from action.CreateSet()) for direct operation
    - Falls back to HSet objects for backward compatibility
    - Supports immediate parameter changes (no staging required for pset)
    - Maintains staging for HSet objects to preserve existing behavior

    Usage Example (pset-based - preferred):
        # 1. Get the ParameterSet for an action (e.g., FindReplace)
        pset = actions.FindReplace.pset

        # 2. Set parameters directly (immediate effect)
        pset.find_string = "foo"
        pset.replace_string = "bar"
        pset.match_case = True

        # 3. Run the action (parameters already set)
        actions.FindReplace.run()

    Usage Example (HSet-based - legacy):
        # Same as before with staging and apply()
        pset = actions.FindReplace.get_pset()
        pset.find_string = "foo"
        pset.apply()
        actions.FindReplace.run(pset)
"""

    # Optional class-level expected SetID. Subclasses can override.
    REQUIRED_SETID: Optional[str] = None

    _property_registry: Dict[str, PropertyDescriptor]  # populated by descriptors

    def __init__(
        self,
        parameterset: Any = None,  # <-- now optional
        *,
        backend_factory: Optional[Callable[[Any], ParameterBackend]] = None,
        initial: Optional[Dict[str, Any]] = None,
        expected_setid: Optional[str] = None,  # <-- new
        app_instance: Any = None,  # <-- new: reference to App instance
        **kwargs,
    ):
        from hwpapi.logging import get_logger
        self.logger = get_logger("parametersets.ParameterSet")
        if backend_factory is None:
            backend_factory = make_backend

        # Expected SetID (instance preference > class default)
        self._expected_setid: Optional[str] = (
            expected_setid
            if expected_setid is not None
            else getattr(self.__class__, "REQUIRED_SETID", None)
        )

        # Store App instance reference for HSet synchronization
        self._app_instance: Any = app_instance

        # Placeholders before binding
        self._raw: Any = None
        self._backend: Optional[ParameterBackend] = None
        self._pset: Any = None
        self._is_pset: bool = False

        # A stable snapshot of current remote values (keyed by descriptor.key)
        self._snapshot: Dict[str, Any] = {}
        # Staged changes (key -> value) not yet applied
        self._staged: Dict[str, Any] = {}
        # Keys marked for deletion
        self._deleted: set[str] = set()
        # Cache of wrapped nested ParameterSets keyed by raw parameter key
        self._wrapper_cache: Dict[str, ParameterSet] = {}

        # Bind immediately if provided, otherwise start unbound
        if parameterset is not None:
            self.bind(parameterset, backend_factory=backend_factory)
            # Take a snapshot of all current values for display/serialization
            self._snapshot = self._take_initial_snapshot()
        else:
            # start empty; snapshot stays empty until bind+reload
            pass

        # Stage initial values (do NOT send yet)
        if initial:
            self.update(initial)
        if kwargs:
            self.update(kwargs)

    def _take_initial_snapshot(self):
        """
        Take a snapshot of all current values from the COM object for display/serialization.
        """
        snapshot = {}
        for name in self._property_registry:
            try:
                value = getattr(self, name, None)
                if isinstance(value, ParameterSet):
                    value = value.serialize()
                elif hasattr(value, "_oleobj_"):
                    value = _safe_com_serialize(value)
                snapshot[name] = value
            except Exception as e:
                snapshot[name] = f"<COM error: {e}>"
        return snapshot

    def bind(
        self,
        parameterset: Any,
        *,
        backend_factory: Optional[Callable[[Any], ParameterBackend]] = None,
    ):
        """
        Bind/attach a raw parameterset to this instance (or re-bind a new one).
        Validates SetID presence and (optionally) equality to expected SetID.
        """
        if parameterset is None:
            raise TypeError("bind(): 'parameterset' must not be None")

        # Must expose SetID
        if not hasattr(parameterset, "SetID") and (
            self._expected_setid is not None
            and self._expected_setid not in str(parameterset.__class__)
        ):
            raise TypeError("bind(): provided object has no 'SetID' attribute")

        # If we know what to expect, enforce it
        if self._expected_setid is not None and (
            (self._expected_setid not in str(parameterset.__class__))
            or (getattr(parameterset, "SetID", None) != self._expected_setid)
        ):
            raise ValueError(
                f"bind(): parameterset.SetID={getattr(parameterset,'SetID',None)!r} "
                f"does not match expected {self._expected_setid!r}"
            )

        if backend_factory is None:
            backend_factory = make_backend

        # Commit the binding
        self._raw = parameterset
        self._backend = backend_factory(parameterset)
        self._pset = parameterset
        self._is_pset = True

        # Refresh snapshot from the newly bound backend
        self.reload()
        return self

    @property
    def parameterset(self):
        """Return the underlying raw object (COM ParameterSet or Python object), or None if unbound."""
        self.apply()
        return self._raw

    def reload(self):
        """Refresh in-memory snapshot from backend and clear staged edits (but keep wrapper cache coherent)."""
        self._snapshot.clear()

        # If not yet bound, nothing to load; keep staged/deleted but clean them for safety.
        if self._backend is None:
            self._staged.clear()
            self._deleted.clear()
            return self

        for name, desc in self._property_registry.items():
            try:
                self._snapshot[desc.key] = self._backend.get(desc.key)
            except Exception:
                self._snapshot[desc.key] = None
        self._staged.clear()
        self._deleted.clear()
        return self

    def apply(
        self,
        overrides: Optional[Dict[str, Any]] = None,
        *,
        require: Literal["error", "warn", "skip"] = "error",
        only_overrides: bool = False,
        parameterset: Any = None,  # <-- NEW: allow binding at apply-time
        **kwargs,
    ):
        """
        Flush staged changes (and deletions) to backend.

        Parameters
        ----------
        overrides : dict | None
            Additional values (by attribute name) to stage just-in-time.
        require : {"error","warn","skip"}
            Control missing-required behavior at apply-time.
        only_overrides : bool
            If True, ignore previously staged edits and apply ONLY overrides/**kwargs.
        parameterset : Any | None
            Optionally (re)bind a raw parameterset right before applying. Required if not already bound.
        **kwargs
            Same as overrides; convenient for inline use.
        """
        # Ensure we are bound; if not, we must bind now and validate SetID
        if self._backend is None:
            if parameterset is None:
                raise RuntimeError(
                    "apply(): no parameterset is bound. Pass 'parameterset=' or call .bind() first."
                )
            self.bind(parameterset)

        # Optionally ignore prior staged changes
        saved_staged = None
        saved_deleted = None
        if only_overrides:
            saved_staged = self._staged
            saved_deleted = self._deleted
            self._staged = {}
            self._deleted = set()

        # Add incoming values to staged
        if overrides:
            self.update(overrides)
        if kwargs:
            self.update(kwargs)

        # Validate requireds
        if require != "skip":
            missing = self._missing_required()
            if missing:
                msg = f"Missing required parameters: {', '.join(missing)}"
                if require == "error":
                    if only_overrides:
                        self._staged = saved_staged or {}
                        self._deleted = saved_deleted or set()
                    raise MissingRequiredError(msg)
                else:
                    print("[ParameterSet] WARN:", msg)

        # Deletes first
        for key in list(self._deleted):
            try:
                self._backend.delete(key)
            finally:
                self._snapshot[key] = None
        self._deleted.clear()

        # Writes next (cascade to nested ParameterSets and unwrap)
        for key, value in list(self._staged.items()):
            if isinstance(value, ParameterSet):
                # Ensure nested staged values are flushed first
                # Only apply if the nested ParameterSet is bound
                if value._backend is not None:
                    value.apply(require=require)
                    raw_value = value.parameterset
                else:
                    # Unbound nested ParameterSet - skip setting it
                    # This can happen with TypedProperty empty wrapped objects
                    # Don't try to set empty dict or None - just skip
                    continue
            else:
                raw_value = value
            self._backend.set(key, raw_value)
            self._snapshot[key] = raw_value
        self._staged.clear()

        # Restore pre-existing staged edits if only_overrides=True
        if only_overrides:
            for k, v in (saved_staged or {}).items():
                if k not in self._snapshot or self._snapshot[k] != v:
                    self._staged[k] = v

        # Special handling for HSet-based parameter sets (e.g., FindReplace, FindDlg, FindAll)
        # These actions use global HParameterSet state instead of local parameter sets
        self._sync_hset_global_state()

        return self

    def discard(self):
        """Drop staged edits and deletions (keep snapshot)."""
        self._staged.clear()
        self._deleted.clear()
        return self


    def __getattr__(self, name: str):
        """
        Attribute access with snake_case/PascalCase/case-insensitive support.
        Uses O(1) lookup table built by ParameterSetMeta.
        """
        if name.startswith('_'):
            raise AttributeError(f"'{type(self).__name__}' object has no attribute '{name}'")

        lookup = type(self).__dict__.get('_attr_lookup') or getattr(type(self), '_attr_lookup', {})
        # Try: exact → lowercase → snake_case lookup
        actual_key = lookup.get(name) or lookup.get(name.lower())
        if actual_key is None:
            # Fallback for snake_case (handles lazy-added entries or legacy)
            pascal_name = _snake_to_pascal(name)
            actual_key = lookup.get(pascal_name) or lookup.get(pascal_name.lower())

        if actual_key is not None:
            descriptor = type(self)._all_properties.get(actual_key)
            if descriptor is not None:
                return descriptor.__get__(self, type(self))

        raise AttributeError(
            f"'{type(self).__name__}' object has no attribute '{name}'. "
            f"Available attributes: {', '.join(sorted(self.attributes_names)[:5])}..."
        )

    def __setattr__(self, name: str, value: Any):
        """
        Attribute assignment with snake_case/PascalCase/case-insensitive support.
        Uses O(1) lookup table built by ParameterSetMeta.
        """
        if name.startswith('_'):
            object.__setattr__(self, name, value)
            return

        lookup = type(self).__dict__.get('_attr_lookup') or getattr(type(self), '_attr_lookup', {})
        actual_key = lookup.get(name) or lookup.get(name.lower())
        if actual_key is None:
            pascal_name = _snake_to_pascal(name)
            actual_key = lookup.get(pascal_name) or lookup.get(pascal_name.lower())

        if actual_key is not None:
            descriptor = type(self)._all_properties.get(actual_key)
            if descriptor is not None:
                descriptor.__set__(self, value)
                return

        # Unknown property: set as regular attribute (backward compat)
        object.__setattr__(self, name, value)

    def create_itemset(self, key: str, setid: str) -> "ParameterSet":
        """
        Create a nested parameter set using CreateItemSet (pset-based) or direct access (HSet-based).

        Args:
            key: The parameter key for the nested set
            setid: The SetID for the nested parameter set

        Returns:
            ParameterSet instance wrapping the nested parameter set
        """
        if self._backend is None:
            raise RuntimeError("create_itemset(): no parameterset is bound. Call .bind() first.")

        # Handle pset-based backend
        if hasattr(self._backend, "create_itemset"):
            # PsetBackend - use CreateItemSet method
            nested_pset = self._backend.create_itemset(key, setid)
            return ParameterSet(nested_pset)

        # Handle HSet-based backend (legacy)
        elif isinstance(self._backend, HParamBackend):
            # Try to get nested parameter set from HSet structure
            try:
                nested_obj = self._backend.get(key)
                return ParameterSet(nested_obj)
            except KeyError:
                # Create new nested parameter set if possible
                raise NotImplementedError(f"Cannot create nested parameter set '{key}' with HSet backend")

        # Handle other backends
        else:
            try:
                nested_obj = self._backend.get(key)
                return ParameterSet(nested_obj)
            except KeyError:
                raise NotImplementedError(f"Cannot create nested parameter set '{key}' with {type(self._backend)} backend")

    def _sync_hset_global_state(self):
        """
        Synchronize staged changes with global HParameterSet state for HSet-based actions.

        The core issue: HSet-based actions use the GLOBAL HParameterSet state, but the
        simplified API creates LOCAL copies via HAction.GetDefault(). This method bridges
        that gap by copying staged values to the global HParameterSet using dotted key paths.
        """
        # Only apply to HParamBackend instances with App reference
        if not (
            self._raw is not None
            and isinstance(self._backend, HParamBackend)
            and self._app_instance
        ):
            return

        try:
            # Get the global HParameterSet
            global_hparam = self._app_instance.api.HParameterSet
            if global_hparam is None:
                return

            # Create a backend for the global HParameterSet
            global_backend = HParamBackend(global_hparam)

            # Determine the HParam node prefix based on the local parameter set type
            hparam_prefix = self._get_hparam_prefix()
            if not hparam_prefix:
                return

            # Sync all staged values to the global HParameterSet using dotted paths
            for property_key, value in self._staged.items():
                # Skip nested ParameterSet objects for now
                if isinstance(value, ParameterSet):
                    continue

                try:
                    # Create the full dotted path: "HFindReplace.FindString"
                    global_key = f"{hparam_prefix}.{property_key}"
                    global_backend.set(global_key, value)
                except Exception as e:
                    # Skip properties that can't be set
                    continue

        except Exception as e:
            # Log the error but don't fail the apply operation
            import logging

            logging.warning(
                f"ParameterSet: Failed to sync with global HParameterSet: {e}"
            )


    # ------ descriptor hooks (staged-aware) ------
    def _ps_get(self, desc: PropertyDescriptor):
        key = desc.key
        if key in self._deleted:
            return None
        if key in self._staged:
            return self._staged[key]

        # For pset backends, try to get live value first
        if isinstance(self._backend, PsetBackend):
            try:
                live_value = self._backend.get(key)
                # Update snapshot with live value
                self._snapshot[key] = live_value
                return live_value
            except Exception:
                # Fall back to snapshot if live read fails
                pass

        return self._snapshot.get(key, None)

    def _ps_set(self, desc: PropertyDescriptor, value: Any):
        key = desc.key
        self._deleted.discard(key)

        # For pset backends, apply immediately (no staging)
        if isinstance(self._backend, PsetBackend):
            try:
                self._backend.set(key, value)
                # Update snapshot to reflect the change
                self._snapshot[key] = value
                # Clear from staged since it's already applied
                self._staged.pop(key, None)
            except Exception as e:
                # If immediate set fails, fall back to staging
                self._staged[key] = value
        else:
            # For other backends (HSet, etc.), use staging
            self._staged[key] = value

        # If setting a nested PS directly, keep cache aligned
        if isinstance(value, ParameterSet):
            self._wrapper_cache[key] = value

    def _ps_del(self, desc: PropertyDescriptor):
        key = desc.key
        self._staged.pop(key, None)
        self._deleted.add(key)
        # Deleting clears any cached wrapper for that key
        self._wrapper_cache.pop(key, None)
        return True

    # ------ conveniences ------
    def __getitem__(self, name: str):
        return getattr(self, name)

    def __setitem__(self, name: str, value):
        setattr(self, name, value)

    def _get_value(self, name):
        """Legacy method - use backend instead."""
        return self._backend.get(name)

    def _set_value(self, name, value):
        """Legacy method - use backend instead."""
        return self._backend.set(name, value)

    def _del_value(self, name):
        """Legacy method - use backend instead."""
        if self._backend is None:
            return False
        return self._backend.delete(name)

    def update(self, data: Dict[str, Any]):
        """Stage multiple values by attribute name (not raw keys)."""
        for n, v in data.items():
            if n in self._property_registry:
                setattr(self, n, v)
        return self

    def to_dict(
        self, *, include_defaults: bool = True, only: Optional[Iterable[str]] = None
    ) -> Dict[str, Any]:
        names = list(only) if only is not None else list(self._property_registry.keys())
        out = {}
        for n in names:
            val = getattr(self, n)  # staged-aware
            # For nested ParameterSets, show their own dict for readability
            if isinstance(val, ParameterSet):
                out[n] = val.to_dict(include_defaults=include_defaults)
            else:
                if include_defaults:
                    out[n] = val
                else:
                    desc = self._property_registry[n]
                    is_staged = desc.key in self._staged or desc.key in self._deleted
                    if is_staged or (desc.default is None or val != desc.default):
                        out[n] = val
        return out

    def _missing_required(self):
        missing = []
        for name, desc in self._property_registry.items():
            if desc.required:
                val = getattr(self, name)  # staged-aware
                is_missing = val in (None, "", [], {}, ())
                # If nested PS is required, also ensure it has no missing requireds
                if not is_missing and isinstance(val, ParameterSet):
                    nested_missing = val._missing_required()
                    is_missing = len(nested_missing) > 0
                if is_missing:
                    missing.append(name)
        return missing

    def dirty(self) -> Dict[str, Any]:
        """Return staged changes as {attr_name: value}, excluding deletions."""
        rev = {d.key: name for name, d in self._property_registry.items()}
        pretty = {}
        for k, v in self._staged.items():
            name = rev.get(k, k)
            pretty[name] = v.to_dict() if isinstance(v, ParameterSet) else v
        return pretty

    def deleted(self) -> set[str]:
        """Return attribute names marked for deletion."""
        rev = {d.key: name for name, d in self._property_registry.items()}
        return {rev.get(k, k) for k in self._deleted if k in rev}

    def __repr__(self):
        """Return human-readable representation of ParameterSet with all properties."""
        return self._format_repr()

    def _format_repr(self, indent=0, max_depth=3):
        """
        Format ParameterSet for display with human-readable values.

        Args:
            indent: Current indentation level
            max_depth: Maximum nesting depth to show

        Returns:
            Formatted string representation
        """
        if indent >= max_depth:
            return f"<{self.__class__.__name__} ...>"

        prefix = "  " * indent
        lines = [f"{self.__class__.__name__}("]

        # Get all properties from the registry
        if hasattr(self, '_property_registry'):
            props = sorted(self._property_registry.items())
        else:
            props = []

        # Show each property with its value
        for prop_name, prop_descriptor in props:
            try:
                value = getattr(self, prop_name)

                # Format value based on type and property descriptor
                if value is None:
                    formatted_value = "None"
                elif isinstance(value, ParameterSet):
                    # Nested ParameterSet - show recursively
                    if indent + 1 < max_depth:
                        nested_repr = value._format_repr(indent + 1, max_depth)
                        formatted_value = f"\n{prefix}  {nested_repr}"
                    else:
                        formatted_value = f"<{value.__class__.__name__} ...>"
                elif isinstance(value, HArrayWrapper):
                    # Array - show as list
                    formatted_value = f"{value.to_list()}"
                elif isinstance(value, bool):
                    # Boolean - show as True/False
                    formatted_value = str(value)
                elif isinstance(value, int):
                    # Check if it's a color, size, or dimension based on property name/type
                    formatted_value = self._format_int_value(prop_name, prop_descriptor, value)
                elif isinstance(value, float):
                    formatted_value = f"{value}"
                elif isinstance(value, str):
                    # String - check if it's from MappedProperty (enum-like)
                    if type(prop_descriptor).__name__ == 'MappedProperty':
                        # Show both numeric value and mapped name
                        # Get raw numeric value from backend or staging
                        raw_value = None

                        # Try backend first (for bound ParameterSets)
                        if self._backend is not None:
                            try:
                                raw_value = self._backend.get(prop_descriptor.key)
                            except:
                                pass

                        # Try staging dict (for unbound ParameterSets)
                        if raw_value is None and hasattr(self, '_staged'):
                            raw_value = self._staged.get(prop_descriptor.key)

                        if raw_value is not None:
                            formatted_value = f"{raw_value} ({value})"
                        else:
                            formatted_value = f'"{value}"'
                    else:
                        # Regular string - show with quotes, limit length
                        if len(value) > 50:
                            formatted_value = f'"{value[:47]}..."'
                        else:
                            formatted_value = f'"{value}"'
                else:
                    # Other types
                    formatted_value = repr(value)

                # Add property description as comment if available
                doc_comment = ""
                if hasattr(prop_descriptor, 'doc') and prop_descriptor.doc:
                    doc_comment = f"        # {prop_descriptor.doc}"

                lines.append(f"{prefix}  {prop_name}={formatted_value}{doc_comment}")

            except (AttributeError, RuntimeError) as e:
                # Property couldn't be accessed (e.g., nested property without backend)
                lines.append(f"{prefix}  {prop_name}=<not accessible>")

        # Add status info
        if hasattr(self, 'dirty') and self.dirty():
            lines.append(f"{prefix}  [staged changes: {len(self._staged)}]")
        if hasattr(self, 'deleted') and self.deleted():
            lines.append(f"{prefix}  [deleted: {len(self._deleted)}]")

        lines.append(f"{prefix})")
        return "\n".join(lines)

    def _format_int_value(self, prop_name: str, prop_descriptor, value: int) -> str:
        """
        Format integer value based on property type.

        Args:
            prop_name: Name of the property
            prop_descriptor: Property descriptor instance
            value: Integer value to format

        Returns:
            Formatted string
        """
        from .functions import from_hwpunit

        # Check property descriptor type
        prop_type_name = type(prop_descriptor).__name__

        # Color properties - convert BBGGRR to #RRGGBB
        if prop_type_name == 'ColorProperty' or 'Color' in prop_name or 'Corlo' in prop_name:
            if value == 0:
                return "#000000 (black)"
            # Convert from BBGGRR to RRGGBB
            bb = (value >> 16) & 0xFF
            gg = (value >> 8) & 0xFF
            rr = value & 0xFF
            return f"#{rr:02X}{gg:02X}{bb:02X}"

        # Font size properties - convert HWPUNIT to pt
        # Check for Size properties (FontSize, Size, etc.)
        if 'Size' in prop_name or prop_name.endswith('Size'):
            # Font sizes are in HWPUNIT (100 HWPUNIT = 1pt)
            if value > 0 and value < 100000:  # Reasonable font size range
                pt_value = value / 100
                return f"{pt_value}pt"

        # Dimension properties - convert HWPUNIT to mm
        if (prop_type_name == 'UnitProperty' or
            'Width' in prop_name or 'Height' in prop_name or
            'Margin' in prop_name or 'Spacing' in prop_name or
            'Offset' in prop_name or 'Indent' in prop_name):
            # Convert HWPUNIT to mm (283 HWPUNIT = 1mm)
            if value != 0:
                try:
                    mm_value = from_hwpunit(value, 'mm')
                    return f"{mm_value}mm"
                except:
                    pass

        # Default - just show the number
        return str(value)

    def __str__(self):
        """Return simple string representation."""
        return f"<{self.__class__.__name__}>"
    @property
    def attributes_names(self):
        """Auto-generated list of attribute names from property registry."""
        return list(self._property_registry.keys())

    # ── Native pset methods (thin wrappers around COM API) ──────────────
    # Per HWP Automation docs (HwpAutomation_2504.pdf, p.63-65), pset COM
    # objects expose Clone, IsEquivalent, Merge, ItemExist, RemoveItem.

    def clone(self):
        """
        Clone this ParameterSet using the native COM Clone() method.

        Faster and more reliable than manual update_from for simple copies.

        Returns:
            New ParameterSet wrapping a cloned pset COM object, or None if
            the underlying object doesn't support Clone.
        """
        if self._raw is None or not hasattr(self._raw, 'Clone'):
            return None
        try:
            cloned_raw = self._raw.Clone()
        except Exception:
            return None
        return type(self)(cloned_raw)

    def is_equivalent(self, other) -> bool:
        """
        Check value equality with another ParameterSet via native IsEquivalent().

        Returns True if both parameter sets have identical items and values.
        """
        if self._raw is None or not hasattr(self._raw, 'IsEquivalent'):
            return False
        other_raw = other._raw if isinstance(other, ParameterSet) else other
        if other_raw is None:
            return False
        try:
            return bool(self._raw.IsEquivalent(other_raw))
        except Exception:
            return False

    def merge(self, other):
        """
        Merge another ParameterSet's items into this one via native Merge().

        Result: self gets "self's items + items unique to other".
        """
        if self._raw is None or not hasattr(self._raw, 'Merge'):
            return
        other_raw = other._raw if isinstance(other, ParameterSet) else other
        if other_raw is None:
            return
        try:
            self._raw.Merge(other_raw)
        except Exception:
            pass

    def item_exists(self, item_id: str) -> bool:
        """Check if an item exists in the underlying pset via native ItemExist()."""
        if self._raw is None or not hasattr(self._raw, 'ItemExist'):
            return False
        try:
            return bool(self._raw.ItemExist(item_id))
        except Exception:
            return False


# Additional methods for ParameterSet class
def update_from(self, pset):
    """
    Update this ParameterSet with values from another ParameterSet instance.

    Only attributes defined in self.attributes_names are updated.
    Nested ParameterSet attributes are updated recursively.
    If a value is None, the attribute is deleted.
    If a value is truthy, it is set on self.
    """
    # Handle both ParameterSet instances and raw COM objects
    if isinstance(pset, ParameterSet):
        # Standard ParameterSet - use getattr normally
        for key in self.attributes_names:
            value = getattr(pset, key, None)
            if isinstance(value, ParameterSet):
                target = getattr(self, key)
                if isinstance(target, ParameterSet):
                    target.update_from(value)
            elif value is None:
                self._del_value(key)
            elif value:
                try:
                    setattr(self, key, value)
                except (ValueError, TypeError) as e:
                    import logging
                    logging.warning(f"Skipping invalid value for '{key}': {value}. Error: {e}")
                    continue
    elif hasattr(pset, '_oleobj_') or str(type(pset)).find('com_gen_py') != -1:
        # Raw COM object - use COM property names
        for key in self.attributes_names:
            # Convert Python attribute name to COM property name
            com_key = self._python_to_com_key(key)
            if com_key and hasattr(pset, com_key):
                try:
                    value = getattr(pset, com_key)
                    if value is not None:
                        setattr(self, key, value)
                except (ValueError, TypeError, AttributeError) as e:
                    import logging
                    logging.warning(f"Skipping invalid COM value for '{key}' (COM: '{com_key}'): {e}")
                    continue
    else:
        raise TypeError("update_from expects a ParameterSet instance or COM object")
    return self

def _python_to_com_key(self, python_key):
    """
    Convert Python attribute name to COM property name.
    Default implementation converts snake_case to PascalCase.
    Subclasses can override this method for specific mappings.
    """
    # Default conversion: snake_case to PascalCase
    # e.g., "find_string" -> "FindString", "match_case" -> "MatchCase"
    parts = python_key.split('_')
    return ''.join(word.capitalize() for word in parts)

ParameterSet.update_from = update_from
ParameterSet._python_to_com_key = _python_to_com_key
ParameterSet.serialize = lambda self: self._serialize_impl()
ParameterSet.__str__ = lambda self: self._format_repr()  # Use same format as __repr__
# ParameterSet.__repr__ = lambda self: self.__str__()  # Removed - using custom __repr__ from ParameterSet class

def _update_from_impl(self, pset):
    """Update this parameter set with values from another parameter set."""
    for key in self.attributes_names:
        value = getattr(pset, key, None)

        if isinstance(value, ParameterSet):
            # Recursively update nested parameter sets
            target = getattr(self, key)
            target.update_from(value)
        elif value is None:
            # Remove the attribute if value is None
            self._del_value(key)
        elif value:
            # Set the attribute if value is truthy, but handle validation errors gracefully
            try:
                setattr(self, key, value)
            except (ValueError, TypeError) as e:
                # Log validation errors but continue with other attributes
                import logging
                logging.warning(f"Skipping invalid value for '{key}': {value}. Error: {e}")
                continue
    return self


def _serialize_impl(self, max_depth=3, _depth=0):
    """
    Robustly convert the parameter set to a dictionary, walking the COM tree and handling COM errors.
    """
    if _depth > max_depth:
        return "<max depth reached>"
    result = {}
    for key in self.attributes_names:
        try:
            value = getattr(self, key, None)
            if isinstance(value, ParameterSet):
                value = value.serialize(max_depth=max_depth, _depth=_depth+1)
            elif hasattr(value, "_oleobj_"):
                # Try to walk COM object attributes (HSet child)
                value = _safe_com_serialize(value, max_depth, _depth+1)
            result[key] = value
        except Exception as e:
            # Handle COM errors or any other error gracefully
            result[key] = f"<COM error: {e}>"
    return result

def _safe_com_serialize(obj, max_depth=3, _depth=0):
    """
    Recursively walk a COM object, returning a dict of its properties, handling COM errors.
    """
    if _depth > max_depth:
        return "<max depth reached>"
    result = {}
    for attr in dir(obj):
        if attr.startswith("_"):
            continue
        try:
            value = getattr(obj, attr)
            if hasattr(value, "_oleobj_"):
                value = _safe_com_serialize(value, max_depth, _depth+1)
            result[attr] = value
        except Exception as e:
            result[attr] = f"<COM error: {e}>"
    return result


def _str_impl(self):
    """String representation of the parameter set."""
    data = {
        "name": self.__class__.__name__,
        "values": self.serialize()
    }
    return pprint.pformat(data, indent=4, width=60)

# Attach the implementations to the class
ParameterSet._update_from_impl = _update_from_impl
ParameterSet._serialize_impl = _serialize_impl
ParameterSet._str_impl = _str_impl


# Static methods for creating properties
@staticmethod
def _typed_prop(key, doc, expected_type):
    """Create a property for typed parameter sets."""
    return TypedProperty(key, doc, expected_type)

@staticmethod
def _int_prop(key, doc, min_val=None, max_val=None):
    """Create a property for integer values with optional range validation."""
    return IntProperty(key, doc, min_val, max_val)

@staticmethod
def _bool_prop(key, doc):
    """Create a property for boolean values (0 or 1)."""
    return BoolProperty(key, doc)

@staticmethod
def _color_prop(key, doc):
    """Create a property for color values with hex conversion."""
    return ColorProperty(key, doc)

@staticmethod
def _unit_prop(key, unit, doc):
    """Create a property for unit-based values with automatic conversion."""
    return UnitProperty(key, unit, doc)

@staticmethod
def _mapped_prop(key, mapping, doc):
    """Create a property for mapped values (string <-> integer)."""
    return MappedProperty(key, mapping, doc)

@staticmethod
def _str_prop(key, doc):
    """Create a property for string values."""
    return StringProperty(key, doc)

@staticmethod
def _int_list_prop(key, doc):
    """Create a property for lists of integers."""
    return ListProperty(key, doc, item_type=int)

@staticmethod
def _tuple_list_prop(key, doc):
    """Create a property for lists of (X, Y) coordinate tuples."""
    return ListProperty(key, doc, item_type=tuple)

@staticmethod
def _gradation_color_prop(key, doc):
    """Create a property for gradation color lists with hex conversion."""
    class GradationColorProperty(ListProperty):
        def __get__(self, instance, owner):
            if instance is None:
                return self
            value = self._get_value(instance)
            if value is None:
                return None
            return [convert_hwp_color_to_hex(color) for color in value]

        def __set__(self, instance, value):
            if value is None:
                return self._del_value(instance)
            if not isinstance(value, (list, tuple)):
                raise TypeError(f"Value for '{self.key}' must be a list or tuple")

            converted_colors = [convert_to_hwp_color(color) for color in value]
            return self._set_value(instance, converted_colors)

    return GradationColorProperty(key, doc, item_type=int, min_length=2, max_length=10)

# Attach static methods to ParameterSet class
ParameterSet._typed_prop = _typed_prop
ParameterSet._int_prop = _int_prop
ParameterSet._bool_prop = _bool_prop
ParameterSet._color_prop = _color_prop
ParameterSet._unit_prop = _unit_prop
ParameterSet._mapped_prop = _mapped_prop
ParameterSet._str_prop = _str_prop
ParameterSet._int_list_prop = _int_list_prop
ParameterSet._tuple_list_prop = _tuple_list_prop
ParameterSet._gradation_color_prop = _gradation_color_prop


class GenericParameterSet(ParameterSet):
    """
    Generic wrapper for unknown ParameterSet types.

    Provides dynamic attribute access to pset objects that don't have
    a dedicated ParameterSet subclass defined.
    """

    def __init__(self, parameterset, pset_id=None):
        super().__init__(parameterset)
        self._pset_id = pset_id or "Unknown"

    def __getattr__(self, name):
        if name.startswith('_'):
            return super().__getattribute__(name)
        try:
            return self._backend.get(name)
        except:
            raise AttributeError(f"'{name}' not found in {self._pset_id}")

    def __setattr__(self, name, value):
        if name.startswith('_') or name in ('logger',):
            super().__setattr__(name, value)
        else:
            self._backend.set(name, value)

    def __repr__(self):
        return f"GenericParameterSet({self._pset_id})"


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
