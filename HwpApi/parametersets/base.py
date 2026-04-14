"""
ParameterSet base class, metaclass, and GenericParameterSet.

This module defines the foundational classes for the ParameterSet system:

- ParameterSetMeta: Metaclass that auto-registers each subclass into
  ``PARAMETERSET_REGISTRY`` and builds a case-insensitive ``_attr_lookup``
  for O(1) snake_case ↔ PascalCase resolution.

- ParameterSet: Base class for all typed parameter wrappers. Supports:
  - snake_case and PascalCase attribute access
  - staged writes (``_staged``) and snapshot reads (``_snapshot``)
  - auto-wrapping of nested ParameterSets
  - native COM methods: ``clone()``, ``is_equivalent()``, ``merge()``, ``item_exists()``

- GenericParameterSet: Fallback wrapper for psets without a dedicated subclass.

- Static helpers attached to ParameterSet: ``update_from``, ``serialize``,
  ``__str__``, and static property factory methods (``_typed_prop`` etc.).
"""
from __future__ import annotations
import pprint
import re
from typing import Any, Dict, List, Optional, Union, Callable, Type, Iterable

from hwpapi.functions import from_hwpunit, to_hwpunit, convert_hwp_color_to_hex, convert_to_hwp_color
from .backends import (
    ParameterBackend, PsetBackend, HParamBackend, ComBackend, AttrBackend,
    _is_com, _looks_like_pset, make_backend,
)
from .properties import (
    MissingRequiredError, PropertyDescriptor, IntProperty, BoolProperty,
    StringProperty, ColorProperty, UnitProperty, MappedProperty, TypedProperty,
    NestedProperty, ArrayProperty, HArrayWrapper, ListProperty,
)


def _pascal_to_snake(name: str) -> str:
    """Convert PascalCase to snake_case."""
    snake = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', name)
    snake = re.sub('([a-z0-9])([A-Z])', r'\1_\2', snake)
    return snake.lower()


def _snake_to_pascal(name: str) -> str:
    """Convert snake_case to PascalCase (simple)."""
    parts = name.split('_')
    return ''.join(word.capitalize() for word in parts)


# PARAMETERSET_REGISTRY is shared with __init__.py via import
# It's populated by ParameterSetMeta.__new__ as each subclass is defined.
from hwpapi.parametersets import PARAMETERSET_REGISTRY


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


