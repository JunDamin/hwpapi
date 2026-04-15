"""Domain-grouped tests: deprecations."""
from hwpapi.core.app import _warn_legacy
from unittest.mock import MagicMock
import pytest
import warnings


def _mock_app():
    """Create an App-like mock with api that supports field calls."""
    from hwpapi.core.app import App
    app = App.__new__(App)
    # api is a @property returning self.engine.impl — stub via a fake engine
    fake_api = MagicMock()
    fake_api.CreateField.return_value = 1
    fake_api.PutFieldText.return_value = 1
    fake_api.GetFieldText.return_value = "value"
    fake_api.FieldExist.return_value = True
    fake_api.MoveToField.return_value = True
    fake_api.RenameField.return_value = True
    fake_api.GetFieldList.return_value = "a\x02b\x02"
    fake_api.Run.return_value = True
    fake_engine = MagicMock()
    fake_engine.impl = fake_api
    app.engine = fake_engine
    app.logger = MagicMock()
    return app


def test_warn_legacy_emits_deprecation():
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        _warn_legacy("old_api", "new_api")
    assert any(issubclass(x.category, DeprecationWarning) for x in w)
    msg = str(w[-1].message)
    assert "old_api" in msg
    assert "new_api" in msg
    assert "v1.0" in msg


def test_warn_legacy_silent_from_hwpapi_module():
    """
    Simulates: hwpapi.classes.X calls App.create_field → _warn_legacy.

    Stack at _warn_legacy:
      frame(0) = _warn_legacy
      frame(1) = App.create_field (hwpapi.core.app)
      frame(2) = hwpapi.fake_accessor (simulated)  ← this is "user" of create_field

    Since frame(2) is in hwpapi.*, should suppress.
    """
    import sys
    import types

    # Build a fake hwpapi module that wraps a legacy method call
    wrapper_src = """
def hwpapi_internal_caller(app):
    # Pretend to be hwpapi internal code calling a legacy App method
    from hwpapi.core.app import App
    App.create_field(app, "x")
"""
    mod = types.ModuleType("hwpapi.fake_accessor")
    mod.__name__ = "hwpapi.fake_accessor"
    exec(wrapper_src, mod.__dict__)
    sys.modules["hwpapi.fake_accessor"] = mod
    try:
        app = _mock_app()
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            mod.hwpapi_internal_caller(app)
        deprecations = [x for x in w if issubclass(x.category, DeprecationWarning)]
        assert not deprecations, (
            f"hwpapi.* caller of legacy method should suppress warnings, got: "
            f"{[str(x.message) for x in deprecations]}"
        )
    finally:
        del sys.modules["hwpapi.fake_accessor"]


def test_create_field_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.create_field(app, "x")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("create_field" in m and "app.fields.add" in m for m in msgs)


def test_set_field_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.set_field(app, "x", "v")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("set_field" in m for m in msgs)


def test_get_field_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.get_field(app, "x")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("get_field" in m for m in msgs)


def test_field_exists_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.field_exists(app, "x")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("field_exists" in m for m in msgs)


def test_move_to_field_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.move_to_field(app, "x")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("move_to_field" in m for m in msgs)


def test_delete_field_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.delete_field(app, "x")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("delete_field" in m for m in msgs)


def test_rename_field_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.rename_field(app, "a", "b")
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("rename_field" in m for m in msgs)


def test_field_names_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        # Access property via class-level descriptor
        App.field_names.fget(app)
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("field_names" in m for m in msgs)


def test_fields_dict_warns():
    from hwpapi.core.app import App
    app = _mock_app()
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        App.fields_dict.fget(app)
    msgs = [str(x.message) for x in w if issubclass(x.category, DeprecationWarning)]
    assert any("fields_dict" in m for m in msgs)


def test_fields_accessor_silent():
    """When the new Fields accessor is used, no DeprecationWarning should fire."""
    from hwpapi.classes.fields import Fields
    app = _mock_app()
    f = Fields(app)
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        # Various operations — these internally call legacy methods but
        # should be silent because they originate from hwpapi.classes.fields
        _ = "x" in f
        _ = list(f)
        f["new"] = "value"      # calls create_field + set_field internally
        _ = f.to_dict()
        f.remove("x")
        f.rename("old", "new")
    deprecations = [x for x in w if issubclass(x.category, DeprecationWarning)]
    assert not deprecations, f"Unexpected deprecations: {[str(x.message) for x in deprecations]}"


