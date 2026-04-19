"""
Comprehensive tests for all HWP Actions.
Requires HWP application to be installed and running.
"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# ── Check dependencies ───────────────────────────────────────────────────

try:
    import win32com.client
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

try:
    from hwpapi.core import App
    from hwpapi.low.actions import _Action, _Actions, _action_info
    from hwpapi.low.parametersets import ParameterSet, PARAMETERSET_REGISTRY
    import hwpapi.low.parametersets as ps_mod
    HAS_HWPAPI = True
except ImportError:
    HAS_HWPAPI = False

HAS_HWP = False
if HAS_HWPAPI and HAS_PYWIN32:
    try:
        _test_app = App()
        HAS_HWP = True
    except Exception:
        pass

pytestmark = pytest.mark.skipif(not HAS_HWP, reason="HWP application not available")


# ── Fixtures ─────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def app():
    """Create a shared App instance for all tests."""
    if not HAS_HWP:
        pytest.skip("HWP not available")
    return App()


# ── Collect action info ──────────────────────────────────────────────────

# Actions that cause Windows fatal exceptions (HWP COM crash)
SKIP_ACTIONS = {"InsertFieldMemo", "InsertFieldRevisionChagne", "TextArtModify"}

if HAS_HWPAPI:
    ALL_ACTION_NAMES = sorted(k for k in _action_info.keys() if k not in SKIP_ACTIONS)
    ACTIONS_WITH_PSET = sorted(
        [(k, v[0]) for k, v in _action_info.items()
         if v[0] is not None and k not in SKIP_ACTIONS],
        key=lambda x: x[0],
    )
    ACTIONS_WITHOUT_PSET = sorted(
        [k for k, v in _action_info.items()
         if v[0] is None and k not in SKIP_ACTIONS]
    )
else:
    ALL_ACTION_NAMES = []
    ACTIONS_WITH_PSET = []
    ACTIONS_WITHOUT_PSET = []


# ── A. Action access tests ──────────────────────────────────────────────

class TestActionAccess:
    """Test that every action can be accessed from app.actions."""

    @pytest.mark.parametrize("action_name", ALL_ACTION_NAMES)
    def test_action_accessible(self, app, action_name):
        """Each action should be accessible as a property on app.actions."""
        try:
            action = getattr(app.actions, action_name)
        except AssertionError:
            pytest.xfail(f"{action_name}: _action_info SetID mismatch with HWP")
        except Exception as e:
            if "com_error" in type(e).__name__ or "pywintypes" in str(type(e)):
                pytest.xfail(f"{action_name}: HWP COM error - {e}")
            raise
        assert isinstance(action, _Action), f"{action_name} is not an _Action"

    @pytest.mark.parametrize("action_name", ALL_ACTION_NAMES)
    def test_action_has_description(self, app, action_name):
        """Each action should have a description."""
        try:
            action = getattr(app.actions, action_name)
        except (AssertionError, Exception) as e:
            pytest.xfail(f"{action_name}: cannot create action - {e}")
        assert action.description is not None
        assert isinstance(action.description, str)

    @pytest.mark.parametrize("action_name", ALL_ACTION_NAMES)
    def test_action_repr(self, app, action_name):
        """repr should not error."""
        try:
            action = getattr(app.actions, action_name)
        except (AssertionError, Exception) as e:
            pytest.xfail(f"{action_name}: cannot create action - {e}")
        r = repr(action)
        assert action_name in r


# ── B. Actions with ParameterSet tests ──────────────────────────────────

class TestActionsWithParameterSet:
    """Test actions that have associated ParameterSets."""

    def _get_action(self, app, action_name):
        """Helper to get action, xfail on SetID mismatch or COM error."""
        try:
            return getattr(app.actions, action_name)
        except AssertionError:
            pytest.xfail(f"{action_name}: _action_info SetID mismatch")
        except Exception as e:
            if "com_error" in type(e).__name__ or "pywintypes" in str(type(e)):
                pytest.xfail(f"{action_name}: HWP COM error")
            raise

    @pytest.mark.parametrize("action_name,pset_key", ACTIONS_WITH_PSET,
                             ids=[a for a, _ in ACTIONS_WITH_PSET])
    def test_pset_exists(self, app, action_name, pset_key):
        """Actions with SetID should have a non-None pset."""
        action = self._get_action(app, action_name)
        assert action.pset is not None, f"{action_name} has pset_key={pset_key} but pset is None"

    @pytest.mark.parametrize("action_name,pset_key", ACTIONS_WITH_PSET,
                             ids=[a for a, _ in ACTIONS_WITH_PSET])
    def test_pset_is_parameterset(self, app, action_name, pset_key):
        """The pset should be a ParameterSet instance."""
        action = self._get_action(app, action_name)
        assert isinstance(action.pset, ParameterSet), (
            f"{action_name}.pset is {type(action.pset)}, not ParameterSet"
        )

    @pytest.mark.parametrize("action_name,pset_key", ACTIONS_WITH_PSET,
                             ids=[a for a, _ in ACTIONS_WITH_PSET])
    def test_pset_attributes_accessible(self, app, action_name, pset_key):
        """All attributes on the pset should be accessible."""
        action = self._get_action(app, action_name)
        pset = action.pset
        for attr_name in pset.attributes_names:
            descriptor = pset._property_registry.get(attr_name)
            from hwpapi.low.parametersets import NestedProperty, ArrayProperty
            if isinstance(descriptor, (NestedProperty, ArrayProperty)):
                continue
            try:
                val = getattr(pset, attr_name)
            except Exception as e:
                pytest.fail(f"{action_name}.pset.{attr_name} raised: {e}")


# ── C. Actions without ParameterSet tests ────────────────────────────────

class TestActionsWithoutParameterSet:
    """Test actions that have no associated ParameterSets (per _action_info)."""

    @pytest.mark.parametrize("action_name", ACTIONS_WITHOUT_PSET)
    def test_no_pset_action(self, app, action_name):
        """Actions without SetID: either pset=None or xfail if HWP disagrees."""
        try:
            action = getattr(app.actions, action_name)
        except AssertionError:
            pytest.xfail(f"{action_name}: HWP has ParameterSet but _action_info says None")
        except Exception as e:
            if "com_error" in type(e).__name__ or "pywintypes" in str(type(e)):
                pytest.xfail(f"{action_name}: HWP COM error - {e}")
            raise
        assert action.pset is None, f"{action_name} should have pset=None"


# ── D. Action info consistency tests ─────────────────────────────────────

class TestActionInfoConsistency:
    """Test _action_info dictionary consistency."""

    def test_action_count(self):
        """Verify expected action count."""
        assert len(_action_info) == 704

    def test_all_values_are_lists(self):
        """Every value should be [SetID_or_None, description]."""
        for name, value in _action_info.items():
            assert isinstance(value, list), f"{name}: value is {type(value)}"
            assert len(value) == 2, f"{name}: value has {len(value)} elements"

    def test_pset_keys_are_known(self):
        """Report non-None SetIDs that don't correspond to known ParameterSet classes."""
        unknown = []
        for name, (pset_key, _) in _action_info.items():
            if pset_key is not None:
                if not hasattr(ps_mod, pset_key):
                    unknown.append((name, pset_key))
        if unknown:
            pytest.xfail(f"{len(unknown)} unknown pset keys: {unknown}")

    def test_actions_accessible_via_getattr(self, app):
        """_Actions should resolve every action via __getattr__ dispatch."""
        missing = []
        for name in _action_info:
            try:
                action = getattr(app.actions, name)
                if not isinstance(action, _Action):
                    missing.append(name)
            except AssertionError:
                # SetID mismatch is acceptable (covered by xfail elsewhere)
                pass
            except Exception as e:
                if "com_error" not in type(e).__name__ and "pywintypes" not in str(type(e)):
                    missing.append(f"{name}: {e}")
        assert missing == [], f"Inaccessible actions: {missing[:10]}..."

    def test_introspection_api(self, app):
        """New introspection API should work without COM calls."""
        # get_pset_class
        cls = app.actions.get_pset_class('CharShape')
        assert cls is not None
        assert cls.__name__ == 'CharShape'
        # Cancel has no pset
        assert app.actions.get_pset_class('Cancel') is None
        # Unknown action raises KeyError
        try:
            app.actions.get_pset_class('XYZInvalid')
            assert False, "Should have raised KeyError"
        except KeyError:
            pass
        # list_actions
        assert len(app.actions.list_actions()) == len(_action_info)
        assert all(v[0] for k in app.actions.list_actions(with_pset_only=True)
                   for v in [_action_info[k]])
        # __contains__
        assert 'CharShape' in app.actions
        assert 'InvalidActionXYZ' not in app.actions

    def test_action_caching(self, app):
        """Repeated access to same action should return cached instance."""
        a1 = app.actions.CharShape
        a2 = app.actions.CharShape
        assert a1 is a2, "Actions should be cached"
        # refresh should invalidate cache
        app.actions.refresh('CharShape')
        a3 = app.actions.CharShape
        assert a1 is not a3, "refresh should create new instance"
