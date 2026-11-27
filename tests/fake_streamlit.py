"""A tiny fake of the Streamlit runtime, enough to execute app.py script runs in tests.

It models the behaviour that matters for state bugs:
  * the whole script re-runs from the top on every interaction;
  * widget state survives only while the widget is rendered (state of widgets that were not
    rendered in a completed run is dropped);
  * a widget whose stored value is no longer a valid option falls back to its default;
  * a button is True only in the run triggered by its click;
  * st.stop() ends the run; st.rerun() abandons it and starts a new one.
Nothing here imports the real streamlit package.
"""
import hashlib
import inspect
import sys
from pathlib import Path
from types import SimpleNamespace

import pandas as pd


class StopApp(Exception):
    pass


class RerunApp(Exception):
    pass


class Upload:
    def __init__(self, name, data):
        self.name, self._data = name, data

    def getvalue(self):
        return self._data


class Ctx:
    """Stands in for columns, expanders, containers, popovers and spinners."""

    def __init__(self, st):
        self._st = st

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def __getattr__(self, name):
        return getattr(self._st, name)


def _make_text(kind):
    def fn(self, text="", *a, **k):
        self.texts.append((kind, str(text)))
    return fn


class FakeStreamlit:
    def __init__(self):
        self.session_state = {}
        self.widget_keys, self.clicks = set(), set()
        self.caches, self.cache_misses = {}, {}      # persist across runs, like st.cache_data
        self.column_config = SimpleNamespace(NumberColumn=lambda **k: ("number", k), CheckboxColumn=lambda **k: ("checkbox", k))
        self._begin_run()

    # ---- run bookkeeping
    def _begin_run(self):
        self.rendered, self.texts, self.tables, self.expanders, self.editors = set(), [], [], [], []
        self.expanded = {}
        self.buttons, self.downloads = {}, {}
        self.active_clicks, self.clicks = self.clicks, set()

    def _end_run(self):
        for key in list(self.widget_keys - self.rendered):
            self.session_state.pop(key, None)
            self.widget_keys.discard(key)

    def _widget(self, key, default, valid=None):
        self.rendered.add(key)
        self.widget_keys.add(key)
        if key in self.session_state and (valid is None or valid(self.session_state[key])):
            return self.session_state[key]
        self.session_state[key] = default
        return default

    # ---- no-ops and text
    def set_page_config(self, **k): pass
    def divider(self): pass

    def cache_data(self, *a, **k):
        """Caches by every argument whose name does not start with an underscore (as Streamlit does)."""
        def decorate(fn):
            store = self.caches.setdefault(fn.__qualname__, {})
            sig = inspect.signature(fn)

            def wrapper(*args, **kw):
                bound = sig.bind(*args, **kw)
                bound.apply_defaults()
                key = tuple((n, hashlib.md5(repr(v).encode()).hexdigest())
                            for n, v in bound.arguments.items() if not n.startswith("_"))
                if key not in store:
                    self.cache_misses[fn.__qualname__] = self.cache_misses.get(fn.__qualname__, 0) + 1
                    store[key] = fn(*args, **kw)
                return store[key]
            return wrapper
        return decorate(a[0]) if len(a) == 1 and callable(a[0]) else decorate

    markdown = _make_text("markdown")
    caption = _make_text("caption")
    error = _make_text("error")
    warning = _make_text("warning")
    info = _make_text("info")
    success = _make_text("success")

    # ---- layout
    def columns(self, spec, **k):
        return [Ctx(self) for _ in range(spec if isinstance(spec, int) else len(spec))]

    def expander(self, label, expanded=False):
        self.expanders.append(label)
        self.expanded[label] = expanded
        return Ctx(self)

    def tabs(self, labels):
        return [Ctx(self) for _ in labels]

    def container(self, **k): return Ctx(self)
    def popover(self, label, **k): return Ctx(self)
    def spinner(self, text=""): return Ctx(self)
    def empty(self): return Ctx(self)

    # ---- flow control
    def stop(self): raise StopApp()
    def rerun(self): raise RerunApp()
    experimental_rerun = rerun

    # ---- widgets
    def file_uploader(self, label, type=None, key=None, **k):
        return self._widget(key, None)

    def selectbox(self, label, options, index=0, key=None, **k):
        options = list(options)
        return self._widget(key, options[index], lambda v: v in options)

    def radio(self, label, options, key=None, **k):
        options = list(options)
        return self._widget(key, options[0], lambda v: v in options)

    def multiselect(self, label, options, default=None, key=None, **k):
        options = list(options)
        return list(self._widget(key, list(default or []), lambda v: all(x in options for x in v)))

    def number_input(self, label, min_value=None, value=0, key=None, **k):
        return self._widget(key, value)

    def slider(self, label, lo, hi, value, key=None, **k):
        return self._widget(key, value)

    def checkbox(self, label, key=None, **k):
        return self._widget(key, False)

    def button(self, label, key=None, type=None, disabled=False, **k):
        key = key or label
        self.buttons[key] = {"disabled": disabled}
        return key in self.active_clicks and not disabled

    def download_button(self, label, data, file_name=None, mime=None, **k):
        self.downloads[label] = {"data": data, "file_name": file_name}
        return False

    def data_editor(self, df, key=None, **k):
        self.editors.append(key)
        edits = self._widget(key, {})
        out = df.copy()
        checkbox = next(c for c in out.columns if out[c].dtype == bool)       # the one tick-box column
        for idx, value in edits.items():
            out.loc[idx, checkbox] = value
        return out

    def dataframe(self, df, **k):
        self.tables.append(df)


def _record_labels(cls):
    """Every widget label and help text counts as text the user can read."""
    import functools
    for name in ("file_uploader", "selectbox", "radio", "multiselect", "number_input", "slider",
                 "checkbox", "button", "download_button"):
        def make(orig):
            @functools.wraps(orig)
            def wrapper(self, label, *a, **k):
                self.texts.append(("label", str(label)))
                if k.get("help"):
                    self.texts.append(("help", str(k["help"])))
                return orig(self, label, *a, **k)
            return wrapper
        setattr(cls, name, make(getattr(cls, name)))
    return cls


FakeStreamlit = _record_labels(FakeStreamlit)


class AppDriver:
    """Runs app.py against the fake runtime and lets a test act like a user."""

    def __init__(self, path):
        self.path = Path(path)
        self.code = compile(self.path.read_text(), str(self.path), "exec")
        self.st = FakeStreamlit()

    @property
    def state(self):
        return self.st.session_state

    def run(self):
        saved = sys.modules.get("streamlit")
        sys.modules["streamlit"] = self.st
        try:
            for _ in range(25):
                self.st._begin_run()
                try:
                    exec(self.code, {"__name__": "__main__", "__file__": str(self.path)})
                except RerunApp:
                    continue
                except StopApp:
                    pass
                self.st._end_run()
                return self
            raise RuntimeError("the app kept rerunning")
        finally:
            if saved is None:
                sys.modules.pop("streamlit", None)
            else:
                sys.modules["streamlit"] = saved

    # ---- user actions (each is followed by run())
    def upload(self, key, name, data):
        self.st.session_state[key] = Upload(name, data)
        self.st.widget_keys.add(key)
        return self

    def clear_upload(self, key):
        self.st.session_state[key] = None
        return self

    def set(self, key, value):
        assert key in self.st.rendered, f"widget {key!r} is not on the page"
        self.st.session_state[key] = value
        return self

    def click(self, label):
        assert label in self.st.buttons, f"no button {label!r}; have {sorted(self.st.buttons)}"
        assert not self.st.buttons[label]["disabled"], f"button {label!r} is disabled"
        self.st.clicks.add(label)
        return self

    def edit(self, key, edits):
        assert key in self.st.editors, f"no editor {key!r}; have {self.st.editors}"
        self.st.session_state[key] = {**self.st.session_state.get(key, {}), **edits}
        return self

    # ---- what the user sees
    def editor_key(self, sid):
        return next(k for k in self.st.editors if k.startswith(f"{sid}_review_"))

    def text(self):
        return "\n".join(t for _, t in self.st.texts)

    def has_text(self, fragment):
        return fragment in self.text()

    def table(self, *columns):
        found = [t for t in self.st.tables if set(columns) <= set(t.columns)]
        assert found, f"no table with columns {columns}"
        return found[-1]

    def button_enabled(self, label):
        return label in self.st.buttons and not self.st.buttons[label]["disabled"]