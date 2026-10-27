from __future__ import annotations

import json
import os
import queue
import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple


def _ensure_tk_runtime_paths() -> None:
    """Set Tcl/Tk library paths when Python installation does not expose them."""
    if os.environ.get("TCL_LIBRARY") and os.environ.get("TK_LIBRARY"):
        return

    base_prefix = Path(os.path.dirname(os.__file__)).parent
    tcl_root = base_prefix / "tcl"
    tcl_dir = tcl_root / "tcl8.6"
    tk_dir = tcl_root / "tk8.6"

    if (tcl_dir / "init.tcl").exists() and (tk_dir / "tk.tcl").exists():
        os.environ.setdefault("TCL_LIBRARY", str(tcl_dir))
        os.environ.setdefault("TK_LIBRARY", str(tk_dir))


_ensure_tk_runtime_paths()

try:
    import pythoncom  # type: ignore
except ImportError:
    pythoncom = None


def _import_win32_client():
    try:
        import win32com.client  # type: ignore

        return win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for CANoe COM access. Install it with: python -m pip install pywin32"
        ) from exc


class CANoeController:
    def __init__(self, visible: bool = True) -> None:
        win32_client = _import_win32_client()
        self.app = win32_client.Dispatch("CANoe.Application")
        self.app.Visible = bool(visible)

    @property
    def measurement_running(self) -> bool:
        return bool(self.app.Measurement.Running)

    @staticmethod
    def _paths_equal(left: Path, right: Path) -> bool:
        try:
            return str(left.resolve()).lower() == str(right.resolve()).lower()
        except Exception:
            return str(left).lower() == str(right).lower()

    def _get_loaded_cfg_path(self) -> Optional[Path]:
        try:
            full_name = self._normalize_existing_path(self.app.Configuration.FullName)
        except Exception:
            return None

        if not full_name:
            return None

        try:
            return Path(full_name).expanduser().resolve()
        except Exception:
            return Path(full_name)

    def load_cfg(self, cfg_path: Path, force_reload: bool = False) -> None:
        cfg_path = cfg_path.expanduser().resolve()
        if not cfg_path.exists():
            raise FileNotFoundError("Configuration file not found: {0}".format(cfg_path))

        loaded_cfg = self._get_loaded_cfg_path()
        if (not force_reload) and loaded_cfg is not None and self._paths_equal(loaded_cfg, cfg_path):
            return

        if self.measurement_running:
            self.stop(timeout_s=15.0)

        try:
            self.app.Open(str(cfg_path))
            return
        except Exception as primary_exc:
            # Retry with explicit autosave/prompt flags to bypass blocking UI prompts.
            try:
                self.app.Open(str(cfg_path), False, False)
                return
            except Exception as retry_exc:
                details_primary = self._format_exception(primary_exc)
                details_retry = self._format_exception(retry_exc)
                raise RuntimeError(
                    "Failed to load cfg '{0}'. First Open error: {1}. Retry error: {2}".format(
                        cfg_path, details_primary, details_retry
                    )
                ) from retry_exc

    def start(self, timeout_s: float = 15.0) -> None:
        if self.measurement_running:
            return
        self.app.Measurement.Start()
        self._wait_measurement_state(expected_running=True, timeout_s=timeout_s)

    def stop(self, timeout_s: float = 15.0) -> None:
        if not self.measurement_running:
            return
        self.app.Measurement.Stop()
        self._wait_measurement_state(expected_running=False, timeout_s=timeout_s)

    def _wait_measurement_state(self, expected_running: bool, timeout_s: float) -> None:
        deadline = time.time() + max(timeout_s, 0.1)
        while time.time() < deadline:
            if self.measurement_running == expected_running:
                return
            time.sleep(0.1)
        expected = "running" if expected_running else "stopped"
        raise TimeoutError("CANoe measurement did not become {0} within {1:.1f}s".format(expected, timeout_s))

    @staticmethod
    def _normalize_existing_path(path_value: object) -> str:
        return str(path_value or "").strip().strip('"')

    @staticmethod
    def _format_exception(exc: Exception) -> str:
        text = str(exc).strip()
        if text:
            return text
        if getattr(exc, "args", None):
            return repr(exc.args)
        return exc.__class__.__name__

    def _get_replay_collection_count(self) -> Optional[int]:
        try:
            collection = self.app.Configuration.SimulationSetup.Buses(1).ReplayCollection
            count = int(collection.Count)
            return count
        except Exception:
            return None

    def _get_replay_block(self, index: int):
        try:
            return self.app.Configuration.SimulationSetup.Buses(1).ReplayCollection.Item(index)
        except Exception as exc:
            count = self._get_replay_collection_count()
            if count is None:
                details = self._format_exception(exc)
                raise RuntimeError(
                    "Unable to access Replay block #{0}. CANoe reported: {1}".format(index, details)
                ) from exc
            raise RuntimeError(
                "Replay block #{0} is not available in this CANoe cfg. Available blocks: {1}".format(index, count)
            ) from exc

    def set_replay_block_path(self, index: int, log_path: Path) -> None:
        resolved = log_path.expanduser().resolve()
        if not resolved.exists():
            raise FileNotFoundError("Replay log not found: {0}".format(resolved))
        replay_block = self._get_replay_block(index)
        replay_block.Path = str(resolved)

    def addReplayBlock_Customer(self, log_path: Path, block_index: int = 3) -> None:
        self.set_replay_block_path(block_index, log_path)

    def addReplayBlock_DiagRequest(self, log_path: Path, block_index: int = 1) -> None:
        self.set_replay_block_path(block_index, log_path)

    def apply_replay_pair(
        self,
        customer_log_path: Path,
        diag_request_log_path: Path,
        customer_block_index: int = 3,
        diag_request_block_index: int = 1,
    ) -> None:
        self.addReplayBlock_Customer(customer_log_path, block_index=customer_block_index)
        self.addReplayBlock_DiagRequest(diag_request_log_path, block_index=diag_request_block_index)

    @staticmethod
    def _build_diag_tester_name(customer_path: Path) -> str:
        source_name = customer_path.name
        replaced = re.sub(r"^diag_customer", "Diag_tester", source_name, flags=re.IGNORECASE)
        if replaced != source_name:
            return replaced
        return "Diag_tester_{0}".format(source_name)

    def Logging_rename(self, customer_log_path: Path) -> str:
        target_name = self._build_diag_tester_name(customer_log_path)
        try:
            logging_collection = self.app.Configuration.OnlineSetup.LoggingCollection
            logging_block = logging_collection.Item(1)
            logging_block.FullName = target_name
        except Exception as exc:
            details = self._format_exception(exc)
            raise RuntimeError(
                "Failed to rename logging block #1 to '{0}'. CANoe reported: {1}".format(target_name, details)
            ) from exc
        return target_name

    def _get_system_namespace(self, namespace_name: str):
        if not namespace_name or not namespace_name.strip():
            raise ValueError("namespace_name must not be empty")

        clean_name = namespace_name.strip()
        namespaces = self.app.System.Namespaces

        # Support both COM access styles across CANoe/pywin32 variants.
        try:
            return namespaces(clean_name)
        except Exception:
            try:
                return namespaces.Item(clean_name)
            except Exception as exc:
                details = self._format_exception(exc)
                raise RuntimeError(
                    "System namespace '{0}' not found or inaccessible. CANoe reported: {1}".format(
                        clean_name, details
                    )
                ) from exc

    def read_system_variable(self, namespace_name: str, variable_name: str) -> Any:
        """Read a CANoe system variable value.

        Example:
            value = controller.read_system_variable("MyNamespace", "MyVariable")
        """
        if not variable_name or not variable_name.strip():
            raise ValueError("variable_name must not be empty")

        clean_var_name = variable_name.strip()
        namespace = self._get_system_namespace(namespace_name)
        variables = namespace.Variables

        try:
            variable = variables(clean_var_name)
        except Exception:
            try:
                variable = variables.Item(clean_var_name)
            except Exception as exc:
                details = self._format_exception(exc)
                raise RuntimeError(
                    "System variable '{0}.{1}' not found or inaccessible. CANoe reported: {2}".format(
                        namespace_name.strip(), clean_var_name, details
                    )
                ) from exc

        try:
            return variable.Value
        except Exception as exc:
            details = self._format_exception(exc)
            raise RuntimeError(
                "Failed to read system variable '{0}.{1}'. CANoe reported: {2}".format(
                    namespace_name.strip(), clean_var_name, details
                )
            ) from exc

    def read_environment_variable(self, variable_name: str) -> Any:
        """Read a CANoe environment variable value.

        Example:
            value = controller.read_environment_variable("EnvDiagSession")
        """
        if not variable_name or not variable_name.strip():
            raise ValueError("variable_name must not be empty")

        clean_var_name = variable_name.strip()
        env = self.app.Environment
        variable_ref = None

        # Support common CANoe COM access styles.
        try:
            variable_ref = env.GetVariable(clean_var_name)
        except Exception:
            try:
                variable_ref = env.GetVariableByName(clean_var_name)
            except Exception:
                try:
                    variables = env.Variables
                    try:
                        variable_ref = variables(clean_var_name)
                    except Exception:
                        variable_ref = variables.Item(clean_var_name)
                except Exception as exc:
                    details = self._format_exception(exc)
                    raise RuntimeError(
                        "Environment variable '{0}' not found or inaccessible. CANoe reported: {1}".format(
                            clean_var_name, details
                        )
                    ) from exc

        # Most COM styles return a variable object with .Value; some may return the value directly.
        try:
            return variable_ref.Value
        except Exception:
            return variable_ref


CUSTOMER_PREFIX = "diag_customer"
REQUEST_PREFIX = "diag_request"
CONTROL_TIMEOUT_S = 15.0
DELTA_THRESHOLD_S = 0.15
MAX_RESET_EVENTS = 64
VSA_2FF_PATTERN = re.compile(r"\bVSA_2FF\b")
TIMESTAMP_PATTERN = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\b")


def extract_vsa_timestamps(log_path: Path) -> List[float]:
    timestamps = []  # type: List[float]
    with log_path.open("r", encoding="utf-8", errors="ignore") as handle:
        for line in handle:
            if not VSA_2FF_PATTERN.search(line):
                continue
            match = TIMESTAMP_PATTERN.match(line)
            if match is None:
                continue
            try:
                timestamps.append(float(match.group(1)))
            except ValueError:
                continue
    return timestamps


def find_reset_events(timestamps: List[float]) -> List[Tuple[float, float]]:
    events = []  # type: List[Tuple[float, float]]
    for idx in range(len(timestamps) - 1):
        off_time = timestamps[idx]
        on_time = timestamps[idx + 1]
        if (on_time - off_time) > DELTA_THRESHOLD_S:
            events.append((off_time, on_time))
    return events


def build_replay_delays_ms(events: List[Tuple[float, float]]) -> Tuple[List[float], List[float]]:
    off_delays_ms = []  # type: List[float]
    on_delays_ms = []  # type: List[float]
    for idx, (off_time, on_time) in enumerate(events):
        if idx == 0:
            off_wait_s = off_time
        else:
            off_wait_s = off_time - events[idx - 1][1]
        on_wait_s = on_time - off_time
        off_delays_ms.append(off_wait_s * 1000.0)
        on_delays_ms.append(on_wait_s * 1000.0)
    return off_delays_ms, on_delays_ms


def summarize_reset_profile(off_delays_ms: List[float], on_delays_ms: List[float], max_events: int = 3) -> str:
    if not off_delays_ms:
        return "-"

    pieces = []
    for idx, (off_ms, on_ms) in enumerate(zip(off_delays_ms, on_delays_ms), start=1):
        if idx > max_events:
            break
        pieces.append("{0}:{1:.0f}/{2:.0f}ms".format(idx, off_ms, on_ms))

    remaining = len(off_delays_ms) - min(len(off_delays_ms), max_events)
    if remaining > 0:
        pieces.append("+{0}".format(remaining))

    return " | ".join(pieces)


class CaplResetProfileUpdater:
    def __init__(self, can_file_path: Path) -> None:
        self.can_file_path = can_file_path

    @staticmethod
    def _format_array(values_ms: List[float]) -> str:
        if not values_ms:
            return "{0.0}"
        return "{" + ", ".join("{0:.3f}".format(value) for value in values_ms) + "}"

    def update_profile(self, off_delays_ms: List[float], on_delays_ms: List[float], source_name: str) -> int:
        if len(off_delays_ms) != len(on_delays_ms):
            raise ValueError("Invalid reset profile: OFF/ON lengths differ")
        if len(off_delays_ms) > MAX_RESET_EVENTS:
            raise ValueError(
                "Too many reset events ({0}). MAX_RESET_EVENTS is {1}.".format(
                    len(off_delays_ms), MAX_RESET_EVENTS
                )
            )

        original = self.can_file_path.read_text(encoding="cp1252", errors="ignore")
        updated = original

        event_count = len(off_delays_ms)
        off_array_text = self._format_array(off_delays_ms)
        on_array_text = self._format_array(on_delays_ms)

        count_pattern = re.compile(r"(?m)^(\s*int\s+igResetEventCount\s*=\s*)\d+(\s*;)")
        off_pattern = re.compile(r"(?m)^(\s*float\s+igOffDelayMs\s*\[\s*64\s*\]\s*=\s*)\{[^}]*\}(\s*;)")
        on_pattern = re.compile(r"(?m)^(\s*float\s+igOnDelayMs\s*\[\s*64\s*\]\s*=\s*)\{[^}]*\}(\s*;)")
        src_pattern = re.compile(r"(?m)^\s*//\s*IG_RESET_PROFILE_SOURCE:.*$")

        updated, count_repl = count_pattern.subn(r"\g<1>{0}\g<2>".format(event_count), updated, count=1)
        updated, off_repl = off_pattern.subn(r"\g<1>{0}\g<2>".format(off_array_text), updated, count=1)
        updated, on_repl = on_pattern.subn(r"\g<1>{0}\g<2>".format(on_array_text), updated, count=1)

        if count_repl != 1 or off_repl != 1 or on_repl != 1:
            raise RuntimeError(
                "Cannot update CAPL reset profile. Missing igResetEventCount/igOffDelayMs/igOnDelayMs declarations."
            )

        source_line = "// IG_RESET_PROFILE_SOURCE: {0} ({1} event(s))".format(source_name, event_count)
        if src_pattern.search(updated):
            updated = src_pattern.sub(source_line, updated, count=1)
        else:
            anchor = "float       igOnDelayMs[64]"
            anchor_idx = updated.find(anchor)
            if anchor_idx >= 0:
                line_end = updated.find("\n", anchor_idx)
                if line_end >= 0:
                    updated = updated[: line_end + 1] + "    " + source_line + "\n" + updated[line_end + 1 :]

        if updated != original:
            self.can_file_path.write_text(updated, encoding="cp1252")

        return event_count


class CanoeReplayPairGui:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("CANoe ReSIM")
        self.root.geometry("1180x700")

        appdata_root = Path(os.environ.get("APPDATA", str(Path.home())))
        self.settings_path = appdata_root / "CanoeReplayPairTool" / "settings.json"

        self.ui_queue = queue.Queue()
        self.worker_thread = None  # type: Optional[threading.Thread]
        self.auto_stop_thread = None  # type: Optional[threading.Thread]
        self.auto_stop_cancel = threading.Event()
        self.auto_run_cancel = threading.Event()
        self._com_task_queue = queue.Queue()
        self._com_thread = threading.Thread(target=self._com_worker_loop, daemon=True)
        self._com_thread.start()
        self._controller = None  # type: Optional[CANoeController]
        self._controller_lock = threading.Lock()
        self._runtime_thread = None  # type: Optional[threading.Thread]
        self._runtime_stop = threading.Event()
        self._runtime_started_at = None  # type: Optional[float]
        self._runtime_target_s = None  # type: Optional[float]
        self._reset_analysis_cache = {}  # type: Dict[str, Dict[str, Any]]

        self.cfg_var = tk.StringVar()
        self.customer_folder_var = tk.StringVar()
        self.request_folder_var = tk.StringVar()
        self.tester_folder_var = tk.StringVar()
        default_can_file = Path(__file__).resolve().parent / "diagtester_Security.can"
        self.can_file_var = tk.StringVar(value=str(default_can_file) if default_can_file.exists() else "")
        self.target_var = tk.StringVar(value="RADAR")
        self.customer_replay_block_var = tk.StringVar(value="3")
        self.request_replay_block_var = tk.StringVar(value="1")
        self.hidden_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready")
        self.runtime_var = tk.StringVar(value="CANoe runtime: 0.0s")
        self.env_maxhold_var = tk.StringVar(value="Env_MaxholdStartup: -")
        self.tester_summary_var = tk.StringVar(value="")
        self.exec_summary_var = tk.StringVar(value="")

        self.pairs = []  # type: List[Dict[str, Any]]

        self._load_settings()
        self._apply_target_defaults(self.target_var.get())
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_queue()

    def _com_worker_loop(self) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        try:
            while True:
                task = self._com_task_queue.get()
                if task is None:
                    break
                task()
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def _run_on_com_thread(self, target) -> None:
        done = threading.Event()
        error_holder = []

        def wrapped() -> None:
            try:
                target()
            except Exception as exc:
                error_holder.append(exc)
            finally:
                done.set()

        self._com_task_queue.put(wrapped)
        done.wait()
        if error_holder:
            raise error_holder[0]

    def _load_settings(self) -> None:
        if not self.settings_path.exists():
            return

        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except Exception:
            return

        self.cfg_var.set(str(data.get("cfg", "")))
        customer_folder = data.get("customer_folder", data.get("log_folder", ""))
        self.customer_folder_var.set(str(customer_folder))
        self.request_folder_var.set(str(data.get("request_folder", "")))
        self.tester_folder_var.set(str(data.get("tester_folder", "")))
        self.can_file_var.set(str(data.get("can_file", self.can_file_var.get().strip())))
        self.target_var.set(str(data.get("target", self.target_var.get())).upper())
        self.customer_replay_block_var.set(str(data.get("customer_replay_block", self.customer_replay_block_var.get())))
        self.request_replay_block_var.set(str(data.get("request_replay_block", self.request_replay_block_var.get())))
        self.hidden_var.set(bool(data.get("hidden", False)))

    def _save_settings(self) -> None:
        payload = {
            "cfg": self.cfg_var.get().strip(),
            "customer_folder": self.customer_folder_var.get().strip(),
            "request_folder": self.request_folder_var.get().strip(),
            "tester_folder": self.tester_folder_var.get().strip(),
            "can_file": self.can_file_var.get().strip(),
            "target": self.target_var.get().strip().upper(),
            "customer_replay_block": self.customer_replay_block_var.get().strip(),
            "request_replay_block": self.request_replay_block_var.get().strip(),
            "hidden": self.hidden_var.get(),
        }

        try:
            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            self.settings_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _on_close(self) -> None:
        self._save_settings()
        self._com_task_queue.put(None)
        if self._com_thread.is_alive():
            self._com_thread.join(timeout=1.0)
        self.root.destroy()

    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        cfg_row = ttk.Frame(container)
        cfg_row.pack(fill=tk.X, pady=3)
        ttk.Label(cfg_row, text="CANoe cfg/cfgx:", width=18).pack(side=tk.LEFT)
        ttk.Entry(cfg_row, textvariable=self.cfg_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(cfg_row, text="Browse", command=self._browse_cfg).pack(side=tk.LEFT)

        customer_logs_row = ttk.Frame(container)
        customer_logs_row.pack(fill=tk.X, pady=3)
        ttk.Label(customer_logs_row, text="Diag_customer folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(customer_logs_row, textvariable=self.customer_folder_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6
        )
        ttk.Button(customer_logs_row, text="Browse", command=self._browse_customer_logs).pack(side=tk.LEFT)

        request_logs_row = ttk.Frame(container)
        request_logs_row.pack(fill=tk.X, pady=3)
        ttk.Label(request_logs_row, text="Diag_request folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(request_logs_row, textvariable=self.request_folder_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6
        )
        ttk.Button(request_logs_row, text="Browse", command=self._browse_request_logs).pack(side=tk.LEFT)

        tester_logs_row = ttk.Frame(container)
        tester_logs_row.pack(fill=tk.X, pady=3)
        ttk.Label(tester_logs_row, text="Diag_tester_logs folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(tester_logs_row, textvariable=self.tester_folder_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6
        )
        ttk.Button(tester_logs_row, text="Browse", command=self._browse_tester_logs).pack(side=tk.LEFT)

        can_file_row = ttk.Frame(container)
        can_file_row.pack(fill=tk.X, pady=3)
        ttk.Label(can_file_row, text="CAPL .can file:", width=18).pack(side=tk.LEFT)
        ttk.Entry(can_file_row, textvariable=self.can_file_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(can_file_row, text="Browse", command=self._browse_can_file).pack(side=tk.LEFT)

        replay_block_row = ttk.Frame(container)
        replay_block_row.pack(fill=tk.X, pady=3)
        ttk.Label(replay_block_row, text="Replay blocks:", width=18).pack(side=tk.LEFT)
        ttk.Label(replay_block_row, text="Customer").pack(side=tk.LEFT)
        ttk.Entry(replay_block_row, textvariable=self.customer_replay_block_var, width=6).pack(side=tk.LEFT, padx=(6, 14))
        ttk.Label(replay_block_row, text="DiagRequest").pack(side=tk.LEFT)
        ttk.Entry(replay_block_row, textvariable=self.request_replay_block_var, width=6).pack(side=tk.LEFT, padx=6)

        options_row = ttk.Frame(container)
        options_row.pack(fill=tk.X, pady=3)
        ttk.Label(options_row, text="Target:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            options_row,
            text="MPC",
            value="MPC",
            variable=self.target_var,
            command=self._on_target_changed,
        ).pack(side=tk.LEFT, padx=(6, 2))
        ttk.Radiobutton(
            options_row,
            text="RADAR",
            value="RADAR",
            variable=self.target_var,
            command=self._on_target_changed,
        ).pack(side=tk.LEFT, padx=(2, 10))
        ttk.Checkbutton(options_row, text="Hide CANoe window", variable=self.hidden_var).pack(side=tk.LEFT, padx=8)

        action_row = ttk.Frame(container)
        action_row.pack(fill=tk.X, pady=5)

        cfg_group = ttk.LabelFrame(action_row, text="Configuration", padding=6)
        cfg_group.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(cfg_group, text="Load CFG", command=self.load_cfg).pack(side=tk.LEFT)

        pair_group = ttk.LabelFrame(action_row, text="Pair Selection", padding=6)
        pair_group.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(pair_group, text="Scan Pairs", command=self.scan_pairs).pack(side=tk.LEFT)
        ttk.Button(pair_group, text="Select All", command=self.select_all_pairs).pack(side=tk.LEFT, padx=6)

        run_group = ttk.LabelFrame(action_row, text="Execution", padding=6)
        run_group.pack(side=tk.LEFT)
        ttk.Button(run_group, text="Apply Selected Pair", command=self.apply_selected_pair).pack(side=tk.LEFT)
        ttk.Button(run_group, text="Auto Run Selected", command=self.auto_execute_selected_pairs).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(run_group, text="Stop Auto Run", command=self.stop_auto_execution).pack(side=tk.LEFT)

        help_text = (
            "Pair rule: files from Diag_customer folder and Diag_request folder are paired by remaining "
            "filename part after prefix (case-insensitive). You can multi-select pairs or use Select All."
        )
        ttk.Label(container, text=help_text, foreground="#555555").pack(anchor="w", pady=(0, 8))

        main_row = ttk.Frame(container)
        main_row.pack(fill=tk.BOTH, expand=True, pady=4)

        tree_frame = ttk.Frame(main_row)
        tree_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        summary_panel = ttk.LabelFrame(main_row, text="Summary", padding=8)
        summary_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))

        tester_summary_frame = ttk.LabelFrame(summary_panel, text="Tester Log", padding=6)
        tester_summary_frame.pack(fill=tk.X, pady=(0, 8))
        self.tester_summary_canvas = tk.Canvas(
            tester_summary_frame,
            width=170,
            height=130,
            bg="#ffffff",
            highlightthickness=1,
            highlightbackground="#d0d0d0",
        )
        self.tester_summary_canvas.pack()
        ttk.Label(tester_summary_frame, textvariable=self.tester_summary_var, justify=tk.LEFT).pack(anchor="w", pady=(6, 0))

        exec_summary_frame = ttk.LabelFrame(summary_panel, text="Executing Status", padding=6)
        exec_summary_frame.pack(fill=tk.X)
        self.exec_summary_canvas = tk.Canvas(
            exec_summary_frame,
            width=170,
            height=130,
            bg="#ffffff",
            highlightthickness=1,
            highlightbackground="#d0d0d0",
        )
        self.exec_summary_canvas.pack()
        ttk.Label(exec_summary_frame, textvariable=self.exec_summary_var, justify=tk.LEFT).pack(anchor="w", pady=(6, 0))

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("pair", "customer", "diag", "latest", "resets", "timings", "tester", "state"),
            show="headings",
            selectmode="extended",
            height=16,
        )
        self.tree.heading("pair", text="Pair Key")
        self.tree.heading("customer", text="Customer Replay")
        self.tree.heading("diag", text="DiagRequest Replay")
        self.tree.heading("latest", text="Auto Stop (s)")
        self.tree.heading("resets", text="Reset Count")
        self.tree.heading("timings", text="Reset Timings (OFF/ON)")
        self.tree.heading("tester", text="Tester Log")
        self.tree.heading("state", text="Executing Status")

        self.tree.column("pair", width=190, anchor=tk.CENTER, stretch=False)
        self.tree.column("customer", width=270)
        self.tree.column("diag", width=270)
        self.tree.column("latest", width=120, anchor=tk.CENTER, stretch=False)
        self.tree.column("resets", width=100, anchor=tk.CENTER, stretch=False)
        self.tree.column("timings", width=340)
        self.tree.column("tester", width=120, anchor=tk.CENTER, stretch=False)
        self.tree.column("state", width=170, anchor=tk.CENTER, stretch=False)

        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Control-a>", self._on_select_all_shortcut)
        self.tree.tag_configure("state_done", background="#d3f9d8")

        bottom_row = ttk.Frame(container)
        bottom_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(bottom_row, textvariable=self.status_var).pack(side=tk.LEFT, anchor="w")

        status_right = ttk.Frame(bottom_row)
        status_right.pack(side=tk.RIGHT, anchor="e")
        ttk.Label(status_right, textvariable=self.runtime_var).pack(anchor="e")
        ttk.Label(status_right, textvariable=self.env_maxhold_var).pack(anchor="e")

        self._refresh_summary_charts()

    def _browse_cfg(self) -> None:
        path = filedialog.askopenfilename(
            title="Select CANoe configuration",
            filetypes=[("CANoe Config", "*.cfg *.cfgx"), ("All files", "*.*")],
        )
        if path:
            self.cfg_var.set(path)
            self._save_settings()

    def _browse_customer_logs(self) -> None:
        path = filedialog.askdirectory(title="Select folder that contains Diag_customer logs")
        if path:
            self.customer_folder_var.set(path)
            self._save_settings()

    def _browse_request_logs(self) -> None:
        path = filedialog.askdirectory(title="Select folder that contains Diag_request logs")
        if path:
            self.request_folder_var.set(path)
            self._save_settings()

    def _browse_tester_logs(self) -> None:
        path = filedialog.askdirectory(title="Select folder that contains Diag_tester logs")
        if path:
            self.tester_folder_var.set(path)
            self._save_settings()
            self._refresh_tester_log_statuses_for_all_rows()
            self._refresh_summary_charts()

    def _browse_can_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select CAPL .can file",
            filetypes=[("CAPL File", "*.can"), ("All files", "*.*")],
        )
        if path:
            self.can_file_var.set(path)
            self._save_settings()

    def _apply_target_defaults(self, target: str) -> None:
        clean_target = str(target or "").strip().upper()
        if clean_target == "MPC":
            self.customer_replay_block_var.set("2")
            self.request_replay_block_var.set("1")
            self.target_var.set("MPC")
            return

        # Default to RADAR mapping.
        self.customer_replay_block_var.set("3")
        self.request_replay_block_var.set("1")
        self.target_var.set("RADAR")

    def _on_target_changed(self) -> None:
        self._apply_target_defaults(self.target_var.get())
        self._save_settings()

    @staticmethod
    def _extract_suffix_after_prefix(file_name: str, prefix: str) -> Optional[str]:
        stem = Path(file_name).stem
        lower_stem = stem.lower()
        if not lower_stem.startswith(prefix):
            return None

        suffix = stem[len(prefix) :].lstrip("_- ")
        if not suffix:
            return None
        return suffix

    @staticmethod
    def _extract_pair_key(file_name: str) -> Optional[str]:
        customer_suffix = CanoeReplayPairGui._extract_suffix_after_prefix(file_name, CUSTOMER_PREFIX)
        if customer_suffix is not None:
            return customer_suffix.lower()

        request_suffix = CanoeReplayPairGui._extract_suffix_after_prefix(file_name, REQUEST_PREFIX)
        if request_suffix is not None:
            return request_suffix.lower()

        return None

    @staticmethod
    def _is_customer_file(file_name: str) -> bool:
        return CanoeReplayPairGui._extract_suffix_after_prefix(file_name, CUSTOMER_PREFIX) is not None

    @staticmethod
    def _is_request_file(file_name: str) -> bool:
        return CanoeReplayPairGui._extract_suffix_after_prefix(file_name, REQUEST_PREFIX) is not None

    @staticmethod
    def _is_log_file(path: Path) -> bool:
        return path.suffix.lower() in {".blf", ".asc"}

    @staticmethod
    def _natural_sort_key(text: str) -> List[Any]:
        parts = re.split(r"(\d+)", text.lower())
        key = []  # type: List[Any]
        for part in parts:
            if part == "":
                continue
            if part.isdigit():
                key.append((0, int(part)))
            else:
                key.append((1, part))
        return key

    @staticmethod
    def _read_latest_asc_timestamp(path: Path) -> Optional[float]:
        latest = None  # type: Optional[float]
        try:
            with path.open("r", encoding="utf-8", errors="ignore") as handle:
                for line in handle:
                    match = re.match(r"^\s*(\d+(?:\.\d+)?)\s+", line)
                    if not match:
                        continue
                    try:
                        timestamp = float(match.group(1))
                    except ValueError:
                        continue
                    latest = timestamp
        except OSError:
            return None

        return latest

    def _validate_cfg(self) -> Path:
        cfg_path = Path(self.cfg_var.get().strip()).expanduser()
        if not cfg_path.exists() or not cfg_path.is_file():
            raise FileNotFoundError("Please choose a valid CANoe .cfg/.cfgx file")
        return cfg_path

    def _validate_customer_folder(self) -> Path:
        customer_folder = Path(self.customer_folder_var.get().strip()).expanduser()
        if not customer_folder.exists() or not customer_folder.is_dir():
            raise FileNotFoundError("Please choose a valid Diag_customer folder")
        return customer_folder

    def _validate_request_folder(self) -> Path:
        request_folder = Path(self.request_folder_var.get().strip()).expanduser()
        if not request_folder.exists() or not request_folder.is_dir():
            raise FileNotFoundError("Please choose a valid Diag_request folder")
        return request_folder

    def _get_tester_folder_optional(self) -> Optional[Path]:
        raw_value = self.tester_folder_var.get().strip()
        if raw_value == "":
            return None
        tester_folder = Path(raw_value).expanduser()
        if not tester_folder.exists() or not tester_folder.is_dir():
            raise FileNotFoundError("Please choose a valid Diag_tester_logs folder")
        return tester_folder

    @staticmethod
    def _tester_log_status_for_pair(customer_path: Path, tester_folder: Optional[Path]) -> str:
        if tester_folder is None:
            return "No Folder"
        expected_name = CANoeController._build_diag_tester_name(customer_path)
        expected_path = tester_folder / expected_name
        return "Available" if expected_path.exists() else "Missing"

    @staticmethod
    def _state_from_tester_status(tester_status: str) -> str:
        return "Done" if str(tester_status).strip().lower() == "available" else "Pending"

    def _validate_can_file(self) -> Path:
        can_file_path = Path(self.can_file_var.get().strip()).expanduser()
        if not can_file_path.exists() or not can_file_path.is_file() or can_file_path.suffix.lower() != ".can":
            raise FileNotFoundError("Please choose a valid CAPL .can file")
        return can_file_path

    @staticmethod
    def _parse_replay_block_index(raw_value: str, field_name: str) -> int:
        text = str(raw_value).strip()
        if text == "":
            raise ValueError("{0} replay block must not be empty".format(field_name))
        try:
            index = int(text)
        except ValueError as exc:
            raise ValueError("{0} replay block must be an integer".format(field_name)) from exc
        if index <= 0:
            raise ValueError("{0} replay block must be >= 1".format(field_name))
        return index

    def _validate_replay_block_indices(self) -> Tuple[int, int]:
        customer_index = self._parse_replay_block_index(self.customer_replay_block_var.get(), "Customer")
        request_index = self._parse_replay_block_index(self.request_replay_block_var.get(), "DiagRequest")
        return customer_index, request_index

    def _extract_reset_profile(self, customer_path: Path) -> Tuple[List[float], List[float]]:
        if customer_path.suffix.lower() != ".asc":
            return [], []

        timestamps = extract_vsa_timestamps(customer_path)
        events = find_reset_events(timestamps)
        off_delays_ms, on_delays_ms = build_replay_delays_ms(events)
        return off_delays_ms, on_delays_ms

    @staticmethod
    def _cache_key_for_path(path: Path) -> str:
        try:
            return str(path.resolve()).lower()
        except Exception:
            return str(path).lower()

    @staticmethod
    def _file_signature(path: Path) -> Optional[str]:
        try:
            stat = path.stat()
        except OSError:
            return None
        return "{0}:{1}".format(stat.st_mtime_ns, stat.st_size)

    def _get_reset_analysis_cached(self, customer_path: Path) -> Tuple[Optional[float], List[float], List[float]]:
        if customer_path.suffix.lower() != ".asc":
            return None, [], []

        cache_key = self._cache_key_for_path(customer_path)
        signature = self._file_signature(customer_path)
        cache_entry = self._reset_analysis_cache.get(cache_key)

        if signature and cache_entry and cache_entry.get("signature") == signature:
            latest_cached = cache_entry.get("latest_customer_ts")
            off_cached = list(cache_entry.get("off_delays_ms", []))
            on_cached = list(cache_entry.get("on_delays_ms", []))
            return latest_cached, off_cached, on_cached

        latest_customer_ts = self._read_latest_asc_timestamp(customer_path)
        off_delays_ms, on_delays_ms = self._extract_reset_profile(customer_path)

        if signature:
            self._reset_analysis_cache[cache_key] = {
                "signature": signature,
                "latest_customer_ts": latest_customer_ts,
                "off_delays_ms": list(off_delays_ms),
                "on_delays_ms": list(on_delays_ms),
            }

        return latest_customer_ts, off_delays_ms, on_delays_ms

    def _refresh_tester_log_statuses_for_all_rows(self) -> None:
        tester_folder = None
        try:
            tester_folder = self._get_tester_folder_optional()
        except Exception:
            tester_folder = None

        try:
            tester_col_index = list(self.tree["columns"]).index("tester")
        except Exception:
            tester_col_index = -1

        if tester_col_index < 0:
            return

        for pair in self.pairs:
            status = self._tester_log_status_for_pair(Path(pair["customer"]), tester_folder)
            pair["tester_log_status"] = status
            pair["state"] = self._state_from_tester_status(status)
            item_id = str(pair.get("item_id", ""))
            if not item_id:
                continue
            values = list(self.tree.item(item_id, "values"))
            state_col_index = -1
            try:
                state_col_index = list(self.tree["columns"]).index("state")
            except Exception:
                state_col_index = -1
            if tester_col_index < len(values):
                values[tester_col_index] = status
                if state_col_index >= 0 and state_col_index < len(values):
                    values[state_col_index] = pair["state"]
                self.tree.item(item_id, values=values)
                tags = ["state_done"] if pair["state"].lower() == "done" else []
                self.tree.item(item_id, tags=tags)

    @staticmethod
    def _format_summary_lines(counts: Dict[str, int], total: int) -> str:
        if total <= 0:
            return "Total: 0"
        lines = ["Total: {0}".format(total)]
        for key in sorted(counts.keys()):
            value = counts[key]
            percent = (float(value) / float(total)) * 100.0
            lines.append("{0}: {1} ({2:.1f}%)".format(key, value, percent))
        return "\n".join(lines)

    def _draw_pie_chart(self, canvas: tk.Canvas, counts: Dict[str, int], color_map: Dict[str, str]) -> None:
        canvas.delete("all")
        width = int(canvas.cget("width"))
        height = int(canvas.cget("height"))
        cx = width // 2
        cy = height // 2
        radius = min(width, height) // 2 - 14
        left = cx - radius
        top = cy - radius
        right = cx + radius
        bottom = cy + radius

        canvas.create_oval(left, top, right, bottom, fill="#f7f7f7", outline="#d6d6d6")

        total = sum(counts.values())
        if total <= 0:
            canvas.create_oval(left, top, right, bottom, fill="#efefef", outline="#cccccc")
            canvas.create_text(cx, cy, text="No Data", fill="#666666")
            return

        start = 0.0
        slices = []  # type: List[Tuple[str, float]]
        for key, value in counts.items():
            if value <= 0:
                continue
            slices.append((key, 360.0 * (float(value) / float(total))))

        # Tkinter can fail to render a 360-degree arc consistently.
        # For a single 100% category, draw a full circle directly.
        if len(slices) == 1:
            only_key = slices[0][0]
            fill_color = color_map.get(only_key, "#cccccc")
            canvas.create_oval(left, top, right, bottom, fill=fill_color, outline="#ffffff")
            return

        for idx, (key, raw_extent) in enumerate(slices):
            extent = raw_extent
            if idx < len(slices) - 1 and extent > 0.0:
                extent = max(extent, 1.0)
            fill_color = color_map.get(key, "#cccccc")
            canvas.create_arc(left, top, right, bottom, start=start, extent=extent, fill=fill_color, outline="#ffffff")
            start += extent

    def _refresh_summary_charts(self) -> None:
        tester_counts = {
            "Available": 0,
            "Missing": 0,
        }
        exec_counts = {
            "Done": 0,
            "Pending": 0,
        }

        for pair in self.pairs:
            tester_key = str(pair.get("tester_log_status", "Missing"))
            if tester_key not in tester_counts:
                tester_key = "Missing"
            tester_counts[tester_key] += 1

            state_key = str(pair.get("state", "Pending")).strip().title()
            if state_key not in exec_counts:
                state_key = "Pending"
            exec_counts[state_key] += 1

        tester_total = sum(tester_counts.values())
        exec_total = sum(exec_counts.values())

        self._draw_pie_chart(
            self.tester_summary_canvas,
            tester_counts,
            {
                "Available": "#69db7c",
                "Missing": "#ff8787",
            },
        )
        self._draw_pie_chart(
            self.exec_summary_canvas,
            exec_counts,
            {
                "Done": "#51cf66",
                "Pending": "#fab005",
            },
        )
        self.tester_summary_var.set(self._format_summary_lines(tester_counts, tester_total))
        self.exec_summary_var.set(self._format_summary_lines(exec_counts, exec_total))

    def scan_pairs(self) -> None:
        self._save_settings()

        try:
            customer_folder = self._validate_customer_folder()
            request_folder = self._validate_request_folder()
            tester_folder = self._get_tester_folder_optional()
        except Exception as exc:
            messagebox.showerror("Invalid Folder", str(exc))
            return

        self.tree.delete(*self.tree.get_children())
        self.pairs = []

        customers = {}  # type: Dict[str, List[Path]]
        requests = {}  # type: Dict[str, List[Path]]

        for file_path in sorted(customer_folder.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file() or not self._is_log_file(file_path):
                continue

            pair_key = self._extract_pair_key(file_path.name)
            if pair_key is None:
                continue

            if self._is_customer_file(file_path.name):
                customers.setdefault(pair_key, []).append(file_path)

        for file_path in sorted(request_folder.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file() or not self._is_log_file(file_path):
                continue

            pair_key = self._extract_pair_key(file_path.name)
            if pair_key is None:
                continue

            if self._is_request_file(file_path.name):
                requests.setdefault(pair_key, []).append(file_path)

        created = 0
        all_keys = sorted(
            set(customers.keys()) | set(requests.keys()),
            key=self._natural_sort_key,
        )
        for pair_key in all_keys:
            customer_candidates = sorted(
                customers.get(pair_key, []), key=lambda p: self._natural_sort_key(p.name)
            )
            request_candidates = sorted(
                requests.get(pair_key, []), key=lambda p: self._natural_sort_key(p.name)
            )

            if not customer_candidates or not request_candidates:
                continue

            customer_path = customer_candidates[0]
            diag_path = request_candidates[0]
            latest_customer_ts = None
            off_delays_ms = []  # type: List[float]
            on_delays_ms = []  # type: List[float]
            if customer_path.suffix.lower() == ".asc":
                latest_customer_ts, off_delays_ms, on_delays_ms = self._get_reset_analysis_cached(customer_path)

            pair = {
                "pair_key": pair_key,
                "customer": customer_path,
                "diag": diag_path,
                "latest_customer_ts": latest_customer_ts,
                "auto_stop_s": latest_customer_ts,
                "off_delays_ms": off_delays_ms,
                "on_delays_ms": on_delays_ms,
                "reset_count": len(off_delays_ms),
                "reset_summary": summarize_reset_profile(off_delays_ms, on_delays_ms),
                "tester_log_status": self._tester_log_status_for_pair(customer_path, tester_folder),
                "state": "Pending",
            }
            pair["state"] = self._state_from_tester_status(pair["tester_log_status"])
            latest_text = ""
            if pair["auto_stop_s"] is not None:
                latest_text = "{0:.6f}".format(pair["auto_stop_s"])
            item_id = self.tree.insert(
                "",
                tk.END,
                values=(
                    pair_key,
                    customer_path.name,
                    diag_path.name,
                    latest_text,
                    pair["reset_count"],
                    pair["reset_summary"],
                    pair["tester_log_status"],
                    pair["state"],
                ),
            )
            pair["item_id"] = item_id
            self.pairs.append(pair)
            created += 1

        if created == 0:
            self.status_var.set("No valid Customer/DiagRequest pairs found in selected folders")
        else:
            unmatched_customers = len([k for k in customers if k not in requests])
            unmatched_requests = len([k for k in requests if k not in customers])
            self.status_var.set(
                "Found {0} replay pair(s). Unmatched: Customer={1}, Request={2}.".format(
                    created, unmatched_customers, unmatched_requests
                )
            )

        self._refresh_summary_charts()

    def _run_async(self, target, start_message: str) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Busy", "Another CANoe operation is running. Please wait.")
            return

        self.status_var.set(start_message)

        def worker_wrapper() -> None:
            try:
                self._run_on_com_thread(target)
            except Exception as exc:
                self._put_result(False, "Operation failed: {0}".format(exc))

        self.worker_thread = threading.Thread(target=worker_wrapper, daemon=True)
        self.worker_thread.start()

    def _put_result(self, success: bool, message: str, row_item_id: Optional[str] = None) -> None:
        self.ui_queue.put(("result", (success, message, row_item_id)))

    def _put_row_state(self, row_item_id: str, state_text: str) -> None:
        self.ui_queue.put(("row_state", (row_item_id, state_text)))

    def _put_progress(self, message: str) -> None:
        self.ui_queue.put(("progress", (message,)))

    def _put_env_maxhold(self, value: Any) -> None:
        self.ui_queue.put(("env_maxhold", (value,)))

    @staticmethod
    def _to_numeric(value: Any) -> float:
        if isinstance(value, bool):
            return float(int(value))
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).strip()
        if text == "":
            raise ValueError("empty value")
        return float(text)

    def _mark_measurement_started(self, target_s: Optional[float] = None) -> None:
        self._runtime_target_s = target_s if (target_s is not None and target_s > 0) else None
        self._runtime_started_at = time.monotonic()
        self._runtime_stop.set()
        self._runtime_stop = threading.Event()
        local_stop = self._runtime_stop

        def monitor() -> None:
            while not local_stop.is_set():
                if self._runtime_started_at is None:
                    break
                elapsed = max(0.0, time.monotonic() - self._runtime_started_at)
                self.ui_queue.put(("runtime", (elapsed, self._runtime_target_s)))
                time.sleep(0.2)

        self._runtime_thread = threading.Thread(target=monitor, daemon=True)
        self._runtime_thread.start()

    def _mark_measurement_stopped(self) -> None:
        self._runtime_stop.set()
        self._runtime_started_at = None
        self._runtime_target_s = None
        self.ui_queue.put(("runtime", (0.0, None)))
        self._put_env_maxhold("-")

    def _ensure_env_maxhold_startup(self, controller: CANoeController, timeout_s: float, max_tries: int = 5) -> float:
        for attempt in range(1, max_tries + 1):
            time.sleep(5.0)
            raw_value = controller.read_environment_variable("Env_MaxholdStartup")
            self._put_env_maxhold(raw_value)
            numeric_value = self._to_numeric(raw_value)
            self._put_progress(
                "Env_MaxholdStartup attempt {0}/{1}: {2}".format(attempt, max_tries, raw_value)
            )

            if numeric_value != 0.0:
                return numeric_value

            if attempt >= max_tries:
                raise RuntimeError(
                    "Env_MaxholdStartup remained 0 after {0} attempts; aborted to avoid endless restart loop".format(
                        max_tries
                    )
                )

            self._put_progress(
                "Env_MaxholdStartup is 0. Restarting CANoe measurement ({0}/{1})".format(attempt, max_tries)
            )
            controller.stop(timeout_s=timeout_s)
            controller.start(timeout_s=timeout_s)

    def _create_controller(self) -> CANoeController:
        visible = not self.hidden_var.get()
        with self._controller_lock:
            if self._controller is None:
                self._controller = CANoeController(visible=visible)
            else:
                try:
                    self._controller.app.Visible = bool(visible)
                except Exception:
                    pass
            return self._controller

    def select_all_pairs(self) -> None:
        items = self.tree.get_children()
        if not items:
            return
        self.tree.selection_set(items)
        self.status_var.set("Selected {0} pair(s)".format(len(items)))

    def _on_select_all_shortcut(self, _event=None):
        self.select_all_pairs()
        return "break"

    def _selected_pairs_in_order(self) -> List[Dict[str, Any]]:
        selected_ids = set(self.tree.selection())
        selected_pairs = []  # type: List[Dict[str, Any]]
        for item_id in self.tree.get_children():
            if item_id not in selected_ids:
                continue
            row = next((p for p in self.pairs if p.get("item_id") == item_id), None)
            if row is not None:
                selected_pairs.append(row)
        return selected_pairs

    def _resolve_pair_duration(self, row: Dict[str, Any], fallback_s: float) -> float:
        value = row.get("auto_stop_s")
        try:
            parsed = float(value)
        except (TypeError, ValueError):
            parsed = 0.0
        if parsed <= 0:
            return fallback_s
        return parsed

    def _update_capl_reset_profile_for_pair(self, row: Dict[str, Any], can_file_path: Path) -> int:
        customer_path = Path(row["customer"])
        off_delays_ms = list(row.get("off_delays_ms", []))
        on_delays_ms = list(row.get("on_delays_ms", []))

        # Recompute only when the source log changed; unchanged logs use cached analysis.
        if customer_path.suffix.lower() == ".asc":
            _, off_delays_ms, on_delays_ms = self._get_reset_analysis_cached(customer_path)

        updater = CaplResetProfileUpdater(can_file_path)
        updated_count = updater.update_profile(off_delays_ms, on_delays_ms, customer_path.name)

        row["off_delays_ms"] = off_delays_ms
        row["on_delays_ms"] = on_delays_ms
        row["reset_count"] = updated_count
        row["reset_summary"] = summarize_reset_profile(off_delays_ms, on_delays_ms)
        return updated_count

    def _on_tree_double_click(self, event) -> None:
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        if column != "#4":
            return

        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        x, y, width, height = self.tree.bbox(item_id, column)
        if width <= 0 or height <= 0:
            return

        current_values = list(self.tree.item(item_id, "values"))
        current_text = ""
        if len(current_values) >= 4:
            current_text = str(current_values[3])

        editor = ttk.Entry(self.tree)
        editor.insert(0, current_text)
        editor.select_range(0, tk.END)
        editor.focus_set()
        editor.place(x=x, y=y, width=width, height=height)

        def commit_edit(_event=None) -> None:
            new_value_text = editor.get().strip()
            try:
                if new_value_text == "":
                    parsed = None
                    display_text = ""
                else:
                    parsed = float(new_value_text)
                    if parsed < 0:
                        raise ValueError("must be >= 0")
                    display_text = "{0:.6f}".format(parsed)
            except ValueError:
                messagebox.showerror("Invalid Value", "Auto stop value must be empty or a non-negative number")
                editor.focus_set()
                return

            if len(current_values) >= 4:
                current_values[3] = display_text
                self.tree.item(item_id, values=current_values)

            row = next((p for p in self.pairs if p.get("item_id") == item_id), None)
            if row is not None:
                row["auto_stop_s"] = parsed

            editor.destroy()

        def cancel_edit(_event=None) -> None:
            editor.destroy()

        editor.bind("<Return>", commit_edit)
        editor.bind("<Escape>", cancel_edit)
        editor.bind("<FocusOut>", commit_edit)

    def _start_auto_stop_monitor(self, auto_stop_s: float) -> None:
        if auto_stop_s <= 0:
            return

        self.auto_stop_cancel.set()
        self.auto_stop_cancel = threading.Event()
        local_cancel = self.auto_stop_cancel

        def monitor() -> None:
            started_at = time.monotonic()
            while not local_cancel.is_set():
                elapsed = time.monotonic() - started_at
                if elapsed >= auto_stop_s:
                    break
                time.sleep(0.1)

            if local_cancel.is_set():
                return

            try:
                def stop_task() -> None:
                    controller = self._create_controller()
                    controller.stop(timeout_s=15.0)

                self._run_on_com_thread(stop_task)
                self._mark_measurement_stopped()
                self.ui_queue.put(("auto_stop", (True, auto_stop_s)))
            except Exception as exc:
                self.ui_queue.put(("auto_stop", (False, "Auto stop failed: {0}".format(exc))))

        self.auto_stop_thread = threading.Thread(target=monitor, daemon=True)
        self.auto_stop_thread.start()

    def _selected_row_auto_stop(self) -> Optional[float]:
        selected = self.tree.selection()
        if not selected:
            return None

        item_id = selected[0]
        row = next((p for p in self.pairs if p.get("item_id") == item_id), None)
        if row is None:
            return None

        value = row.get("auto_stop_s")
        if value is None:
            return None

        try:
            parsed = float(value)
        except (TypeError, ValueError):
            return None

        if parsed <= 0:
            return None
        return parsed

    def load_cfg(self) -> None:
        self._save_settings()

        try:
            cfg_path = self._validate_cfg()
        except Exception as exc:
            messagebox.showerror("Invalid CFG", str(exc))
            return

        def worker() -> None:
            try:
                controller = self._create_controller()
                controller.load_cfg(cfg_path)
                self._put_result(True, "CFG loaded in CANoe")
            except Exception as exc:
                self._put_result(False, "Load cfg failed: {0}".format(exc))

        self._run_async(worker, "Loading cfg into CANoe...")

    def query_status(self) -> None:
        def worker() -> None:
            try:
                controller = self._create_controller()
                state = "running" if controller.measurement_running else "stopped"
                env_value = controller.read_environment_variable("Env_MaxholdStartup")
                self._put_env_maxhold(env_value)
                self._put_result(True, "Measurement is {0}. Env_MaxholdStartup={1}".format(state, env_value))
            except Exception as exc:
                self._put_result(False, "Read status failed: {0}".format(exc))

        self._run_async(worker, "Reading CANoe measurement state...")

    def start_measurement(self) -> None:
        timeout = CONTROL_TIMEOUT_S

        auto_stop_s = self._selected_row_auto_stop()

        if auto_stop_s is not None:
            self.status_var.set("Starting measurement... auto-stop at {0:.6f}s".format(auto_stop_s))
        else:
            self.status_var.set("Starting measurement...")

        def worker() -> None:
            try:
                controller = self._create_controller()
                controller.start(timeout_s=timeout)
                self._mark_measurement_started(target_s=auto_stop_s)
                env_value = self._ensure_env_maxhold_startup(controller, timeout_s=timeout, max_tries=5)
                if auto_stop_s is not None:
                    self._start_auto_stop_monitor(auto_stop_s)
                    self._put_result(
                        True,
                        "Measurement started. Env_MaxholdStartup={0}. Auto-stop armed at {1:.6f}s".format(
                            env_value, auto_stop_s
                        ),
                    )
                else:
                    self.auto_stop_cancel.set()
                    self._put_result(True, "Measurement started. Env_MaxholdStartup={0}".format(env_value))
            except Exception as exc:
                self._mark_measurement_stopped()
                self._put_result(False, "Start failed: {0}".format(exc))

        self._run_async(worker, "Starting measurement...")

    def stop_measurement(self) -> None:
        timeout = CONTROL_TIMEOUT_S

        def worker() -> None:
            try:
                controller = self._create_controller()
                self.auto_stop_cancel.set()
                controller.stop(timeout_s=timeout)
                self._mark_measurement_stopped()
                self._put_result(True, "Measurement stopped")
            except Exception as exc:
                self._put_result(False, "Stop failed: {0}".format(exc))

        self._run_async(worker, "Stopping measurement...")

    def apply_selected_pair(self) -> None:
        self._save_settings()

        try:
            cfg_path = self._validate_cfg()
            can_file_path = self._validate_can_file()
            customer_block_index, request_block_index = self._validate_replay_block_indices()
        except Exception as exc:
            messagebox.showerror("Invalid Input", str(exc))
            return

        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Select one replay pair first")
            return

        item_id = selected[0]
        row = next((p for p in self.pairs if p["item_id"] == item_id), None)
        if row is None:
            messagebox.showerror("Invalid Selection", "Selected replay pair is not available")
            return

        customer_path = Path(row["customer"])
        diag_path = Path(row["diag"])

        def worker() -> None:
            try:
                controller = self._create_controller()
                reset_count = self._update_capl_reset_profile_for_pair(row, can_file_path)
                controller.load_cfg(cfg_path)
                controller.apply_replay_pair(
                    customer_path,
                    diag_path,
                    customer_block_index=customer_block_index,
                    diag_request_block_index=request_block_index,
                )
                logging_name = controller.Logging_rename(customer_path)
                self._put_result(
                    True,
                    "Replay pair applied (Customer block #{0}, DiagRequest block #{1}). "
                    "CAPL reset profile updated ({2} event(s)). Logging renamed to {3}".format(
                        customer_block_index, request_block_index, reset_count, logging_name
                    ),
                    row_item_id=item_id,
                )
            except Exception as exc:
                self._put_result(False, "Apply pair failed: {0}".format(exc), row_item_id=item_id)

        self._run_async(worker, "Applying replay pair {0}...".format(row["pair_key"]))

    def stop_auto_execution(self) -> None:
        self.auto_run_cancel.set()
        self.status_var.set("Stopping auto execution...")

    def auto_execute_selected_pairs(self) -> None:
        self._save_settings()

        try:
            cfg_path = self._validate_cfg()
            can_file_path = self._validate_can_file()
            customer_block_index, request_block_index = self._validate_replay_block_indices()
        except Exception as exc:
            messagebox.showerror("Invalid Input", str(exc))
            return

        timeout = CONTROL_TIMEOUT_S

        selected_pairs = self._selected_pairs_in_order()
        if not selected_pairs:
            messagebox.showwarning("No Selection", "Select one or more replay pairs first")
            return

        self.auto_run_cancel.clear()

        def worker() -> None:
            controller = None
            total = len(selected_pairs)
            failed_pairs = []  # type: List[str]
            try:
                controller = self._create_controller()
                controller.load_cfg(cfg_path)

                for idx, row in enumerate(selected_pairs, 1):
                    if self.auto_run_cancel.is_set():
                        raise RuntimeError("Auto execution cancelled by user")

                    current_item_id = str(row["item_id"])
                    self._put_row_state(current_item_id, "Pending")

                    customer_path = Path(row["customer"])
                    diag_path = Path(row["diag"])
                    run_duration = self._resolve_pair_duration(row, timeout)

                    measurement_started = False
                    try:
                        reset_count = self._update_capl_reset_profile_for_pair(row, can_file_path)
                        self._put_progress(
                            "[{0}/{1}] Applying pair {2} with {3} reset event(s)".format(
                                idx, total, row["pair_key"], reset_count
                            )
                        )
                        controller.apply_replay_pair(
                            customer_path,
                            diag_path,
                            customer_block_index=customer_block_index,
                            diag_request_block_index=request_block_index,
                        )

                        logging_name = controller.Logging_rename(customer_path)
                        self._put_progress(
                            "[{0}/{1}] Logging renamed to {2}".format(idx, total, logging_name)
                        )

                        controller.start(timeout_s=timeout)
                        measurement_started = True
                        self._mark_measurement_started(target_s=run_duration)
                        env_value = self._ensure_env_maxhold_startup(controller, timeout_s=timeout, max_tries=5)

                        self._put_progress(
                            "[{0}/{1}] Measurement started (Env_MaxholdStartup={2}), waiting {3:.6f}s".format(
                                idx, total, env_value, run_duration
                            )
                        )
                        self._put_env_maxhold(env_value)
                        started_at = time.monotonic()
                        while (time.monotonic() - started_at) < run_duration:
                            if self.auto_run_cancel.is_set():
                                raise RuntimeError("Auto execution cancelled by user")
                            time.sleep(0.1)

                        controller.stop(timeout_s=timeout)
                        measurement_started = False
                        self._mark_measurement_stopped()

                        self._put_row_state(current_item_id, "Done")
                        self._put_progress("[{0}/{1}] Completed pair {2}".format(idx, total, row["pair_key"]))
                    except Exception as pair_exc:
                        if measurement_started:
                            try:
                                controller.stop(timeout_s=timeout)
                                self._mark_measurement_stopped()
                            except Exception:
                                pass
                        self._put_row_state(current_item_id, "Pending")
                        failed_pairs.append(str(row["pair_key"]))
                        self._put_progress(
                            "[{0}/{1}] Failed pair {2}: {3}".format(idx, total, row["pair_key"], pair_exc)
                        )

                if failed_pairs:
                    self._put_result(
                        False,
                        "Auto execution finished with failures ({0}/{1}). Failed: {2}".format(
                            len(failed_pairs), total, ", ".join(failed_pairs)
                        ),
                    )
                else:
                    self._put_result(True, "Auto execution completed for {0} pair(s)".format(total))
            except Exception as exc:
                self._put_result(False, "Auto execution failed: {0}".format(exc))

        self._run_async(worker, "Running selected replay pairs automatically...")

    def _set_row_state(self, item_id: str, state_text: str) -> None:
        values = list(self.tree.item(item_id, "values"))
        if not values:
            return
        try:
            state_index = list(self.tree["columns"]).index("state")
        except ValueError:
            state_index = len(values) - 1
        values[state_index] = state_text

        tester_index = -1
        try:
            tester_index = list(self.tree["columns"]).index("tester")
        except ValueError:
            tester_index = -1

        if tester_index >= 0:
            row = next((p for p in self.pairs if p["item_id"] == item_id), None)
            if row is not None:
                tester_folder = None
                try:
                    tester_folder = self._get_tester_folder_optional()
                except Exception:
                    tester_folder = None
                tester_status = self._tester_log_status_for_pair(Path(row["customer"]), tester_folder)
                row["tester_log_status"] = tester_status
                if tester_index < len(values):
                    values[tester_index] = tester_status

                    derived_state = self._state_from_tester_status(tester_status)
                    row["state"] = derived_state
                    if state_index < len(values):
                        values[state_index] = derived_state

        self.tree.item(item_id, values=values)

        state_key = str(state_text).strip().lower()
        tags = ["state_done"] if state_key == "done" else []
        self.tree.item(item_id, tags=tags)

        for pair in self.pairs:
            if pair["item_id"] == item_id:
                # Executing status is derived from tester log availability.
                pair["state"] = self._state_from_tester_status(str(pair.get("tester_log_status", "")))
                break

        self._refresh_summary_charts()

    def _poll_queue(self) -> None:
        while True:
            try:
                event, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            if event == "result":
                success, message, row_item_id = payload
                if row_item_id:
                    self._set_row_state(row_item_id, "Pending")
                self.status_var.set(message)
                if not success:
                    messagebox.showerror("CANoe Operation Failed", message)
            elif event == "row_state":
                row_item_id, state_text = payload
                self._set_row_state(row_item_id, state_text)
            elif event == "progress":
                message = payload[0]
                self.status_var.set(str(message))
            elif event == "auto_stop":
                success, detail = payload
                if success:
                    self.status_var.set("Measurement auto-stopped at {0:.6f}s".format(detail))
                else:
                    self.status_var.set(str(detail))
                    messagebox.showerror("CANoe Auto Stop Failed", str(detail))
            elif event == "runtime":
                elapsed = float(payload[0])
                target_s = payload[1]
                if target_s is not None:
                    try:
                        target_value = float(target_s)
                    except (TypeError, ValueError):
                        target_value = 0.0
                    if target_value > 0:
                        percent = min(100.0, max(0.0, (elapsed / target_value) * 100.0))
                        self.runtime_var.set(
                            "CANoe runtime: {0:.1f}s / {1:.1f}s ({2:.1f}%)".format(elapsed, target_value, percent)
                        )
                    else:
                        self.runtime_var.set("CANoe runtime: {0:.1f}s".format(elapsed))
                else:
                    self.runtime_var.set("CANoe runtime: {0:.1f}s".format(elapsed))
            elif event == "env_maxhold":
                self.env_maxhold_var.set("Env_MaxholdStartup: {0}".format(payload[0]))

        self.root.after(120, self._poll_queue)


def main() -> None:
    root = tk.Tk()
    root.state("zoomed")
    root.minsize(980, 580)
    CanoeReplayPairGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
