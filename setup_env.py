#!/usr/bin/env python3
"""Cross-platform setup helper for the PESU course downloader."""

from __future__ import annotations

import argparse
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / "venv"
REQUIREMENTS_FILE = ROOT / "requirements.txt"
MAIN_SCRIPT = ROOT / "interactive_download.py"


# ── Terminal helpers (no external deps — setup runs before they exist) ────────

def _supports_color() -> bool:
    if platform.system() == "Windows":
        try:
            import ctypes
            kernel = ctypes.windll.kernel32
            kernel.SetConsoleMode(kernel.GetStdHandle(-11), 7)
            return True
        except Exception:
            return False
    return hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

_USE_COLOR = _supports_color()

def _c(code: str, text: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _USE_COLOR else text

def green(t: str)  -> str: return _c("32", t)
def yellow(t: str) -> str: return _c("33", t)
def cyan(t: str)   -> str: return _c("36", t)
def red(t: str)    -> str: return _c("31", t)
def bold(t: str)   -> str: return _c("1",  t)
def dim(t: str)    -> str: return _c("2",  t)

TICK  = green("✓")
CROSS = red("✗")
ARROW = cyan("→")
SKIP  = dim("·")


# ── Step banner ───────────────────────────────────────────────────────────────

_TOTAL_STEPS = 3   # venv · deps · launch

def step(n: int, label: str) -> None:
    """Print a numbered step header."""
    print(f"\n{bold(f'[{n}/{_TOTAL_STEPS}]')} {cyan(label)}", flush=True)


# ── Errors ────────────────────────────────────────────────────────────────────

class SetupError(RuntimeError):
    """Raised when setup cannot continue safely."""


def format_command(command: list[str]) -> str:
    if platform.system() == "Windows":
        return subprocess.list2cmdline(command)
    import shlex
    return shlex.join(command)


def run(command: list[str], *, cwd: Path = ROOT,
        env: dict[str, str] | None = None,
        capture: bool = False) -> subprocess.CompletedProcess:
    completed = subprocess.run(
        command, cwd=str(cwd), env=env, check=False,
        capture_output=capture, text=capture,
    )
    if completed.returncode != 0:
        raise SetupError(
            f"Command failed (exit {completed.returncode}): {format_command(command)}"
        )
    return completed


# ── OS / Python detection ─────────────────────────────────────────────────────

def read_os_release() -> dict[str, str]:
    os_release = Path("/etc/os-release")
    if not os_release.exists():
        return {}
    data: dict[str, str] = {}
    for line in os_release.read_text(encoding="utf-8").splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        data[key] = value.strip().strip('"')
    return data


def detect_environment() -> str:
    system = platform.system()
    if system == "Windows":
        return "Windows"
    if system == "Darwin":
        return "macOS"
    if system == "Linux":
        os_release = read_os_release()
        distro = os_release.get("ID", "linux")
        return "Ubuntu/Debian" if distro in {"ubuntu", "debian"} else f"Linux ({distro})"
    return system or "unknown"


def candidate_python_commands() -> list[list[str]]:
    system = platform.system()
    commands: list[list[str]] = []
    if system == "Windows":
        commands.extend([["py", "-3.12"], ["py", "-3.11"]])
    else:
        commands.extend([["python3.12"], ["python3.11"], ["python3"]])
    commands.append([sys.executable])
    commands.append(["python"])
    unique: list[list[str]] = []
    seen: set[tuple[str, ...]] = set()
    for command in commands:
        key = tuple(command)
        if key not in seen:
            seen.add(key)
            unique.append(command)
    return unique


def python_version(command: list[str]) -> tuple[int, int, int] | None:
    probe = "import sys; print('%d.%d.%d' % sys.version_info[:3])"
    try:
        completed = subprocess.run(
            [*command, "-c", probe],
            cwd=str(ROOT), check=False, capture_output=True, text=True,
        )
    except (FileNotFoundError, OSError):
        return None
    if completed.returncode != 0:
        return None
    try:
        major, minor, patch = completed.stdout.strip().split(".", 2)
        return int(major), int(minor), int(patch)
    except ValueError:
        return None


def is_supported_python(version: tuple[int, int, int]) -> bool:
    major, minor, _ = version
    if major != 3:
        return False
    if platform.system() == "Windows":
        return minor in {11, 12}
    return minor >= 11


def find_python() -> tuple[list[str], tuple[int, int, int]]:
    attempted: list[str] = []
    for command in candidate_python_commands():
        version = python_version(command)
        attempted.append(format_command(command))
        if version and is_supported_python(version):
            return command, version
    if platform.system() == "Windows":
        requirement = "Python 3.11 or 3.12 is required on Windows (windows-curses limit)."
    else:
        requirement = "Python 3.11 or newer is required."
    raise SetupError(f"{requirement}\nTried: {', '.join(attempted)}")


# ── Virtual environment ───────────────────────────────────────────────────────

def venv_python_path() -> Path:
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def venv_bin_path() -> Path:
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts"
    return VENV_DIR / "bin"


def activated_venv_env() -> dict[str, str]:
    env = os.environ.copy()
    env["VIRTUAL_ENV"] = str(VENV_DIR)
    env["PATH"] = f"{venv_bin_path()}{os.pathsep}{env.get('PATH', '')}"
    env.pop("PYTHONHOME", None)
    return env


def existing_venv_is_usable() -> bool:
    python_path = venv_python_path()
    if not python_path.exists():
        return False
    version = python_version([str(python_path)])
    return bool(version and is_supported_python(version))


def is_debian_like() -> bool:
    if platform.system() != "Linux":
        return False
    os_release = read_os_release()
    ids = {os_release.get("ID", "")}
    ids.update(os_release.get("ID_LIKE", "").split())
    return bool(ids & {"ubuntu", "debian"})


def install_ubuntu_venv_support() -> None:
    if not is_debian_like():
        raise SetupError("Automatic python3-venv install is only supported on Ubuntu/Debian.")
    if shutil.which("sudo") is None:
        raise SetupError("sudo not found. Install python3-venv manually, then rerun.")
    run(["sudo", "apt", "update"])
    run(["sudo", "apt", "install", "-y", "python3-venv"])


def create_venv(python_command: list[str], *, install_system_deps: bool) -> None:
    try:
        run([*python_command, "-m", "venv", str(VENV_DIR)])
    except SetupError:
        if not install_system_deps:
            if is_debian_like():
                raise SetupError(
                    "Failed to create venv. On Ubuntu/Debian run "
                    "`sudo apt install python3-venv` or rerun with --install-system-deps."
                )
            raise
        install_ubuntu_venv_support()
        run([*python_command, "-m", "venv", str(VENV_DIR)])


def ensure_venv(*, recreate: bool, install_system_deps: bool) -> Path:
    step(1, "Virtual environment")

    if recreate and VENV_DIR.exists():
        if VENV_DIR.resolve().parent != ROOT.resolve() or VENV_DIR.name != "venv":
            raise SetupError(f"Refusing to remove unexpected venv path: {VENV_DIR}")
        print(f"  {ARROW} Removing old venv…", flush=True)
        shutil.rmtree(VENV_DIR)

    if existing_venv_is_usable():
        vp = venv_python_path()
        ver = python_version([str(vp)])
        ver_str = ".".join(map(str, ver)) if ver else "?"
        print(f"  {TICK} Reusing existing venv  {dim(f'Python {ver_str}')}", flush=True)
        return vp

    if VENV_DIR.exists():
        raise SetupError(
            f"Existing venv is not usable: {VENV_DIR}\n"
            "Rerun with --recreate to rebuild it."
        )

    python_command, version = find_python()
    env_name = detect_environment()
    ver_str = ".".join(map(str, version))
    print(f"  {dim('Platform:')} {env_name}   {dim('Python:')} {ver_str}", flush=True)
    print(f"  {ARROW} Creating venv…", end=" ", flush=True)
    create_venv(python_command, install_system_deps=install_system_deps)
    print(TICK, flush=True)
    return venv_python_path()


# ── Dependency pre-check and smart install ────────────────────────────────────

def _normalize_name(name: str) -> str:
    """PEP 503 canonical package name."""
    return re.sub(r"[-_.]+", "-", name).lower()


# Requirement entry: (display_name, pip_spec_with_version, marker_or_None)
_Req = tuple[str, str, str | None]


def _venv_marker_env(venv_python: Path) -> dict[str, str]:
    """
    Query platform/version info FROM the venv Python itself.

    Critical: the script may be launched with a different Python than the venv
    (e.g. `py setup.py` resolves to Python 3.13 but the venv was created with
    3.12).  Markers like `python_version<"3.13"` must therefore be evaluated
    against the *venv* Python, not the setup-script Python, or conditional
    packages like windows-curses get incorrectly skipped.
    """
    probe = (
        "import sys, platform; "
        "print(platform.system()); "
        "print(sys.platform); "
        "print(f'{sys.version_info.major}.{sys.version_info.minor}')"
    )
    result = subprocess.run(
        [str(venv_python), "-c", probe],
        capture_output=True, text=True, check=False,
    )
    if result.returncode == 0:
        parts = result.stdout.strip().splitlines()
        if len(parts) == 3:
            return {
                "platform_system": parts[0],
                "sys_platform":    parts[1],
                "python_version":  parts[2],
            }
    # Fallback: use the setup script's own environment
    return {
        "platform_system": platform.system(),
        "sys_platform":    sys.platform,
        "python_version":  f"{sys.version_info.major}.{sys.version_info.minor}",
    }


def _evaluate_marker(marker: str, env: dict[str, str]) -> bool:
    """
    Evaluate a simple PEP 508 environment marker against an explicit env dict.
    Handles the patterns that appear in this project's requirements.txt.
    """
    expr = marker
    for key, val in env.items():
        expr = expr.replace(key, repr(val))
    try:
        return bool(eval(expr))   # noqa: S307 — only our own marker strings
    except Exception:
        return True   # unknown marker → assume it applies (safe default)


def parse_requirements(path: Path, marker_env: dict[str, str] | None = None) -> list[_Req]:
    """
    Parse requirements.txt into (name, pip_spec, marker) triples.

    - Skips blank lines, pure comments, and -r / --flag lines.
    - Strips inline comments that follow whitespace+# (but not ; markers).
    - Returns only entries whose environment marker evaluates True.
    - marker_env: platform/version dict from _venv_marker_env(); falls back to
      the current process environment if not supplied.
    """
    if marker_env is None:
        marker_env = {
            "platform_system": platform.system(),
            "sys_platform":    sys.platform,
            "python_version":  f"{sys.version_info.major}.{sys.version_info.minor}",
        }
    entries: list[_Req] = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or line.startswith("-"):
            continue

        # Separate inline comment (space + #) — must NOT split on ; (marker separator)
        spec_and_marker = re.split(r"\s+#", line, maxsplit=1)[0].strip()

        # Split off PEP 508 marker: "pkg>=1; platform_system=='Windows'"
        if ";" in spec_and_marker:
            spec_part, marker = spec_and_marker.split(";", 1)
            spec_part = spec_part.strip()
            marker    = marker.strip()
        else:
            spec_part = spec_and_marker
            marker    = None

        # Skip if the marker says "not this platform / python"
        if marker and not _evaluate_marker(marker, marker_env):
            continue

        # Bare package name (strip version, extras, whitespace)
        name = re.split(r"[>=<!;\[\s]", spec_part)[0].strip()
        if name:
            entries.append((name, spec_part, marker))

    return entries


def get_installed_packages(venv_python: Path) -> dict[str, str]:
    """Return {normalised-name: version} for every package in the venv."""
    result = subprocess.run(
        [str(venv_python), "-m", "pip", "list", "--format=json"],
        capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        return {}
    try:
        return {
            _normalize_name(pkg["name"]): pkg["version"]
            for pkg in json.loads(result.stdout)
        }
    except (json.JSONDecodeError, KeyError):
        return {}


def _pkg_installed_version(venv_python: Path, name: str) -> str:
    """Return the installed version string for a package, or '' if not found."""
    result = subprocess.run(
        [str(venv_python), "-m", "pip", "show", name],
        capture_output=True, text=True, check=False,
    )
    for line in result.stdout.splitlines():
        if line.lower().startswith("version:"):
            return line.split(":", 1)[1].strip()
    return ""


def _install_one(venv_python: Path, spec: str) -> tuple[bool, str]:
    """
    Install a single package spec (e.g. 'requests>=2.28').
    Returns (success, version_installed).
    """
    result = subprocess.run(
        [str(venv_python), "-m", "pip", "install", spec,
         "--progress-bar", "off", "-q"],
        capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        return False, result.stderr.strip().splitlines()[-1] if result.stderr.strip() else "unknown error"

    # Extract version from "Successfully installed foo-1.2.3"
    for line in result.stdout.splitlines():
        if line.startswith("Successfully installed"):
            parts = line.removeprefix("Successfully installed").split()
            if parts:
                return True, parts[0]   # e.g. "requests-2.31.0"

    # Already satisfied (pip printed nothing) — look it up
    name = re.split(r"[>=<!;\[\s]", spec)[0]
    ver  = _pkg_installed_version(venv_python, name)
    return True, ver or "already satisfied"


def install_requirements(venv_python: Path) -> None:
    step(2, "Dependencies")

    if not REQUIREMENTS_FILE.exists():
        raise SetupError(f"Missing requirements file: {REQUIREMENTS_FILE}")

    # Evaluate markers against the VENV Python, not the setup-script Python.
    # e.g. `py setup.py` may resolve to Python 3.13 while the venv is 3.12 —
    # without this, `python_version<"3.13"` wrongly excludes windows-curses.
    menv = _venv_marker_env(venv_python)
    if menv["python_version"] != f"{sys.version_info.major}.{sys.version_info.minor}":
        print(
            f"  {dim('launcher Python')} {sys.version_info.major}.{sys.version_info.minor}"
            f"  {dim('→  venv Python')} {menv['python_version']}"
        )

    required  = parse_requirements(REQUIREMENTS_FILE, marker_env=menv)
    installed = get_installed_packages(venv_python)

    missing:    list[_Req] = []
    up_to_date: list[_Req] = []

    for entry in required:
        name, spec, _ = entry
        if _normalize_name(name) in installed:
            up_to_date.append(entry)
        else:
            missing.append(entry)

    # ── Already-good packages on one dim line ─────────────────────────────
    if up_to_date:
        print("  " + "  ".join(f"{SKIP} {dim(n)}" for n, _, _ in up_to_date), flush=True)

    if not missing:
        print(f"  {TICK} All {len(required)} packages already installed", flush=True)
        return

    # ── Upgrade pip silently first ────────────────────────────────────────
    subprocess.run(
        [str(venv_python), "-m", "pip", "install", "--upgrade", "pip", "-q"],
        check=False, capture_output=True,
    )

    print(
        f"  {ARROW} Installing {len(missing)} package(s)"
        f"  {dim(f'({len(up_to_date)} already up to date)')}",
        flush=True,
    )

    # ── One-by-one with live line update ──────────────────────────────────
    #   Print "  → pkg_name   " (no newline), then overwrite with ✓ or ✗
    failed: list[str] = []
    COL = 28   # fixed column width so the tick column lines up

    for name, spec, _ in missing:
        label = f"{name}"
        pending = f"  {cyan('→')} {label:<{COL}}"
        print(pending, end="", flush=True)

        ok, detail = _install_one(venv_python, spec)

        if ok:
            print(f"\r  {TICK} {green(label):<{COL}}  {dim(detail)}", flush=True)
        else:
            print(f"\r  {CROSS} {red(label):<{COL}}  {red(detail)}", flush=True)
            failed.append(name)

    # ── Final result ──────────────────────────────────────────────────────
    if failed:
        raise SetupError(f"Failed to install: {', '.join(failed)}")


# ── Summary and launch ────────────────────────────────────────────────────────

def print_setup_summary(venv_python: Path, *, setup_only: bool) -> None:
    print()
    print(bold("Setup complete."), flush=True)

    env_file = ROOT / ".env"
    if not env_file.exists():
        print(
            f"  {yellow('⚠')} No .env found — the downloader will prompt for credentials.\n"
            f"  {dim('Create .env with:')}\n"
            f"      PESU_USERNAME=your_srn\n"
            f"      PESU_PASSWORD=your_password",
            flush=True,
        )

    if setup_only:
        print(flush=True)
        print("To run the downloader later:", flush=True)
        print(f"  {cyan(format_command([str(venv_python), str(MAIN_SCRIPT)]))}", flush=True)


def run_downloader(venv_python: Path) -> None:
    step(3, "Launching downloader")
    run([str(venv_python), str(MAIN_SCRIPT)], env=activated_venv_env())
    print(f"\n{dim('Downloader finished.')}", flush=True)


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Set up the PESU course downloader.")
    parser.add_argument("--recreate",           action="store_true",
                        help="Delete and rebuild the venv.")
    parser.add_argument("--install-system-deps", action="store_true",
                        help="On Ubuntu/Debian, install python3-venv via apt if needed.")
    parser.add_argument("--setup-only",         action="store_true",
                        help="Set up the environment without launching the downloader.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    print(bold("\n  PESU Academy — Setup"), flush=True)
    print(dim("  " + "─" * 38), flush=True)

    try:
        venv_python = ensure_venv(
            recreate=args.recreate,
            install_system_deps=args.install_system_deps,
        )
        install_requirements(venv_python)
        print_setup_summary(venv_python, setup_only=args.setup_only)

        if not args.setup_only:
            run_downloader(venv_python)

    except SetupError as exc:
        print(f"\n{CROSS} {red('Setup failed:')} {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print(f"\n{yellow('Setup cancelled.')}", file=sys.stderr)
        return 130

    return 0


if __name__ == "__main__":
    raise SystemExit(main())