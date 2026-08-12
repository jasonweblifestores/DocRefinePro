"""Putting the computer to sleep once a long run has finished.

Rebranding a few thousand files takes hours, so the machine is usually left
running unattended overnight. This offers what a torrent client offers: finish
the work, then sleep.

Nothing here decides *whether* to sleep — that judgement lives in the caller,
because it depends on how the run ended. A run that was stopped by hand or that
failed must never put the machine to sleep: the person is either still at the
keyboard or needs to read an error.
"""
import ctypes
import subprocess

from .config import SystemUtils, log_app

SLEEP = "sleep"
HIBERNATE = "hibernate"
ACTIONS = (SLEEP, HIBERNATE)

# How long the user gets to change their mind. Long enough to cross a room,
# short enough not to waste the saving.
COUNTDOWN_SECONDS = 60


def describe(action=SLEEP):
    """Wording for the checkbox and the countdown, in this platform's terms."""
    if action == HIBERNATE:
        if SystemUtils.IS_MAC:
            # macOS has no separate user-facing hibernate; sleeping is the whole
            # story, so promising otherwise would be a lie.
            return "sleep"
        return "hibernate"
    return "sleep"


def available():
    """Can we actually do this here? Used to hide the option rather than fail late."""
    if SystemUtils.IS_WIN:
        return True
    if SystemUtils.IS_MAC:
        return bool(_which("pmset"))
    return bool(_which("systemctl"))


def _which(exe):
    import shutil
    return shutil.which(exe)


def _run(cmd):
    """Run a command detached enough that our own exit doesn't cancel it."""
    subprocess.run(cmd, check=True, timeout=30,
                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)


def suspend(action=SLEEP):
    """Put this machine to sleep now. Returns (ok, message).

    Never raises: the caller is a UI timer finishing a long, successful run, and
    an exception there would be a poor reward for it.
    """
    want_hib = (action == HIBERNATE)
    try:
        if SystemUtils.IS_WIN:
            return _suspend_windows(want_hib)
        if SystemUtils.IS_MAC:
            # `pmset sleepnow` needs no elevation for the console user.
            _run(["pmset", "sleepnow"])
            return True, "Sleeping."
        _run(["systemctl", "hibernate" if want_hib else "suspend"])
        return True, "Sleeping."
    except Exception as e:
        log_app(f"Could not put the computer to sleep: {e}", "ERROR")
        return False, str(e)


def _suspend_windows(want_hib):
    """SetSuspendState first, falling back to the shell.

    SetSuspendState(bHibernate, bForce, bWakeupEventsDisabled). Hibernating also
    needs the shutdown privilege and a hiberfil, neither guaranteed — so if it
    refuses we fall back, and finally settle for plain sleep rather than leaving
    the machine awake all night.
    """
    try:
        ok = ctypes.windll.powrprof.SetSuspendState(1 if want_hib else 0, 0, 0)
        if ok:
            return True, "Sleeping."
    except Exception as e:
        log_app(f"SetSuspendState unavailable ({e}) — falling back.", "WARNING")

    if want_hib:
        try:
            _run(["shutdown", "/h"])
            return True, "Hibernating."
        except Exception as e:
            log_app(f"Hibernate refused ({e}) — sleeping instead.", "WARNING")

    try:
        # The documented shell equivalent; the 0,1,0 forces it past apps that
        # would otherwise veto.
        _run(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"])
        return True, "Sleeping."
    except Exception as e:
        return False, str(e)
