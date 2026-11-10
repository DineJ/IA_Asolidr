"""Check pipeline orchestration and failure handling without real subprocesses."""
import subprocess

import pytest

import main


def test_run_step_uses_current_python(monkeypatch):
    calls = []
    monkeypatch.setattr(subprocess, 'run', lambda cmd, **kwargs: calls.append((cmd, kwargs)))
    main.run_step('extract.py', ['--input', 'data/raw'])
    assert calls[0][0][0] == main.sys.executable
    assert calls[0][0][1].endswith('extract.py')
    assert calls[0][1]['check'] is True


def test_run_step_propagates_failure(monkeypatch):
    def fail(cmd, **kwargs):
        raise subprocess.CalledProcessError(2, cmd)
    monkeypatch.setattr(subprocess, 'run', fail)
    with pytest.raises(subprocess.CalledProcessError):
        main.run_step('extract.py', [])
