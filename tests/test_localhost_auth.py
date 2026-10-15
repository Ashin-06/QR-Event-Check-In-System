"""Exercise the actual host check without starting the eventlet application."""
import ast
from pathlib import Path
from types import SimpleNamespace

import pytest


def localhost_check(host, remote='127.0.0.1', forwarded=None):
    source = Path(__file__).resolve().parents[1]/'app.py'
    module = ast.parse(source.read_text(encoding='utf-8'))
    function = next(node for node in module.body if isinstance(node, ast.FunctionDef) and node.name == 'is_localhost_request')
    request = SimpleNamespace(host=host, remote_addr=remote,
                              headers={'X-Forwarded-For':forwarded} if forwarded else {})
    namespace = {'request':request}
    exec(compile(ast.Module(body=[function], type_ignores=[]), str(source), 'exec'), namespace)
    return namespace['is_localhost_request']()


@pytest.mark.parametrize('host', ['localhost','localhost:5001','127.0.0.1:5001','[::1]:5001'])
def test_exact_loopback_hosts(host):
    assert localhost_check(host)


@pytest.mark.parametrize('host', ['localhost.attacker.example','127.0.0.1.attacker.example',
                                'notlocalhost.example','attacker-localhost.example','[broken',
                                'attacker@localhost','localhost:invalid','localhost/path'])
def test_deceptive_hostnames_do_not_bypass_dashboard_auth(host):
    assert not localhost_check(host)


def test_remote_and_forwarded_requests_do_not_bypass_dashboard_auth():
    assert not localhost_check('localhost:5001', remote='192.0.2.1')
    assert not localhost_check('localhost:5001', forwarded='192.0.2.1')
