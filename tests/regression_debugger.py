r"""Live regression on the explicitly selected, disposable DebuggerTrial project.

Open tests/fixtures/debugger/DebuggerTrial.vbp in an owned IDE before running.
It sets session breakpoints and runs only that fixture. Use --project for an
explicitly selected copy. It never opens or switches projects.
"""
import argparse
import json
import pathlib
import time
import urllib.request
import uuid


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--port", type=int, default=8767)
    parser.add_argument("--output", type=pathlib.Path, required=True)
    parser.add_argument("--project", type=pathlib.Path,
                        default=pathlib.Path(__file__).resolve().parent /
                        "fixtures" / "debugger" / "DebuggerTrial.vbp",
                        help="Expected disposable fixture project (default: bundled fixture)")
    args = parser.parse_args()
    records = []
    request_id_prefix = uuid.uuid4().hex

    def tool(name, arguments=None):
        arguments = dict(arguments or {}, _target=args.pid)
        request = {"jsonrpc": "2.0", "id": "%s-%d" % (request_id_prefix, len(records) + 1), "method": "tools/call",
                   "params": {"name": name, "arguments": arguments}}
        started = time.monotonic()
        req = urllib.request.Request(
            "http://127.0.0.1:%d/rpc" % args.port,
            data=json.dumps(request).encode(), headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=12) as response:
            envelope = json.load(response)
        content = envelope.get("result", {}).get("content", [])
        result = json.loads(content[0]["text"]) if content else envelope
        records.append({"tool": name, "arguments": arguments,
                        "seconds": round(time.monotonic() - started, 3), "result": result})
        args.output.write_text(json.dumps(records, indent=2), encoding="utf-8")
        return result

    def wait_mode(expected):
        for _ in range(15):
            result = tool("ide_state")
            if result.get("mode") == expected:
                return
            time.sleep(0.1)
        raise AssertionError((expected, result))

    def execute(mode):
        result = tool("control_execution", {"mode": mode})
        assert result.get("ok") and result.get("executed"), result
        time.sleep(0.15)

    project = tool("get_project_info")
    assert project.get("name") == "DebuggerTrial", "Expected the loaded DebuggerTrial fixture: %r" % project
    path = pathlib.Path(project["projectPath"])
    assert path.resolve() == args.project.resolve(), "Refusing an unexpected project: %s" % path
    wait_mode("design")
    source = tool("get_component_code", {"name": "MTrial", "coords": "ide"})["text"]
    lines = source.splitlines()
    assert "    total = -999" in lines, "Expected the module/local shadowing fixture"
    breakpoint_line = lines.index("    Touch total") + 1
    for control in ("#2555", "#2554"):
        result = tool("command_bars", {"action": "execute", "bar": "View", "control": control})
        assert result.get("executed"), result
    tool("control_window", {"window": "Immediate", "action": "close"})
    tool("activate_component", {"name": "MTrial"})
    tool("native_breakpoint", {"action": "clear_all"})
    result = tool("native_breakpoint", {"action": "toggle", "component": "MTrial", "line": breakpoint_line})
    assert result.get("executed") and result.get("paneTextUnchanged"), result
    result = tool("run_until_break", {"timeout_ms": 3000})
    assert result.get("startedPending"), result
    wait_mode("break")
    stack = tool("get_callstack")
    assert [frame["proc"] for frame in stack["frames"]] == ["Inner", "Outer", "Main"], stack
    print("PASS: native breakpoint hit and call stack")

    result = tool("control_window", {"window": "Locals", "action": "show"})
    assert result.get("ok") and not result.get("windowModelTouched"), result
    assert tool("get_locals").get("reason") == "break_eval_requires_opt_in"
    result = tool("get_locals", {"allowUnsafeInBreak": True})
    assert result.get("ok"), result
    locals_ = {item["expr"]: item for item in result["locals"]}
    expected = {"inputValue": ("12", "Long"), "copy": ("12", "Long"),
                "label": ("ready", "String"), "total": ("36", "Long"),
                "emptyText": ("", "String"), "emptyVariant": ("", "Empty"), "fraction": ("1.25", "Double"),
                "flag": ("True", "Boolean")}
    for name, (value, type_name) in expected.items():
        assert locals_[name]["value"] == value, (name, locals_[name])
        assert locals_[name]["typeName"] == type_name, (name, locals_[name])
        assert locals_[name]["valueRead"], (name, locals_[name])
    assert locals_["label"]["declaredType"] == "String", locals_["label"]
    result = tool("eval_at_frame", {"expr": "total * 2", "allowUnsafeEval": True})
    assert result.get("ok") and result.get("value") == "72", result
    result = tool("eval_at_frame", {"expr": "Probe()", "allowUnsafeEval": True})
    assert result.get("ok") and result.get("value") == "1", result
    result = tool("immediate_eval", {"expression": "evaluations"})
    assert result.get("ok") and result.get("value") == "1", result
    result = tool("eval_at_frame", {"expr": "total", "frame_index": 1, "allowUnsafeEval": True})
    assert result.get("reason") == "frame_binding_forbidden", result
    print("PASS: ByVal/ByRef/Optional parameters, scalar types, empty string, local shadowing and evaluation")

    execute("step_over")
    result = tool("immediate_eval", {"expression": "total"})
    assert result.get("ok") and result.get("value") == "37", result
    execute("step_out")
    stack = tool("get_callstack")
    assert [frame["proc"] for frame in stack["frames"]] == ["Outer", "Main"], stack
    execute("continue")
    wait_mode("design")
    print("PASS: Step Over, Step Out and Continue")

    result = tool("run_until_break", {"timeout_ms": 3000})
    assert result.get("startedPending"), result
    wait_mode("break")
    execute("step_into")
    stack = tool("get_callstack")
    assert stack["frames"][0]["proc"] == "Touch", stack
    execute("step_out")
    execute("continue")
    wait_mode("design")
    after = tool("get_component_code", {"name": "MTrial", "coords": "ide"})["text"]
    assert after == source, "Debugger operations changed fixture source"
    tool("activate_component", {"name": "MTrial"})
    tool("native_breakpoint", {"action": "clear_all"})
    print("PASS: Step Into, repeated run and unchanged source")


if __name__ == "__main__":
    main()
