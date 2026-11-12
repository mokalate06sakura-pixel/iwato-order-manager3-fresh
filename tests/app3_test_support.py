"""UIを実行せずapp3.pyの関数と定数を読み込む。"""
import ast
import io
from pathlib import Path
import types

APP_PATH = Path(__file__).resolve().parents[1] / "app3.py"


def load_logic():
    source = ast.parse(APP_PATH.read_text(encoding="utf-8"))
    nodes = [
        node for node in source.body
        if isinstance(node, (ast.Import, ast.ImportFrom, ast.FunctionDef))
        or (isinstance(node, ast.Assign) and all(
            isinstance(target, ast.Name) and target.id.isupper() for target in node.targets
        ))
    ]
    module = types.ModuleType("app3_logic")
    module.__file__ = str(APP_PATH)
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(APP_PATH), "exec"), module.__dict__)
    return module


class Upload(io.BytesIO):
    def __init__(self, payload, name):
        super().__init__(payload)
        self.name = name
        self.size = len(payload)

    @classmethod
    def from_path(cls, path):
        return cls(path.read_bytes(), path.name)
