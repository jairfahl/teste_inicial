import importlib

MODULES = [
    "tecnoloc_reconciliation",
    "tecnoloc_reconciliation.cli",
    "tecnoloc_reconciliation.loader",
    "tecnoloc_reconciliation.reports",
]


def test_imports():
    for module_name in MODULES:
        importlib.import_module(module_name)
