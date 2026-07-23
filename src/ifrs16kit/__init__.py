"""
ifrs16kit — IFRS 16 auditor re-performance (LIQUID edition).

Public API
----------
    import ifrs16kit as lease
    # or:  from ifrs16kit import (
        LeaseInputs, cross_check,
        build_template, read_template, build_calculation_workbook,
        COUNTRIES, FREQUENCIES,
    )

Command line
------------
    ifrs16kit           # guided interactive flow
    ifrs16kit --demo    # CALISMA golden case end-to-end
"""

from .core import (
    LeaseInputs,
    create_input,
    create_calculation,
    demo,
    cross_check,
    build_template,
    read_template,
    build_calculation_workbook,
    build_inputs_sheet,
    print_summary,
    COUNTRIES,
    FREQUENCIES,
    MAX_PERIODS,
    main,
    run,
)

__version__ = "1.0.2"

__all__ = [
    "LeaseInputs", "create_input", "create_calculation", "demo", "cross_check", "build_template", "read_template",
    "build_calculation_workbook", "build_inputs_sheet", "print_summary",
    "COUNTRIES", "FREQUENCIES", "MAX_PERIODS", "main", "run", "__version__",
]
