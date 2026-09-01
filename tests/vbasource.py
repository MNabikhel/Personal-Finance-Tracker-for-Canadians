"""A reader for the VBA sources in ``vba/``.

Most of the workbook's code needs Excel's object model and so cannot be run
outside Excel.  What can still be checked without running it is that the code
agrees with itself: that every qualified call has something to call and hands
it the right number of arguments, that every constant and class member exists,
and that the sheet and table names the macros use are the ones the builder
writes.  Those are the mistakes a refactor makes, and they are exactly what a
compiler would catch if one were available here.

The parsing is deliberately shallow.  It understands declarations, line
continuations, comments and string literals, which is enough to resolve names;
it does not try to understand expressions.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

VBA_DIR = Path(__file__).resolve().parent.parent / "vba"

CLASS_PREFIX = "cls"

PROCEDURE = re.compile(
    r"^(?:(?P<scope>Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"(?P<name>[A-Za-z_]\w*)\s*(?:\((?P<args>.*)\))?\s*(?:As\s+\w+)?\s*$",
    re.IGNORECASE)

CONSTANT = re.compile(
    r"^(?:(?P<scope>Public|Private)\s+)?Const\s+(?P<name>[A-Za-z_]\w*)",
    re.IGNORECASE)

PUBLIC_FIELD = re.compile(
    r"^Public\s+(?P<name>[A-Za-z_]\w*)\s+As\s+(?P<type>[\w.]+)\s*$",
    re.IGNORECASE)

# "Dim txn As clsTxn", "ByVal profile As clsProfile", "Set rule = New clsRule".
TYPED_NAME = re.compile(
    r"\b(?:Dim|Private|Public|Static|ByVal|ByRef)\s+(?P<name>[A-Za-z_]\w*)\s*(?:\(\))?"
    r"\s+As\s+(?P<type>cls[A-Za-z_]\w*)\b",
    re.IGNORECASE)

OPTION_EXPLICIT = re.compile(r"^Option\s+Explicit\s*$", re.IGNORECASE | re.MULTILINE)
NAME_ATTRIBUTE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.MULTILINE)

UPPER_CONSTANT = re.compile(r"\b([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+)\b")

# Anything that is a qualifier but not one of our modules.
HOST_OBJECTS = {"Application", "Err", "Debug", "Excel", "VBA", "ThisWorkbook",
                "Selection", "Me"}


@dataclass
class Procedure:
    module: str
    name: str
    kind: str
    scope: str
    required: int
    limit: Optional[int]        # None when the last argument is a ParamArray

    @property
    def public(self) -> bool:
        return self.scope.lower() != "private"

    def accepts(self, count: int) -> bool:
        if count < self.required:
            return False
        return self.limit is None or count <= self.limit


@dataclass
class Module:
    name: str
    path: Path
    text: str                   # as written
    code: str                   # comments and string contents removed
    lines: List[str] = field(default_factory=list)
    procedures: Dict[str, Procedure] = field(default_factory=dict)
    constants: Dict[str, str] = field(default_factory=dict)
    fields: Dict[str, str] = field(default_factory=dict)

    @property
    def is_class(self) -> bool:
        return self.path.suffix == ".cls"

    def member(self, name: str) -> bool:
        key = name.lower()
        return key in self.procedures or key in self.fields or key in self.constants


def strip_line(line: str) -> str:
    """The line with string contents emptied and any comment removed."""
    out: List[str] = []
    index = 0
    length = len(line)
    while index < length:
        char = line[index]
        if char == '"':
            index += 1
            while index < length:
                if line[index] == '"':
                    if index + 1 < length and line[index + 1] == '"':
                        index += 2
                        continue
                    break
                index += 1
            out.append('""')
            index += 1
        elif char == "'":
            break
        else:
            out.append(char)
            index += 1
    return "".join(out).rstrip()


def joined_lines(text: str) -> List[str]:
    """Logical lines: comments gone, string contents gone, continuations joined."""
    out: List[str] = []
    pending = ""
    for raw in text.splitlines():
        line = strip_line(raw)
        if re.search(r"\s_$", line):
            pending += line[:-1]
            continue
        combined = (pending + line).strip()
        pending = ""
        if combined:
            out.append(re.sub(r"\s+", " ", combined))
    if pending.strip():
        out.append(re.sub(r"\s+", " ", pending.strip()))
    return out


def split_arguments(text: str) -> List[str]:
    """Top level comma separated pieces of an argument list."""
    pieces: List[str] = []
    depth = 0
    current = ""
    for char in text:
        if char in "([":
            depth += 1
        elif char in ")]":
            depth -= 1
        if char == "," and depth == 0:
            pieces.append(current.strip())
            current = ""
        else:
            current += char
    if current.strip():
        pieces.append(current.strip())
    return pieces


def _arity(argument_text: Optional[str]) -> Tuple[int, Optional[int]]:
    if not argument_text or not argument_text.strip():
        return 0, 0
    arguments = split_arguments(argument_text)
    required = 0
    unbounded = False
    for argument in arguments:
        lowered = argument.lower()
        if lowered.startswith("paramarray"):
            unbounded = True
        elif lowered.startswith("optional"):
            pass
        else:
            required += 1
    if unbounded:
        return required, None
    return required, len(arguments)


def load(directory: Path = VBA_DIR) -> Dict[str, Module]:
    """Every module in ``vba/``, parsed, keyed by its declared name."""
    modules: Dict[str, Module] = {}
    for path in sorted(directory.iterdir()):
        if path.suffix not in (".bas", ".cls"):
            continue
        text = path.read_text(encoding="utf-8")
        match = NAME_ATTRIBUTE.search(text)
        name = match.group(1) if match else path.stem
        module = Module(name=name, path=path, text=text,
                        code="\n".join(joined_lines(text)))
        module.lines = joined_lines(text)
        _parse(module)
        modules[name] = module
    return modules


def _parse(module: Module) -> None:
    for line in module.lines:
        found = PROCEDURE.match(line)
        if found and not line.lower().startswith(("exit ", "end ")):
            required, limit = _arity(found.group("args"))
            kind = re.sub(r"\s+", " ", found.group("kind"))
            module.procedures[found.group("name").lower()] = Procedure(
                module=module.name, name=found.group("name"), kind=kind,
                scope=found.group("scope") or "Public",
                required=required, limit=limit)
            continue

        found = CONSTANT.match(line)
        if found:
            module.constants[found.group("name").lower()] = found.group("scope") or "Private"
            continue

        found = PUBLIC_FIELD.match(line)
        if found:
            module.fields[found.group("name").lower()] = found.group("type")


def typed_locals(module: Module) -> List[Tuple[str, Dict[str, str]]]:
    """Each procedure's lines, paired with the cls-typed names in scope there.

    Scoped per procedure because the same name means different things in
    different procedures - ``out`` is a clsProfile in one and a Collection in
    the next - and a module-wide map would confuse the two.
    """
    out: List[Tuple[str, Dict[str, str]]] = []
    current: Dict[str, str] = {}
    module_level: Dict[str, str] = {}
    in_procedure = False

    for line in module.lines:
        if PROCEDURE.match(line) and not line.lower().startswith(("exit ", "end ")):
            current = dict(module_level)
            in_procedure = True
        elif re.match(r"^End (Sub|Function|Property)\b", line, re.IGNORECASE):
            in_procedure = False
            current = dict(module_level)
        for found in TYPED_NAME.finditer(line):
            current[found.group("name").lower()] = found.group("type")
            if not in_procedure:
                module_level[found.group("name").lower()] = found.group("type")
        out.append((line, dict(current)))
    return out


@dataclass
class Reference:
    module: str                 # where the reference was written
    line: str
    qualifier: str              # the module or variable it was written against
    name: str
    arguments: Optional[int]    # None when the count cannot be read off

    def __str__(self) -> str:
        return f"{self.module}: {self.qualifier}.{self.name} in {self.line!r}"


def references(module: Module, qualifiers: Set[str]) -> List[Reference]:
    """Every ``qualifier.Name`` in ``module`` for a qualifier we know about."""
    out: List[Reference] = []
    pattern = re.compile(r"\b(" + "|".join(sorted(map(re.escape, qualifiers)))
                         + r")\.([A-Za-z_]\w*)")
    for line in module.lines:
        for found in pattern.finditer(line):
            qualifier, name = found.group(1), found.group(2)
            out.append(Reference(module=module.name, line=line,
                                 qualifier=qualifier, name=name,
                                 arguments=_argument_count(line, found.end())))
    return out


def local_references(module: Module) -> List[Reference]:
    """Every unqualified call in ``module`` to something ``module`` declares.

    A procedure calls its own module's procedures by bare name, so these are
    the calls that qualified-reference scanning cannot see - and they are the
    majority of the calls in the codebase.
    """
    out: List[Reference] = []
    names = {procedure.name for procedure in module.procedures.values()}
    if not names:
        return out
    pattern = re.compile(r"(?<![\w.])(" + "|".join(sorted(map(re.escape, names)))
                         + r")\b")
    for line in module.lines:
        if PROCEDURE.match(line) and not line.lower().startswith(("exit ", "end ")):
            continue                    # the declaration itself
        for found in pattern.finditer(line):
            count = _argument_count(line, found.end())
            if count is None:
                continue
            out.append(Reference(module=module.name, line=line,
                                 qualifier=module.name, name=found.group(1),
                                 arguments=count))
    return out


def _argument_count(line: str, after: int) -> Optional[int]:
    """How many arguments the call at ``after`` was given, if it can be told."""
    rest = line[after:]
    stripped = rest.lstrip()
    if stripped.startswith("="):
        return None                     # an assignment or a comparison
    if stripped.startswith("("):
        offset = len(rest) - len(stripped)
        depth = 0
        for index, char in enumerate(rest[offset:], start=offset):
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    return len(split_arguments(rest[offset + 1:index]))
        return None                     # unbalanced: leave it alone

    # A Sub called as a statement takes its arguments without parentheses, but
    # only where the call is the whole statement; anywhere else the name is
    # being read rather than called and there is nothing to count.
    start = line.rfind(" ", 0, after)
    head = line[:after]
    if re.fullmatch(r"(?:Call\s+)?[\w.]+", head, re.IGNORECASE):
        return len(split_arguments(stripped))
    del start
    return None
