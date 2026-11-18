"""Check procedure names and Dim declarations without interpreting the module's code."""

from itertools import groupby
import re
import warnings

from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter import tokenize
from pyopenvba.interpreter._preprocess import strip_directives


# MS-VBAL 3.3.5.2 defines the words excluded from an ordinary IDENTIFIER.
# https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/7df907cb-ab6c-40d3-aa81-272742ce00c3
RESERVED = frozenset("""
    abs addressof and any array as attribute boolean byref byte byval call case cbool cbyte ccur
    cdate cdbl cdec cdecl cint circle clng clnglng clngptr close const csng cstr currency cvar cverr
    date debug decimal declare defbool defbyte defcur defdate defdbl defdec defint deflng deflnglng
    deflngptr defobj defsng defstr defvar dim do doevents double each else elseif empty end endif
    enum eqv erase event exit false fix for friend function get global gosub goto if imp implements
    in input inputb int integer is lbound len lenb let like lineinput lock long longlong longptr loop
    lset me mod new next not nothing null on open option optional or paramarray preserve print private
    pset public put raiseevent redim rem resume return rset scale seek select set sgn shared single
    spc static stop string sub tab then to true type typeof ubound unlock until variant vb_base
    vb_control vb_creatable vb_customizable vb_description vb_exposed vb_ext_key vb_globalnamespace
    vb_helpid vb_invoke_func vb_invoke_property vb_invoke_propertyput vb_invoke_propertyputref
    vb_memberflags vb_name vb_predeclaredid vb_procdata vb_templatederived vb_usermemid
    vb_vardescription vb_varhelpid vb_varmemberflags vb_varprocdata vb_varusermemid wend while
    with withevents write xor
""".split())
DIRECTIVE = re.compile(r'^\s*#\s*(if\b|end\s*if\b)', re.IGNORECASE)


def unconditional_source(source):
    """Retain unconditional declarations if a custom directive expression exceeds the preprocessor."""
    depth = 0
    lines = []
    for line in source.splitlines():
        directive = DIRECTIVE.match(line)
        if directive:
            depth += 1 if directive[1].lower() == 'if' else -1
            if depth < 0:
                raise ValueError('Unmatched #End If in VBA conditional declarations.')
        lines.append('' if directive or depth else line)
    if depth:
        raise ValueError('Unclosed #If in VBA conditional declarations.')
    return '\n'.join(lines)


def declaration_sources(source):
    if not any(DIRECTIVE.match(line) for line in source.splitlines()):
        return (source,)
    # VBA7 Mac, 32-bit Windows and 64-bit Windows each compile one side of platform directives.
    profiles = ({'Mac': True, 'MacOffice16': True, 'Win32': False, 'Win64': False},
                {'Win64': False}, {'Win64': True})
    try:
        return tuple(dict.fromkeys(strip_directives(source, constants=profile) for profile in profiles))
    except VBACompileError as error:
        # The pinned preprocessor supports fewer constant expressions than Excel.
        plain_source = unconditional_source(source)
        warnings.warn(f'Conditional VBA declarations require Excel verification ({error}).', stacklevel=2)
        return (plain_source,)


def check_dim(tokens):
    # The first token and those following top-level commas are declaration names.
    depth = 0
    want_name = True
    for token in tokens:
        if want_name:
            if token.kind == 'ident' and token.lower == 'withevents':
                continue
            name = token.lower.rstrip('%&^!#@$')
            if token.kind == 'ident' and name in RESERVED:
                raise ValueError(f'Line {token.line}: reserved VBA variable name {token.text!r}.')
            want_name = False
        if token.kind == 'op':
            if token.text == '(':
                depth += 1
            elif token.text == ')':
                depth -= 1
            elif token.text == ',' and depth == 0:
                want_name = True


def audit_vba_declarations(source):
    """Catch duplicate procedures and bare reserved Dim names in each selected compilation branch."""
    for active in declaration_sources(source):
        procedures = {}
        member_block = None
        tokens = tokenize(active, retain_errors=True)
        for boundary, group in groupby(tokens, lambda token: token.kind in ('eos', 'eof')):
            if boundary:
                continue
            statement = list(group)
            words = [token.lower if token.kind == 'ident' else '' for token in statement]
            if member_block:
                if words == ['end', member_block]:
                    member_block = None
                continue
            index = 0
            while index < len(words) and words[index] in ('public', 'private', 'friend', 'static', 'declare', 'ptrsafe'):
                index += 1
            if index == len(words):
                continue
            declaration = words[index]
            if declaration in ('type', 'enum'):
                member_block = declaration
                continue
            if declaration == 'dim':
                first = index + 1
                if words[first:first + 1] == ['shared']:
                    first += 1
                check_dim(statement[first:])
                continue
            if declaration in ('sub', 'function'):
                kind = 'procedure'
                index += 1
            elif declaration == 'property' and index + 1 < len(words) and words[index + 1] in ('get', 'let', 'set'):
                kind = words[index + 1]
                index += 2
            else:
                continue
            if index >= len(statement) or statement[index].kind not in ('ident', 'bracket'):
                continue
            token = statement[index]
            name = token.lower.rstrip('%&^!#@$') if token.kind == 'ident' else token.lower
            previous = procedures.setdefault(name, {})
            if previous and (kind in previous or kind == 'procedure' or 'procedure' in previous):
                first = min(previous.values())
                raise ValueError(f'Line {token.line}: duplicate VBA procedure {name!r} (first at line {first}).')
            previous[kind] = token.line
