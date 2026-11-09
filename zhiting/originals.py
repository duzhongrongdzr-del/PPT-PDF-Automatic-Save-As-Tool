"""Verify temporary-file/original-file relationships by bytes, never by guesswork."""
from pathlib import Path
import re
from .archive import EXTENSIONS, digest_file, signature, temporary_name


def title_matches_name(title, filename):
    name = Path(filename).name
    stem = Path(filename).stem
    reader = r"(?:WPS|Microsoft PowerPoint|PowerPoint|Adobe|Acrobat|SumatraPDF|Foxit|福昕|PDF-XChange|Microsoft Edge|Google Chrome|Mozilla Firefox)"
    tail = r"(?:\s*(?:\[[^\]]*\]|\([^)]*\)|（[^）]*）))*\s*(?:[-—–]\s*(?:[^-—–]{1,80}\s[-—–]\s)?" + reader + r".*)?"
    if Path(name).suffix.lower() in EXTENSIONS:
        prefixes = [re.escape(name), re.escape(stem)]
    else:
        # Recent shortcuts often omit the document extension in their own name.
        extensions = "|".join(re.escape(ext) for ext in sorted(EXTENSIONS))
        prefixes = [re.escape(name) + f"(?:{extensions})?"]
    return any(re.fullmatch(prefix + tail, title.strip(), flags=re.IGNORECASE)
               for prefix in prefixes)


def stable_digest(path, cache):
    before = signature(path)
    key = str(path)
    if key in cache and cache[key][0] == before:
        return cache[key][1]
    digest = digest_file(path)
    if signature(path) != before:
        raise OSError("文件在原名校验期间改变")
    cache[key] = (before, digest)
    if len(cache) > 256:
        cache.pop(next(iter(cache)))
    return digest


def prefer_originals(paths, candidates=(), cache=None):
    cache = {} if cache is None else cache
    normal = {p: method for p, method in paths.items() if not temporary_name(p)}
    temporary = [p for p in paths if temporary_name(p)]
    possible = list(dict.fromkeys([*normal, *(p for p in candidates if not temporary_name(p))]))
    unresolved = []
    recovered = []
    for temp in temporary:
        matches = []
        try:
            size = signature(temp)[0]
            for candidate in possible:
                try:
                    if signature(candidate)[0] == size and stable_digest(candidate, cache) == stable_digest(temp, cache):
                        matches.append(candidate)
                except OSError:
                    continue
        except OSError:
            continue
        # Existing document paths outrank history. Ambiguous historical names are not guessed.
        opened = [p for p in matches if p in normal]
        if opened:
            recovered.append({"temporary": temp, "original": opened[0]})
        elif matches and len({Path(p).name.casefold() for p in matches}) == 1:
            normal[matches[0]] = "原路径（与临时文件内容一致）"
            recovered.append({"temporary": temp, "original": matches[0]})
        else:
            unresolved.append(temp)
    for temp in unresolved:
        normal[temp] = paths[temp]
    return normal, unresolved, recovered
