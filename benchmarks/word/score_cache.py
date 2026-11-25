"""On-disk cache of per-document scores and rendered reference pages.

Each entry is keyed by a digest of everything its result depends on: the candidate PDF's SHA-256,
the reference digests, the rasteriser and text extractor versions, the font files Poppler can
substitute, the scoring settings and the SHA-256 of the scoring code. A PDF that has not changed is
therefore not rasterised or scored again, whichever engine wrote it.

Entries are published with a rename from a staging directory in the run's own work directory on
the same disk, so parallel scoring workers and an interrupted run never leave a half-written entry.
Delete the cache directory to reclaim its space; nothing else depends on it.
"""
from hashlib import file_digest, sha256
import json
import os
from pathlib import Path
import shutil
import subprocess
from shutil import which
from tempfile import NamedTemporaryFile, mkdtemp

# Raise when the layout or meaning of the entries changes.
FORMAT = 1


def digest(value):
    return sha256(json.dumps(value, sort_keys=True).encode()).hexdigest()


def file_sha256(path):
    with Path(path).open('rb') as file:
        return file_digest(file, 'sha256').hexdigest()


def code_digests(paths, root):
    return {Path(path).relative_to(root).as_posix(): file_sha256(path) for path in paths}


def rendering_files(env=None):
    """The pdftoppm binary and the font files and configuration Poppler can substitute fonts from,
    as (path, size, mtime_ns), and the fontconfig environment variables."""
    def lines(command):
        return subprocess.run(command, capture_output=True, text=True, check=True, env=env).stdout.splitlines()

    paths = {Path(which('pdftoppm')).resolve()}
    if which('fc-list'):
        paths.update(Path(name) for name in lines(['fc-list', '-f', '%{file}\n']) if name)
    if which('fc-conflist'):
        for line in lines(['fc-conflist']):
            if line.startswith('+ '):
                paths.add(Path(line[2:].split(':', 1)[0]))
    stats = []
    for path in sorted(paths):
        if path.is_file():
            stat = path.stat()
            stats.append((str(path), stat.st_size, stat.st_mtime_ns))
    font_env = [os.environ.get(key) for key in ('FONTCONFIG_FILE', 'FONTCONFIG_PATH', 'XDG_CONFIG_HOME')]
    return stats, font_env


class ScoreCache:
    """Results under directory/scores and reference pages under directory/references.

    staging is a directory of the current run on the same file system, where entries are written
    before they are renamed into place. With read=False every lookup misses, and entries are
    written again.
    """

    def __init__(self, directory, staging, *, read=True):
        self.directory = Path(directory)
        self.staging = Path(staging)
        self.read = read

    def _entry(self, key):
        return self.directory / 'scores' / key[:2] / f'{key}.json'

    def get(self, key):
        if not self.read:
            return None
        try:
            saved = json.loads(self._entry(key).read_text(encoding='utf-8'))
            return saved['result'] if saved['key'] == key else None
        except (OSError, ValueError, KeyError, TypeError):
            return None

    def put(self, key, result):
        target = self._entry(key)
        target.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile('w', encoding='utf-8', dir=self.staging, suffix='.json', delete=False) as file:
            staged = Path(file.name)
            try:
                json.dump({'format': FORMAT, 'key': key, 'result': result}, file)
                file.flush()
                os.fsync(file.fileno())
            except BaseException:
                file.close()
                staged.unlink(missing_ok=True)
                raise
        try:
            staged.replace(target)
        finally:
            staged.unlink(missing_ok=True)

    def _pages(self, entry):
        try:
            names = json.loads((entry / 'pages.json').read_text(encoding='utf-8'))
        except (OSError, ValueError):
            return None
        pages = [entry / name for name in names]
        return pages if pages and all(page.is_file() for page in pages) else None

    def reference_pages(self, key, render):
        """The cached page images for key, or those render(directory) writes to a fresh directory
        and returns, which are then cached."""
        entry = self.directory / 'references' / key
        pages = self._pages(entry) if self.read else None
        if pages is not None:
            return pages
        staged = Path(mkdtemp(prefix='reference-', dir=self.staging))
        try:
            rendered = render(staged)
            if not rendered:
                return []
            if any(Path(page).parent != staged for page in rendered):
                raise ValueError('Reference pages must be written to the directory they are given')
            (staged / 'pages.json').write_text(json.dumps([Path(page).name for page in rendered]), encoding='utf-8')
            entry.parent.mkdir(parents=True, exist_ok=True)
            try:
                staged.rename(entry)
            except OSError:
                # Another worker published the entry first, or an earlier one is invalid or being
                # replaced (read=False): keep a valid one, otherwise replace it.
                pages = self._pages(entry) if self.read else None
                if pages is not None:
                    return pages
                old = Path(mkdtemp(prefix='replaced-', dir=self.staging))
                entry.rename(old / 'entry')
                staged.rename(entry)
            return self._pages(entry)
        finally:
            shutil.rmtree(staged, ignore_errors=True)
