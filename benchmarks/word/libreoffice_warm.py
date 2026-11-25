"""Measure consecutive PDF completions from one CLI batch (LibreOffice or office2pdf)."""
import ctypes
import os
from pathlib import Path
import select
from struct import Struct
import subprocess
from time import perf_counter

IN_CLOSE_WRITE = 0x00000008
IN_MOVED_TO = 0x00000080
EVENT = Struct('iIII')


def libreoffice_command(workspace, output, sources):
    return [
        'soffice', '-env:UserInstallation=' + (workspace / 'profile').as_uri(),
        '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', output, *sources,
    ]


def batch_times(jobs, workspace, command=libreoffice_command, label='LibreOffice', extension='pdf', env=None):
    """Time each job; a batch that stops early returns times for the jobs it completed."""
    output = workspace / 'pdfs'
    output.mkdir()
    sources = []
    for index, (source, _) in enumerate(jobs):
        # The CLI derives the output filename from the input basename.
        alias = workspace / f'{index}.docx'
        alias.hardlink_to(source)
        sources.append(alias)
    targets = {f'{i}.{extension}': Path(target) for i, (_, target) in enumerate(jobs)}
    completed = {}

    libc = ctypes.CDLL(None, use_errno=True)
    watcher = libc.inotify_init1(os.O_CLOEXEC | os.O_NONBLOCK)
    if watcher < 0:
        raise OSError(ctypes.get_errno(), 'inotify_init1')
    try:
        if libc.inotify_add_watch(watcher, os.fsencode(output), IN_CLOSE_WRITE | IN_MOVED_TO) < 0:
            raise OSError(ctypes.get_errno(), 'inotify_add_watch')
        with (workspace / 'batch.log').open('w+') as log:
            start = perf_counter()
            with subprocess.Popen(command(workspace, output, sources), stdout=log, stderr=log, env=env) as process:
                try:
                    while len(completed) < len(targets):
                        remaining = 300 - (perf_counter() - start)
                        if remaining <= 0:
                            raise TimeoutError(f'{label} warm batch timed out')
                        if select.select([watcher], [], [], min(0.1, remaining))[0]:
                            now = perf_counter()
                            events = os.read(watcher, 65536)
                            offset = 0
                            while offset < len(events):
                                _, _, _, length = EVENT.unpack_from(events, offset)
                                offset += EVENT.size
                                name = os.fsdecode(events[offset:offset + length].rstrip(b'\0'))
                                offset += length
                                if name in targets and name not in completed:
                                    completed[name] = now
                        elif process.poll() is not None:
                            break
                    status = process.wait(timeout=10)
                except BaseException:
                    process.kill()
                    raise
            # A batch that stops early returns the documents it completed, so the caller can
            # resume after the document it stopped on. Out-of-order completions are an error.
            if list(completed) != list(targets)[:len(completed)]:
                log.seek(0)
                raise RuntimeError(f'{label} warm batch did not complete in input order: {log.read()}')
            if len(completed) < len(targets):
                print(f'{label} warm batch stopped after {len(completed)} of {len(targets)} documents (status {status})',
                      flush=True)
    finally:
        os.close(watcher)

    for name in completed:
        (output / name).replace(targets[name])
    timestamps = [start, *completed.values()]
    return [(end - begin) * 1000 for begin, end in zip(timestamps, timestamps[1:])]
