"""Download and verify the pinned office2pdf release binary for the benchmark."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import tarfile
from urllib.request import urlopen

OFFICE2PDF_VERSION = '0.11.0'
# The statically linked musl build runs on older glibc hosts than the gnu build (GLIBC_2.43).
ARCHIVE = f'office2pdf-v{OFFICE2PDF_VERSION}-x86_64-unknown-linux-musl'
URL = f'https://github.com/developer0hye/office2pdf/releases/download/v{OFFICE2PDF_VERSION}/{ARCHIVE}.tar.gz'
SHA256 = 'a5eb3ee023c1e43e4d6e7303914aaee17492f87d0c9211d5d44d8358f2b5f2ed'
TARGET = Path(__file__).resolve().parents[3] / 'out' / 'benchmark-word' / 'office2pdf-bin' / 'office2pdf'


def main():
    with urlopen(URL, timeout=120) as response:
        data = response.read()
    if sha256(data).hexdigest() != SHA256:
        raise ValueError(f'SHA-256 mismatch for {URL}')
    with tarfile.open(fileobj=BytesIO(data)) as archive:
        binary = archive.extractfile(f'{ARCHIVE}/office2pdf').read()
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    TARGET.write_bytes(binary)
    TARGET.chmod(0o755)
    print(f'Installed office2pdf {OFFICE2PDF_VERSION} at {TARGET}')


if __name__ == '__main__':
    main()
