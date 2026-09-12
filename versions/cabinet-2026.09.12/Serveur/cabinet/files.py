"""Copies immuables sur volume local NAS, avant validation SQL des publications."""
from pathlib import Path, PureWindowsPath
import hashlib
import os
import tempfile
import zipfile
from .domain import Refus

MAX_FILE = 32 * 1024 * 1024

class Documents:
    def __init__(self, root: Path, unc: str):
        self.root = root.resolve()
        self.unc = PureWindowsPath(unc)
        if not str(unc).startswith('\\\\'):
            raise ValueError('CABINET_UNC doit etre le partage NAS.')
        self.objects = self.root / 'Documents'

    def resoudre(self, value: str) -> Path:
        path = PureWindowsPath(value)
        try:
            relative = path.relative_to(self.unc)
        except ValueError as exc:
            raise Refus('Document situe hors du partage du cabinet.', 422) from exc
        if '..' in relative.parts or any(':' in x for x in relative.parts):
            raise Refus('Chemin de document interdit.', 422)
        resolved = self.root.joinpath(*relative.parts).resolve()
        if not resolved.is_relative_to(self.root):
            raise Refus('Document situe hors du volume NAS.', 422)
        return resolved

    def conserver(self, source: str, extension: str) -> tuple[str, str]:
        path = self.resoudre(source)
        if path.suffix.lower() != extension or not path.is_file() or not 0 < path.stat().st_size <= MAX_FILE:
            raise Refus('Document absent, vide ou trop volumineux.', 422)
        # Lire une fois : le contenu valide est exactement celui copie.
        with path.open('rb') as f: data = f.read(MAX_FILE + 1)
        if len(data) > MAX_FILE: raise Refus('Document trop volumineux.', 422)
        if extension == '.pdf' and not data.startswith(b'%PDF-'):
            raise Refus('PDF invalide.', 422)
        if extension == '.docx':
            import io
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    if len(z.infolist()) > 4096 or sum(x.file_size for x in z.infolist()) > 128 * 1024 * 1024:
                        raise ValueError()
                    if 'word/document.xml' not in z.namelist() or z.testzip():
                        raise ValueError()
                    if any('vbaProject' in x for x in z.namelist()): raise ValueError()
            except (zipfile.BadZipFile, ValueError) as exc:
                raise Refus('DOCX invalide ou contenant des macros.', 422) from exc
        sha = hashlib.sha256(data).hexdigest()
        self.objects.mkdir(parents=True, exist_ok=True)
        target = self.objects / (sha + extension)
        if not target.exists():
            fd, tmp = tempfile.mkstemp(dir=self.objects, suffix='.tmp')
            try:
                with os.fdopen(fd, 'wb') as f:
                    f.write(data); f.flush(); os.fsync(f.fileno())
                os.replace(tmp, target)
                if hasattr(os, 'O_DIRECTORY'):
                    fd = os.open(self.objects, os.O_DIRECTORY)
                    try: os.fsync(fd)
                    finally: os.close(fd)
            finally:
                if os.path.exists(tmp): os.unlink(tmp)
        elif hashlib.sha256(target.read_bytes()).hexdigest() != sha:
            raise Refus('Archive NAS alteree : publication arretee.', 503)
        return str(self.unc / 'Documents' / target.name), sha
