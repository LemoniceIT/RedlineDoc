"""Local application preferences, independent of document contents."""

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile

from i18n import LANGUAGES


@dataclass(frozen=True)
class Preferences:
    language: str = "en"
    ignore_header: bool = False
    ignore_footer: bool = False

    @classmethod
    def load(cls, path: Path):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return cls()
        if not isinstance(data, dict):
            return cls()
        language = data.get("language")
        return cls(
            language if isinstance(language, str) and language in LANGUAGES else "en",
            data.get("ignore_header") is True,
            data.get("ignore_footer") is True,
        )

    def save(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(
            dir=path.parent, mode="w", encoding="utf-8", delete=False,
        ) as stream:
            temporary_path = Path(stream.name)
            try:
                json.dump(asdict(self), stream)
                stream.close()
                temporary_path.replace(path)
            finally:
                stream.close()
                temporary_path.unlink(missing_ok=True)
