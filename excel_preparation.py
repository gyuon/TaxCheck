from dataclasses import dataclass
from enum import Enum


class ExcelPreparationStatus(str, Enum):
    IDLE = "idle"
    GENERATING = "generating"
    READY = "ready"
    FAILED = "failed"


@dataclass(frozen=True)
class DownloadButtonPresentation:
    text: str
    icon: str
    disabled: bool
    pulsing: bool


class ExcelPreparation:
    def __init__(self) -> None:
        self.status = ExcelPreparationStatus.IDLE
        self.data: bytes | None = None
        self.error: str | None = None
        self._generation = 0

    @property
    def presentation(self) -> DownloadButtonPresentation:
        if self.status == ExcelPreparationStatus.GENERATING:
            return DownloadButtonPresentation("결과 파일 준비 중", "sym_o_overview", True, True)
        if self.status == ExcelPreparationStatus.READY:
            return DownloadButtonPresentation("결과 파일 다운로드", "download", False, False)
        if self.status == ExcelPreparationStatus.FAILED:
            return DownloadButtonPresentation("파일 준비 다시 시도", "refresh", False, False)
        return DownloadButtonPresentation("결과 파일 준비", "hourglass_empty", True, False)

    def start(self) -> int:
        self._generation += 1
        self.status = ExcelPreparationStatus.GENERATING
        self.data = None
        self.error = None
        return self._generation

    def complete(self, generation: int, data: bytes) -> bool:
        if generation != self._generation:
            return False
        self.status = ExcelPreparationStatus.READY
        self.data = data
        self.error = None
        return True

    def fail(self, generation: int, error: str) -> bool:
        if generation != self._generation:
            return False
        self.status = ExcelPreparationStatus.FAILED
        self.data = None
        self.error = error
        return True

    def invalidate(self) -> None:
        self._generation += 1
        self.status = ExcelPreparationStatus.IDLE
        self.data = None
        self.error = None
