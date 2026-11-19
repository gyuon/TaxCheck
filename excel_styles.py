"""워크시트 안에서 완성된 셀 서식을 재사용한다."""

from copy import copy


class StyleTemplates:
    """key는 적용할 서식 조합을 식별하며, 캐시는 통합문서마다 새로 만든다."""

    def __init__(self):
        self._templates = {}
        self._workbook = None

    def apply_with(self, cell, key, configure):
        workbook = cell.parent.parent
        if self._workbook is None:
            self._workbook = workbook
        elif self._workbook is not workbook:
            raise ValueError('서식 템플릿은 같은 통합문서 안에서만 사용할 수 있습니다.')
        # pandas가 미리 지정한 인덱스 서식 등도 보존한다.
        base = tuple(cell._style) if cell._style is not None else None
        cache_key = (key, base)
        if cache_key not in self._templates:
            configure(cell)
            self._templates[cache_key] = copy(cell._style)
        # StyleArray는 변경 가능하므로 셀마다 복사해 이후 서식 수정의 전파를 막는다.
        # openpyxl 내부 서식 접근은 이 모듈로 한정한다.
        cell._style = copy(self._templates[cache_key])

    def apply(self, cell, key, **components):
        def configure(target):
            for name, value in components.items():
                setattr(target, name, value)

        self.apply_with(cell, key, configure)
