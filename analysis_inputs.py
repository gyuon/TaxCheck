"""페이지별 입력 준비. 진행 중인 작업을 공유하고 오래된 결과는 재사용하지 않는다."""
import asyncio
import time


class PreparedInput:
    def __init__(self, loader, *, ttl=None, clock=time.monotonic):
        self.loader = loader
        self.ttl = ttl
        self.clock = clock
        self._key = None
        self._task = None
        self._loaded_at = None

    def start(self, key, *args, force=False):
        task = self._task
        if not force and key == self._key and task is not None:
            if not task.done():
                return task
            if not task.cancelled() and task.exception() is None:
                if self.ttl is None or self.clock() - self._loaded_at < self.ttl:
                    return task

        async def load():
            result = await self.loader(*args)
            if self._task is asyncio.current_task():
                self._loaded_at = self.clock()
            return result

        self._key = key
        self._loaded_at = None
        self._task = asyncio.create_task(load())
        # 미리 준비하다 실패한 경우에도 예외를 회수한다. get()에서는 그대로 전달된다.
        self._task.add_done_callback(lambda done: None if done.cancelled() else done.exception())
        return self._task

    async def get(self, key, *args):
        return await asyncio.shield(self.start(key, *args))
