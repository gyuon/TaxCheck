import asyncio
import unittest
from unittest.mock import AsyncMock

from analysis_inputs import PreparedInput


class TestPreparedInput(unittest.IsolatedAsyncioTestCase):
    async def test_inflight_and_completed_requests_are_shared(self):
        release = asyncio.Event()
        async def load(value):
            await release.wait()
            return value
        loader = AsyncMock(side_effect=load)
        cache = PreparedInput(loader)
        cache.start(1, 'file')
        a = asyncio.create_task(cache.get(1, 'file'))
        b = asyncio.create_task(cache.get(1, 'file'))
        release.set()
        self.assertEqual(await asyncio.gather(a, b), ['file', 'file'])
        self.assertEqual(await cache.get(1, 'file'), 'file')
        loader.assert_awaited_once()

    async def test_file_replacement_does_not_publish_old_completion(self):
        old_release = asyncio.Event()
        async def load(value):
            if value == 'old':
                await old_release.wait()
            return value
        cache = PreparedInput(load)
        old = cache.start(1, 'old')
        self.assertEqual(await cache.get(2, 'new'), 'new')
        old_release.set()
        self.assertEqual(await old, 'old')
        self.assertEqual(await cache.get(2, 'new'), 'new')

    async def test_expiration_url_change_and_failure_retry(self):
        now = [0]
        loader = AsyncMock(side_effect=['first', 'fresh', RuntimeError('offline'), 'retry'])
        cache = PreparedInput(loader, ttl=300, clock=lambda: now[0])
        self.assertEqual(await cache.get('url1', 'url1'), 'first')
        now[0] = 299
        self.assertEqual(await cache.get('url1', 'url1'), 'first')
        now[0] = 300
        self.assertEqual(await cache.get('url1', 'url1'), 'fresh')
        with self.assertRaisesRegex(RuntimeError, 'offline'):
            await cache.get('url2', 'url2')
        self.assertEqual(await cache.get('url2', 'url2'), 'retry')

    async def test_cancelled_waiter_does_not_cancel_shared_preparation(self):
        release = asyncio.Event()
        async def load():
            await release.wait()
            return 'ready'
        cache = PreparedInput(load)
        waiter = asyncio.create_task(cache.get('key'))
        await asyncio.sleep(0)
        waiter.cancel()
        with self.assertRaises(asyncio.CancelledError):
            await waiter
        release.set()
        self.assertEqual(await cache.get('key'), 'ready')

    async def test_expired_snapshot_is_not_served_when_refresh_fails(self):
        now = [0]
        loader = AsyncMock(side_effect=['old', RuntimeError('offline')])
        cache = PreparedInput(loader, ttl=300, clock=lambda: now[0])
        self.assertEqual(await cache.get('url'), 'old')
        now[0] = 300
        with self.assertRaisesRegex(RuntimeError, 'offline'):
            await cache.get('url')
