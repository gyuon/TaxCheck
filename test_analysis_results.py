import unittest
from nicegui import json
from nicegui.observables import ObservableDict
import pandas as pd

from analysis_results import encode_result, decode_result


class TestAnalysisResults(unittest.TestCase):
    def test_json_and_legacy_snapshots_restore_identical_data(self):
        frames = [pd.DataFrame({'이름': ['가나다', 'A'], '금액': [1000, None], '코드': ['01', '21']}) for _ in range(4)]
        summary = {'counts': {'미납': 1}, 'duration': 1.2, 'source_filename': 'input_26.11.4.xlsx'}
        encoded = encode_result(*frames[:3], 'download.xlsx', summary, frames[3])
        self.assertIsInstance(encoded, str)
        legacy = json.loads(encoded)
        for stored in [encoded, legacy, ObservableDict(legacy)]:
            restored = decode_result(stored)
            for original, actual in zip(frames, [*restored[:3], restored[5]]):
                pd.testing.assert_frame_equal(original, actual)
            self.assertEqual(restored[3:5], ('download.xlsx', summary))

    def test_session_does_not_wrap_each_record_and_supports_missing_legacy_raw(self):
        frame = pd.DataFrame({'x': [1, 2]})
        encoded = encode_result(frame, frame, frame, 'a.xlsx', {}, frame)
        session = ObservableDict({'result': encoded})
        self.assertIsInstance(session['result'], str)
        legacy = json.loads(encoded)
        del legacy['df_view']
        self.assertTrue(decode_result(legacy)[5].empty)

    def test_empty_results_can_be_restored(self):
        empty = pd.DataFrame()
        encoded = encode_result(empty, empty, empty, 'a.xlsx', {}, empty)
        result = decode_result(encoded)
        self.assertTrue(all(result[i].empty for i in [0, 1, 2, 5]))
