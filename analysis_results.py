"""분석 결과의 서버 세션 저장과 이전 저장 형식 복원."""
import pandas as pd
from nicegui import json


def encode_result(df_errors, df_first, df_summary, dl_name, summary, df_view):
    return json.dumps({
        'df_errors': df_errors.to_dict('records'),
        'df_first': df_first.to_dict('records'),
        'df_summary': df_summary.to_dict('records'),
        'df_view': df_view.to_dict('records') if df_view is not None else [],
        'dl_name': dl_name,
        'summary': summary,
    })


def decode_result(stored):
    result = json.loads(stored) if isinstance(stored, str) else stored
    return (
        pd.DataFrame(result['df_errors']),
        pd.DataFrame(result['df_first']),
        pd.DataFrame(result['df_summary']),
        result['dl_name'], result['summary'],
        pd.DataFrame(result.get('df_view', [])),
    )
