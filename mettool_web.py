import logging
import sys
import os
import io  # Use io for in-memory files
import json
from pandas.io import json as pd_json # For NaN-safe JSON
import pandas as pd
import numpy as np
from scipy import stats
from copy import deepcopy

# --- Setup Logging ---
logging.basicConfig(level=logging.DEBUG,
                    format='[%(asctime)s] [%(levelname)s] – %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    handlers=[logging.StreamHandler(sys.stderr)])

class Mettool:
    def __init__(self, input_path: str | io.BytesIO | None = None, sheet_name: str = 'Input'):
        self.input_data_path = input_path
        self.input_sheet_name = sheet_name
        self.date_column = 'Date'
        self.data = pd.DataFrame()
        self.original_data = pd.DataFrame()
        self.column_list = []

    def read_data(self) -> bool:
        if not self.input_data_path:
            logging.error('read_data: invalid path')
            return False

        engine = None
        if isinstance(self.input_data_path, str):
            ext = os.path.splitext(self.input_data_path)[1].lower()
            if ext in {'.xlsm', '.xlsx'}:
                engine = 'openpyxl'
            elif ext == '.xls':
                engine = 'xlrd'
            else:
                raise ValueError(f'Unsupported extension {ext}')
        else:
            engine = 'openpyxl'

        try:
            self.data = pd.read_excel(self.input_data_path, sheet_name=self.input_sheet_name, engine=engine)
            self.original_data = deepcopy(self.data)
            self._remove_junk_columns(0.8)
            logging.info('Loaded %s rows × %s cols', *self.data.shape)
            return True
        except Exception:
            logging.error('read_data failed', exc_info=True)
            return False

    def _remove_junk_columns(self, threshold: float):
        nan_frac = self.data.isna().mean()
        numeric_cols = self.data.select_dtypes(include=np.number).columns
        zero_frac = pd.Series(0, index=self.data.columns)
        if not numeric_cols.empty:
            zero_frac[numeric_cols] = (self.data[numeric_cols] == 0).mean()

        keep_mask = (nan_frac + zero_frac) < threshold
        if 'Date' in self.data.columns:
            keep_mask['Date'] = True

        self.data = self.data.loc[:, keep_mask]
        logging.debug('Cols kept: %s', list(self.data.columns))

    def select_data(self, start: str, end: str, cols: list[str]):
        dayfirst = '/' in start
        start_dt = pd.to_datetime(start, dayfirst=dayfirst)
        end_dt = pd.to_datetime(end, dayfirst=dayfirst)

        if start_dt > end_dt:
            raise ValueError('Start > End')

        df = self.original_data.copy()
        if self.date_column not in df.columns:
            raise KeyError(f'"{self.date_column}" not found')

        df[self.date_column] = pd.to_datetime(df[self.date_column], errors='coerce')
        mask = (df[self.date_column] >= start_dt) & (df[self.date_column] <= end_dt)

        if self.date_column not in cols:
            cols = cols + [self.date_column]

        self.data = df.loc[mask, cols].copy()
        self.column_list = cols
        logging.info('select_data → %s rows', len(self.data))

    def filter_data(self, replace_zero=True, remove_outliers=False, outlier_method='zscore', outlier_threshold=3.0):
        num = [c for c in self.column_list if c != self.date_column]
        for c in num:
            self.data[c] = pd.to_numeric(self.data[c], errors='coerce')
            if replace_zero:
                self.data.loc[self.data[c] == 0, c] = np.nan

            if remove_outliers and not self.data[c].isna().all():
                if outlier_method.lower() == 'zscore':
                    col_data = self.data[c].dropna()
                    if not col_data.empty:
                        z = np.abs(stats.zscore(col_data))
                        z_series = pd.Series(z, index=col_data.index)
                        self.data.loc[z_series > outlier_threshold, c] = np.nan

                elif outlier_method.lower() == 'iqr':
                    q1 = self.data[c].quantile(0.25)
                    q3 = self.data[c].quantile(0.75)
                    iqr = q3 - q1
                    lo = q1 - (outlier_threshold * iqr)
                    hi = q3 + (outlier_threshold * iqr)
                    self.data.loc[(self.data[c] < lo) | (self.data[c] > hi), c] = np.nan

    def export_filtered_data(self, sheet: str) -> bytes:
        try:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as wt:
                self.data.to_excel(wt, sheet_name=sheet, index=False)
            logging.debug('Exported %s to in-memory buffer', sheet)
            return buf.getvalue()
        except Exception:
            logging.error('export_filtered_data failed', exc_info=True)
            return b""

    def correlation(self):
        numeric_cols = self.data.select_dtypes(include=np.number).columns
        return self.data[numeric_cols].corr()

    def cusum(self, col: str):
        return (self.data[col] - self.data[col].mean()).cumsum()

    def control_limits(self, col: str, conf: float):
        s = self.data[col].dropna()
        if len(s) < 2:
            return np.nan, np.nan

        m = s.mean()
        t = stats.t.ppf(1 - (1 - (conf / 100.0)) / 2, df=len(s) - 1)
        se = s.std(ddof=1) / np.sqrt(len(s))

        upper = m + t * se
        lower = m - t * se
        return upper, lower

    def get_data(self):
        return self.data.copy()

# --- Global Mettool Instance ---
mettool_instance = Mettool()


# --- Web-Facing Functions (Called from JavaScript) ---

def get_sheet_names_from_buffer(buffer) -> list:
    try:
        py_buffer = io.BytesIO(buffer.to_py())
        xls = pd.ExcelFile(py_buffer)
        return xls.sheet_names
    except Exception as e:
        logging.error(f"Failed to get sheet names: {e}")
        return []

def load_data_from_buffer(buffer, sheet_name: str) -> None:
    global mettool_instance
    py_buffer = io.BytesIO(buffer.to_py())
    mettool_instance = Mettool(py_buffer, sheet_name)
    if not mettool_instance.read_data():
        raise RuntimeError("Failed to read Excel data")

def get_columns() -> list:
    global mettool_instance
    return list(mettool_instance.get_data().columns)

def clear_data() -> None:
    global mettool_instance
    mettool_instance = Mettool()

def run_correlation(params_proxy: dict) -> dict:
    params = params_proxy.to_py()
    global mettool_instance

    m = mettool_instance
    m.date_column = params['dateCol']
    m.select_data(params['start'], params['end'], params['cols'])
    m.filter_data()

    corr_matrix = m.correlation()

    plot_data = {
        'z': corr_matrix.values.tolist(),
        'x': corr_matrix.columns.tolist(),
        'y': corr_matrix.columns.tolist()
    }

    file_buffer = m.export_filtered_data('Correlation')

    return {'plot_data_json': pd_json.dumps(plot_data), 'file_buffer': file_buffer}


def run_cusum(params_proxy: dict) -> dict:
    params = params_proxy.to_py()
    global mettool_instance

    m = mettool_instance
    m.date_column = params['dateCol']
    m.select_data(params['start'], params['end'], params['cols'])
    m.filter_data()

    df = m.get_data().sort_values(by=m.date_column)

    traces = []
    for col in params['cols']:
        traces.append({
            'name': f'CUSUM {col}',
            'data': m.cusum(col).tolist()
        })

    plot_data = {
        'dateCol': df[m.date_column].dt.strftime('%Y-%m-%d').tolist(),
        'dateColName': m.date_column,
        'traces': traces
    }

    file_buffer = m.export_filtered_data('CUSUM')

    return {'plot_data_json': pd_json.dumps(plot_data), 'file_buffer': file_buffer}


def run_control_graph(params_proxy: dict) -> dict:
    params = params_proxy.to_py()
    global mettool_instance

    col = params['col']
    conf = params['conf']
    periods = params['periods']

    plot_traces = []

    palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
               '#8c564b', '#e377c2', '#7f7f0f', '#bcbd22', '#17becf']

    full_period_data = []

    for i, period in enumerate(periods):
        s, e = period['start'], period['end']
        color = palette[i % len(palette)]

        tmp = Mettool(mettool_instance.input_data_path, mettool_instance.input_sheet_name)
        tmp.date_column = params['dateCol']
        tmp.original_data = deepcopy(mettool_instance.original_data)

        tmp.select_data(s, e, [col, tmp.date_column])
        tmp.filter_data()

        d = tmp.get_data().sort_values(by=tmp.date_column)
        full_period_data.append(d)

        # Data points (in color)
        plot_traces.append({
            'x': d[tmp.date_column].dt.strftime('%Y-%m-%d').tolist(),
            'y': d[col].tolist(),
            'name': f'{s} – {e}',
            'mode': 'markers',
            'color': color, # Use the period color
            'dateColName': tmp.date_column
        })

        period_start_dt = pd.to_datetime(s, dayfirst=True).strftime('%Y-%m-%d')
        period_end_dt = pd.to_datetime(e, dayfirst=True).strftime('%Y-%m-%d')

        if params['showLim']:
            up, lo = tmp.control_limits(col, conf)
            if not np.isnan(up):
                # UCL Line (neutral gray)
                plot_traces.append({
                    'x': [period_start_dt, period_end_dt],
                    'y': [up, up],
                    'name': f'UCL ({s}-{e})',
                    'mode': 'lines',
                    'line': {'color': '#888888', 'dash': 'dash', 'width': 2}, # <-- THE FIX: Neutral color
                    'showlegend': False
                })
                # LCL Line (neutral gray)
                plot_traces.append({
                    'x': [period_start_dt, period_end_dt],
                    'y': [lo, lo],
                    'name': f'LCL ({s}-{e})',
                    'mode': 'lines',
                    'line': {'color': '#888888', 'dash': 'dash', 'width': 2}, # <-- THE FIX: Neutral color
                    'showlegend': False
                })
            else:
                plot_traces.append({'x': [], 'y': [], 'mode': 'lines'})
                plot_traces.append({'x': [], 'y': [], 'mode': 'lines'})


        if params['showAvg']:
            avg = d[col].mean()
            if not np.isnan(avg):
                # Average Line (neutral gray)
                plot_traces.append({
                    'x': [period_start_dt, period_end_dt],
                    'y': [avg, avg],
                    'name': f'Avg ({s}-{e})',
                    'mode': 'lines',
                    'line': {'color': '#888888', 'dash': 'dot', 'width': 3}, # <-- THE FIX: Neutral color
                    'showlegend': False
                })
            else:
                plot_traces.append({'x': [], 'y': [], 'mode': 'lines'})

    if full_period_data:
        mettool_instance.data = pd.concat(full_period_data)
    else:
        mettool_instance.data = pd.DataFrame()

    file_buffer = mettool_instance.export_filtered_data('Control')

    return {'plot_data_json': pd_json.dumps(plot_traces), 'file_buffer': file_buffer}