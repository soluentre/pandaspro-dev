import os
from datetime import datetime


class VersionParser:
    def __init__(self, path, class_prefix, id_expression='%Y-%m-%d', file_type='csv', fiscal_year_end='06-30'):
        if path.endswith(('/', r'\\')):
            raise ValueError(r'path cannot end with either / or \\')
        self.path = path
        self.class_prefix = class_prefix
        self.id_expression = id_expression
        self.file_type = file_type
        self.fiscal_year_end_month = int(fiscal_year_end.split('-')[0])
        self.fiscal_year_end_day = int(fiscal_year_end.split('-')[1])

    def check_single_file(self, version):
        filename_start = self.class_prefix + '_' + version
        self.filename_start = filename_start
        files = os.listdir(self.path)
        matching_files = [f for f in files if f.startswith(self.filename_start) and (f.endswith(f'.{self.file_type}'))]
        if len(matching_files) != 1:
            raise Exception(
                f"Expected exactly one file starting with '{self.filename_start}' but found {len(matching_files)}")
        else:
            self.suffix = matching_files[0][len(self.filename_start) + 1:]
            return matching_files[0], self.suffix

    def get_file(self, version):
        return self.check_single_file(version)[0]

    def get_suffix(self, version):
        return self.check_single_file(version)[1]

    def _can_parse_date(self, date_str):
        try:
            datetime.strptime(date_str, self.id_expression)
            return True
        except ValueError:
            return False

    def _filter_by_frequency(self, dates, freq):
        if freq == 'fiscal_year':
            return [(file, date) for file, date in dates if
                    date.month == self.fiscal_year_end_month and date.day == self.fiscal_year_end_day]
        elif freq == 'year':
            return [(file, date) for file, date in dates if date.month == 12 and date.day == 31]
        elif freq == 'quarter':
            return [(file, date) for file, date in dates if date.month in [3, 6, 9, 12] and date.day == 31]
        elif freq == 'month':
            return [(file, date) for file, date in dates if date.day in (30, 31)]
        elif freq == 'day':
            return [(file, date) for file, date in dates if date.hour == 23 and date.minute == 59 and date.second == 59]
        elif freq == 'hour':
            return [(file, date) for file, date in dates if date.minute == 59 and date.second == 59]
        elif freq == 'minute':
            return [(file, date) for file, date in dates if date.second == 59]
        else:
            raise ValueError(f"Unknown frequency: {freq}")

    def get_latest_file(self, freq='none'):
        # Configure matching files
        files = os.listdir(self.path)
        matching_files = [
            f for f in files if
            f.startswith(self.class_prefix + '_') and
            f.endswith(f'.{self.file_type}') and
            self._can_parse_date(f.split('_')[1])
        ]
        if not matching_files:
            raise ValueError('No matching files detected')

        # Configure dates list and apply use max method
        dates = [(file, datetime.strptime(file.split('_')[1], self.id_expression)) for file in matching_files]
        if freq == 'none':
            return max(dates, key=lambda x: x[1])[0] if dates else None
        freq_filtered_dates = self._filter_by_frequency(dates, freq)

        return max(freq_filtered_dates, key=lambda x: x[1])[0] if freq_filtered_dates else None


if __name__ == '__main__':
    vp = VersionParser(
        path = r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\Staff on Board Database\csv',
        class_prefix = 'SOB',
        id_expression = '%Y%m%d',
        fiscal_year_end = '05-31'
    )
    print(vp.path)
    print(vp.check_single_file('20240715'))
    print(vp.get_latest_file('fiscal_year'))
    print(vp._can_parse_date('no date'))
