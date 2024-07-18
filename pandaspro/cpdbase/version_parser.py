import os
from datetime import datetime


class VersionParser:
    def __init__(self, path, class_prefix, id_expression='%Y-%m-%d', file_type='csv'):
        if path.endswith(('/', r'\\')):
            raise ValueError(r'path cannot end with either / or \\')
        self.path = path
        self.class_prefix = class_prefix
        self.id_expression = id_expression
        self.file_type = file_type

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

    def get_file(self, class_prefix, version):
        return self.check_single_file(version)[0]

    def get_suffix(self, class_prefix, version):
        return self.check_single_file(version)[1]

    def _can_parse_date(self, date_str):
        try:
            datetime.strptime(date_str, self.id_expression)
            return True
        except ValueError:
            return False

    def get_latest_file(self, freq='none'):
        # Configure matching files
        files = os.listdir(self.path)
        matching_files = [
            f for f in files if
            f.startswith(self.class_prefix + '_') and
            f.endswith(f'.{self.file_type}') and
            self._can_parse_date(f.split('_')[1])
        ]
        if len(matching_files) == 0:
            raise ValueError('Database files should match format: class_vid[_suffix].csv/.xlsx, no such file detected')
        print(matching_files)

        # Configure dates for comparison
        dates = [(file, datetime.strptime(file.split('_')[1], self.id_expression)) for file in matching_files]

        if freq == 'none':
            return max(dates, key=lambda x: x[1])[0] if dates else None

        now = datetime.now()
        if freq == 'fiscal_year':
            fiscal_dates = [(file, date) for file, date in dates if date.month == 6 and date.day == 30]
            return max(fiscal_dates, key=lambda x: x[1])[0] if fiscal_dates else None

        elif freq == 'year':
            year_dates = [(file, date) for file, date in dates if date.month == 12 and date.day == 31]
            return max(year_dates, key=lambda x: x[1])[0] if year_dates else None

        elif freq == 'month':
            month_dates = [(file, date) for file, date in dates if date.day == 30 or date.day == 31]
            return max(month_dates, key=lambda x: x[1])[0] if month_dates else None

        elif freq == 'day':
            day_dates = [(file, date) for file, date in dates if
                         date.hour == 23 and date.minute == 59 and date.second == 59]
            return max(day_dates, key=lambda x: x[1])[0] if day_dates else None

        elif freq == 'hour':
            hour_dates = [(file, date) for file, date in dates if date.minute == 59 and date.second == 59]
            return max(hour_dates, key=lambda x: x[1])[0] if hour_dates else None

        elif freq == 'minute':
            minute_dates = [(file, date) for file, date in dates if date.second == 59]
            return max(minute_dates, key=lambda x: x[1])[0] if minute_dates else None

        else:
            raise ValueError(f"Unknown frequency: {freq}")


if __name__ == '__main__':
    vp = VersionParser(
        path = r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\Staff on Board Database\csv',
        class_prefix = 'SOB',
        id_expression = '%Y%m%d'
    )
    print(vp.path)
    print(vp.check_single_file('20240715'))
    print(vp.get_latest_file('year'))
    print(vp._can_parse_date('no date'))
