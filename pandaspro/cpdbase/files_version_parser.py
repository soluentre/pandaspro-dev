import os
from datetime import datetime
from pandaspro.utils.cpd_logger import cpdLogger


@cpdLogger
class FilesVersionParser:
    def __init__(self, path, class_prefix, dateid_expression='%Y-%m-%d', file_type='csv', fiscal_year_end='06-30'):
        if path.endswith(('/', r'\\')):
            raise ValueError(r'path cannot end with either / or \\.')
        self.path = path
        self.class_prefix = class_prefix
        self.dateid_expression = dateid_expression
        self.file_type = file_type
        self.fiscal_year_end_month = int(fiscal_year_end.split('-')[0])
        self.fiscal_year_end_day = int(fiscal_year_end.split('-')[1])

        # Get the granularity of the id_expression
        self.granularity = self.get_granularity()

        # Perform duplicate check on initialization
        self.check_for_duplicates()

    def get_granularity(self):
        if '%S' in self.dateid_expression:
            return 'second'
        elif '%M' in self.dateid_expression:
            return 'minute'
        elif '%H' in self.dateid_expression:
            return 'hour'
        elif '%d' in self.dateid_expression:
            return 'day'
        elif '%m' in self.dateid_expression or '%B' in self.dateid_expression or '%b' in self.dateid_expression:
            return 'month'
        elif '%Y' in self.dateid_expression:
            return 'year'
        else:
            raise ValueError('Invalid dateid_expression provided.')

    def _can_parse_date(self, date_str):
        try:
            datetime.strptime(date_str, self.dateid_expression)
            return True
        except ValueError:
            return False

    def _filter_by_frequency(self, dates, freq):
        if freq == 'fiscal_year':
            return [(file, date) for file, date in dates if date.month == self.fiscal_year_end_month and date.day == self.fiscal_year_end_day]
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
            raise ValueError(f"Unknown frequency: <<{freq}>>.")

    def list_all_files(self):
        try:
            files = os.listdir(self.path)
            matching_files = [
                f for f in files if
                f.startswith(self.class_prefix + '_') and
                f.endswith(f'.{self.file_type}') and
                self._can_parse_date(f.split('_')[1])
            ]
            return matching_files
        except Exception as e:
            print(f"Error: {e}")
            return []

    def get_latest_file(self, freq='none'):
        # Define granularity levels
        granularity_levels = ['second', 'minute', 'hour', 'day', 'month', 'year']

        # Check if the provided frequency is valid
        if freq != 'none' and freq not in granularity_levels:
            raise ValueError(f"Unknown frequency: {freq}")

        # Check if the provided frequency is higher or same level as the current granularity
        if freq != 'none' and granularity_levels.index(freq) <= granularity_levels.index(self.granularity):
            raise ValueError(
                f"Invalid frequency: {freq}. The frequency must be higher than the current granularity: {self.granularity}")

        # Configure matching files
        matching_files = self.list_all_files()
        if len(matching_files) == 0:
            raise ValueError('No matching files detected.')

        # Configure dates list and apply use max method
        dates = [(file, datetime.strptime(file.split('_')[1], self.dateid_expression)) for file in matching_files]
        if freq == 'none':
            return max(dates, key=lambda x: x[1])[0] if dates else None
        freq_filtered_dates = self._filter_by_frequency(dates, freq)

        return max(freq_filtered_dates, key=lambda x: x[1])[0] if freq_filtered_dates else None

    @staticmethod
    def _find_duplicates(items):
        seen = {}
        duplicates = []
        for i, item in enumerate(items):
            if item in seen:
                duplicates.append(i)
                duplicates.append(seen[item])
            else:
                seen[item] = i
        return duplicates

    def check_for_duplicates(self):
        files = self.list_all_files()
        dates = [f.split('_')[1] for f in files]
        parsed_dates = [datetime.strptime(date, self.dateid_expression) for date in dates]
        duplicates = FilesVersionParser._find_duplicates(parsed_dates)

        if duplicates:
            duplicate_files = [files[i] for i in duplicates]
            print(f'Note your data tables should be unique at the <<{self.granularity}>> level.')
            print(f'Duplicate dates detected in the database folder for files: {duplicate_files}.')
            raise ValueError('See info printed above: go and fix the duplicates')

    def check_single_file(self, version):
        filename_start = self.class_prefix + '_' + version
        self.filename_start = filename_start
        files = os.listdir(self.path)
        matching_files = [f for f in files if f.startswith(self.filename_start) and (f.endswith(f'.{self.file_type}'))]
        if len(matching_files) != 1:
            raise Exception(
                f"Expected exactly one file starting with '{self.filename_start}' but found {len(matching_files)}.")
        else:
            self.suffix = matching_files[0][len(self.filename_start) + 1:]
            return matching_files[0], self.suffix

    def get_file(self, version):
        return self.check_single_file(version)[0]

    def get_suffix(self, version):
        return self.check_single_file(version)[1]

    def get_version_str(self, version):
        return self.get_file(version).split('_')[1]

    def get_version_dt(self, version):
        return datetime.strptime(self.get_version_str(version), self.dateid_expression)

## create a class: cpdDateStr - dt.year, dt.B, dt.b, dt.showtypes, dt.type1 + maya.parse
## A new package: datepro

if __name__ == '__main__':
    vp = FilesVersionParser(
        path = r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\Staff on Board Database\csv',
        class_prefix = 'SOB',
        dateid_expression = '%Y-%m-%d',
        fiscal_year_end = '05-31'
    )
    print(vp.path)
    print(vp.check_single_file('20240715'))
    print(vp.get_latest_file())
    # print(vp._can_parse_date('no date'))
    # print(vp.list_all_files())

