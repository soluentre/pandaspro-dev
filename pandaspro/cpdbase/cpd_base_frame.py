import pandas as pd
import pandaspro as cpd
from abc import ABC
import inspect
from pandaspro.cpdbase.design import cpdBaseFrameDesign
from pandaspro.cpdbase.files_version_parser import FilesVersionParser


def extract_params(func):
    signature = inspect.signature(func)
    pos_params = [
        name for name, param in signature.parameters.items()
        if param.default == inspect.Parameter.empty
    ]
    kw_params_with_defaults = {
        name: param.default for name, param in signature.parameters.items()
        if param.default != inspect.Parameter.empty
    }
    return pos_params, kw_params_with_defaults


def cpdBaseFrame(
        path: str = None,
        default_version: str = 'latest',
        dateid: str = '%Y-%m-%d',
        file_type: str = 'csv',
        fiscal_year_end: str = '06-30'
):
    def decorator(myclass):
        class CombinedClass(myclass, cpdBaseFrameDesign, ABC):
            @classmethod
            def get_file_versions_parser(cls):
                return FilesVersionParser(
                    path = cls.get_path(),
                    class_prefix = cls.__name__,
                    dateid_expression = dateid,
                    file_type = file_type,
                    fiscal_year_end = fiscal_year_end
                )

            @classmethod
            def get_path(cls):
                if path and (hasattr(myclass, 'get_path') or hasattr(myclass, 'path')):
                    raise AttributeError('Declaring @decorator path argument and in-class path/get_path is not allowed, only declare one')
                elif path and not (hasattr(myclass, 'get_path') or hasattr(myclass, 'path')):
                    return path
                else:
                    if hasattr(myclass, 'get_path') and hasattr(myclass, 'path'):
                        raise AttributeError('Declaring in-class path and get_path is not allowed, only declare one')
                    elif hasattr(myclass, 'path'):
                        return myclass.path
                    elif hasattr(myclass, 'get_path'):
                        return myclass.get_path()
                    else:
                        raise TypeError("Can't instantiate abstract class MyConcreteClass with abstract method get_path")

            @classmethod
            def get_filename(cls):
                if default_version == 'latest':
                    filename = cls.get_file_versions_parser().get_latest_file()
                elif 'latest' in default_version:
                    freq = default_version.split('_')[1]
                    filename = cls.get_file_versions_parser().get_latest_file(freq)
                else:
                    filename = cls.get_file_versions_parser().get_file(default_version)

                return filename

            @classmethod
            def load(cls, version='latest'):
                filename = cls.get_filename()
                if file_type == 'csv':
                    return cpd.pwread(cls.get_path() + f'/{filename}', low_memory=False)[0]
                elif file_type == 'xlsx':
                    return cpd.pwread(cls.get_path() + f'/{filename}')[0]
                else:
                    raise ValueError('Invalid file type, can only read .csv/.xlsx format.')

            def __init__(self, *args, **kwargs):
                cpd_kwargs = extract_params(CombinedClass.load)[1]
                cpd_kwargs = {key: kwargs.pop(key, value) for key, value in cpd_kwargs.items()}
                if args or kwargs:
                    super(CombinedClass, self).__init__(*args, **kwargs)
                else:
                    this_frame = CombinedClass.load(**cpd_kwargs)
                    super(CombinedClass, self).__init__(this_frame)  # Ensure DataFrame initialization

            @property
            def _constructor(self):
                return CombinedClass

            @property
            def df(self):
                return pd.DataFrame(self)

        CombinedClass.__name__ = myclass.__name__
        return CombinedClass

    return decorator


@cpdBaseFrame(default_version='latest_month', dateid='%Y%m%d')
class SOB(pd.DataFrame):
    path = r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\Staff on Board Database\csv'
    pass
    # @classmethod
    # def get_path(cls):
    #     return "345"

    # @classmethod
    # def load(cls, version='latest'):
    #     return cpd.pwread(cls.get_path() + f'/{version}.csv', low_memory=False)[0]




# @cpdBaseFrame("Greetings", 456, region=None)
# class MyDataFrame2(pd.DataFrame):
#     pass


# 测试
df1 = SOB()
print(df1.shape)

# df2 = MyDataFrame2(region="Asia")
# print(df2)
