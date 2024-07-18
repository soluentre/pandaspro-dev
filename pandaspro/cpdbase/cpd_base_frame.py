import pandas as pd
import pandaspro as cpd
from abc import ABC
import inspect
from pandaspro.cpdbase.design import cpdBaseFrameDesign
from pandaspro.cpdbase.files_version_parser import FilesVersionParser
import textwrap

from pandaspro.utils.cpd_logger import cpdLogger


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
        load: str = None,
        default_version: str = 'latest',
        dateid: str = '%Y%m%d',
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
                    raise AttributeError('Declaring both path argument (1) in @decorator and (2) in-class path/get_path at the same time is not allowed, please only declare one.')
                elif path and not (hasattr(myclass, 'get_path') or hasattr(myclass, 'path')):
                    return path
                else:
                    if hasattr(myclass, 'get_path') and hasattr(myclass, 'path'):
                        raise AttributeError('Declaring in-class path and get_path is not allowed, only declare one.')
                    elif hasattr(myclass, 'path'):
                        return myclass.path
                    elif hasattr(myclass, 'get_path'):
                        return myclass.get_path()
                    else:
                        raise TypeError("Can't instantiate abstract class MyConcreteClass with abstract method get_path.")

            @classmethod
            def get_filename(cls, version):
                if version == 'latest':
                    filename = cls.get_file_versions_parser().get_latest_file()
                elif 'latest' in version:
                    freq = version.split('_')[1]
                    filename = cls.get_file_versions_parser().get_latest_file(freq)
                else:
                    filename = cls.get_file_versions_parser().get_file(version)

                return filename

            @classmethod
            def read_table(cls, version):
                filename = cls.get_filename(version)
                if file_type == 'csv':
                    return cpd.pwread(cls.get_path() + f'/{filename}', low_memory=False)
                elif file_type == 'xlsx':
                    return cpd.pwread(cls.get_path() + f'/{filename}')
                else:
                    raise ValueError('Invalid file type, can only read .csv/.xlsx format.')

            @staticmethod
            def load(data, **kwargs):
                return data.head(5)

            @classmethod
            def get_process_method(cls):
                if load and hasattr(myclass, 'load'):
                    raise AttributeError('Declaring both load argument (1) in @decorator and (2) in-class load attribute/method at the same time is not allowed, please only declare one.')
                elif hasattr(myclass, 'load'):
                    return myclass.load
                elif load:
                    return load
                else:
                    return CombinedClass.load

            def __init__(self, *args, **kwargs):
                cpd_kwargs = extract_params(CombinedClass.get_process_method())[1]
                version_kwarg = {'version': kwargs.pop('version', default_version)}
                other_kwargs = {key: kwargs.pop(key, value) for key, value in cpd_kwargs.items()}
                # self.debug.info(f'[cpd_kwargs]: {cpd_kwargs}')
                # self.debug.info(f'[version_kwarg]: {version_kwarg}')
                # self.debug.info(f'[other_kwargs]: {other_kwargs}')
                # self.debug.info(f'[kwargs]: {kwargs}')
                # self.debug.info(f'[args]: {args}')
                if args or kwargs:
                    # self.debug_info_lv1('Inside __init__')
                    # self.logger.info(f'Entered Above Part of init: args: **{type(args)}**, kwargs: **{type(kwargs)}**')
                    # self.logger.debug(f'Seeing values -> args: **{args}**, kwargs: **{kwargs}**')
                    try:
                        super(CombinedClass, self).__init__(*args, **kwargs)
                    except ValueError as e:
                        raise ValueError(textwrap.dedent(f'''
                            --------------------------------------
                            {e}
                            Please only pass key-word arguments when you want to create new instances of the <{myclass.__name__}>. 
                            You are passing {args} as positional arguments, which can not be parsed by the pandas DataFrame constructor
                        '''))
                    except TypeError as e:
                        raise TypeError(textwrap.dedent(f'''
                            --------------------------------------
                            {e}
                            Please check that you don't pass any extra key-word arguments besides those you declared in load defined in class <{myclass.__name__}>.
                            For the load method defined, the class constructor can only take the following kwargs: {list(other_kwargs.keys())}  
                        '''))
                else:
                    # self.logger.info('Entered Below Part of init: no args or kwargs detected')
                    raw_frame, name_map = CombinedClass.read_table(**version_kwarg)
                    processed_frame = CombinedClass.get_process_method()(raw_frame, **other_kwargs)
                    super(CombinedClass, self).__init__(processed_frame)  # Ensure DataFrame initialization

                    self.filename = CombinedClass.get_filename(version_kwarg['version'])
                    self.version =  CombinedClass.get_file_versions_parser().get

            @property
            def _constructor(self):
                return CombinedClass

            @property
            def df(self):
                return pd.DataFrame(self)

        CombinedClass.__name__ = myclass.__name__
        return CombinedClass

    return decorator


@cpdBaseFrame(default_version='latest_month')
class SOB(pd.DataFrame):
    path = r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\Staff on Board Database\csv'

    @staticmethod
    def load(data, region=None):
        print(region)
        return data.head(30)

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
df1 = SOB(region='balabala')
print(df1.shape)

# df2 = MyDataFrame2(region="Asia")
# print(df2)