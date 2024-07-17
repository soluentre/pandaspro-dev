from abc import ABCMeta, abstractmethod
from typing import Iterable, Any


class cpdBase(metaclass=ABCMeta):
    @classmethod
    @abstractmethod
    def path(cls):
        pass

    @classmethod
    @abstractmethod
    def cpdbase_engine(cls):
        pass


if __name__ == '__main__':
    from pandaspro import FramePro

    class mydata(FramePro, cpdBase):
        @classmethod
        def path(cls):
            return 'abc'

        @classmethod
        def cpdbase_engine(cls):
            return 'abc'

        def __init__(self):
            super().__init__({'a': [1, 2, 3], 'b': [2, 3, 4]})

    d = mydata()
    m = mydata.path
    print(mydata.path())  # 使用类来访问类方法
    print(mydata.cpdbase_engine())  # 使用类来访问类方法