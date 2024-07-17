from abc import ABCMeta, abstractmethod


class cpdBaseFrameDesign(metaclass=ABCMeta):
    @classmethod
    def instruction(cls):
        print("If the folder is holding several versions of the csv/xlsx, name it as  ... ")
        print(">> [class name]_[date/time] + [any potential meta data necessary].xlsx/csv")
        print("Note: no need to add other info in the file name so it's easier to read")

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

    class mydata(FramePro, cpdBaseFrameDesign):
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
    print(mydata.path())
    print(mydata.cpdbase_engine())
