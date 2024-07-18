from abc import ABCMeta, abstractmethod, ABC


class cpdBaseFrameDesign(metaclass=ABCMeta):
    @classmethod
    def instruction(cls):
        print("If the folder is holding several versions of the csv/xlsx, name it as  ... ")
        print(">> [class name]_[date/time] + [any potential meta data necessary].xlsx/csv")
        print("Note: no need to add other info in the file name so it's easier to read")

    @classmethod
    @abstractmethod
    def get_path(cls):
        pass

    @classmethod
    @abstractmethod
    def load(cls):
        pass