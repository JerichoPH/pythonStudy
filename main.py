class SingletonType(type):
    _instance = None

    def __call__(self, *args, **kwargs):
        if not self._instance:
            self._instance = self.__new__(self)

        return self._instance


class SingletonCls(object, metaclass=SingletonType):
    _a = 0

    def __init__(self):
        pass

    @property
    def get_val(self) -> int:
        return self._a

    def set_val(self, a: int) -> __init__:
        self._a = a
        return self


a = SingletonCls()
a.set_val(123)
print(a.get_val)

b = SingletonCls()
print(b.get_val)
