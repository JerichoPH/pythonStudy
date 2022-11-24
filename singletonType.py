class SingletonType(type):
    _instance = None

    def __call__(self, *args, **kwargs):
        if not self._instance:
            self._instance = self.__new__(self)

        self.__init__(self._instance, *args, **kwargs)

        return self._instance
