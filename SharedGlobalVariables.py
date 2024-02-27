headers = ["Recorder","Audio Name",	"Statement","StatementID"]
class Singleton:
    _instance = None
    activeFilename=None
    activeWB = None
    activeSheet = None
    username = None

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls, *args, **kwargs)
            cls.activeFilename=''
            cls.activeWB = ''
            cls.activeSheet = []
            cls.username = '' 
        return cls._instance
