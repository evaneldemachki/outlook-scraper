class Error(Exception):
    pass

class UiRuntimeError(Error):
    def __str__(self):
        msg = self.args[0]
        return msg

class NoConfigFile(Error):
    def __str__(self):
        msg = "configuration file has not yet been generated"
        return msg

class ExistingConfig(Error):
    def __str__(self):
        msg = "configuration file already generated"
        return msg

class InvalidLimit(Error):
    def __str__(self):
        msg = "failed to parse config limit '{0}'".format(self.args[0])
        return msg

class InvalidEntryID(Error):
    def __str__(self):
        msg = "row with specified entry ID does not exist"
        return msg

class DuplicateEntryID(Error):
    def __str__(self):
        msg = "specified entry ID already exists in table"
        return msg

class InvalidMirrorTable(Error):
    def __str__(self):
        msg = "specified mirror table does not exist"
