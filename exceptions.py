class Error(Exception):
    pass

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
