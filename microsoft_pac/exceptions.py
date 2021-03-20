class UnknownValue(Exception):
    """ Raised when incorrect value supplied """
    def __init__(self, value, valid_values, message="Incorrect value passed"):
        self.value = value
        self.message = f"{message} {self.value}: Valid Options = {valid_values}"
        super().__init__(self.message)