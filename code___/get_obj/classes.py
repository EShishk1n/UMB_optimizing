
class ObjectFromRaports:

    def __init__(self, dzo: str, fieldname: str, padname: str, wellname: str):
        self.dzo = dzo
        self.fieldname = fieldname
        self.padname = padname
        self.wellname = wellname

    def __repr__(self):
        return (f"ДО: {self.dzo}; {self.fieldname} "
                f"к.{self.padname}, {self.wellname}")
