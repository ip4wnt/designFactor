import zipfile


class ZipReader:

    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path, "r")
        self.names = set(self.z.namelist())

    def open(self, name):
        return self.z.open(name)

    def read(self, name):
        return self.z.read(name)

    def exists(self, name):
        return name in self.names

    def list(self, prefix):
        return [
            name
            for name in sorted(self.names)
            if name.startswith(prefix)
        ]

    def close(self):
        self.z.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
