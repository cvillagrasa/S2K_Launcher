from s2k_launcher import S2KLauncherHandler
from s2k_general import S2KGeneralHandler


handlers = [S2KLauncherHandler, S2KGeneralHandler]


class S2KModel(*handlers):
    def __init__(self, **kwargs):
        S2KGeneralHandler.__init__(self, **kwargs)
        S2KLauncherHandler.__init__(self, **kwargs)

        self.check_filename_extension()
        self.launch()
        self.setup_model()
        self.setup_project()

    def __call__(self):
        return self.s2k_model


if __name__ == '__main__':
    s2k_model = S2KModel()
    print('SAP 2000 up and running!')
