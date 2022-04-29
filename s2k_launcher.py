import sys
import time


class S2KLauncherHandler:
    def __init__(self, **kwargs):
        self.t0 = time.time()
        self.attach = kwargs.get('attach', False)
        self.program_path = kwargs.get('program_path')
        self.path = kwargs.get('path')
        self.filename = kwargs.get('filename')
        self.remote_computer = kwargs.get('remote_computer')
        self.client = kwargs.get('client', 'COM')
        self.s2k_helper = None
        self.s2k_object = None
        self.s2k_model = None
        if self.client == 'NET':
            self.add_net_reference()

    def __call__(self):
        if self.client == 'COM':
            return self.s2k_model
        else:
            from SAP2000v1 import cSapModel
            return cSapModel(self.s2k_model)

    def add_net_reference(self):
        import clr
        clr.AddReference('System.Runtime.InteropServices')
        clr.AddReference(r'C:\Program Files\Computers and Structures\SAP2000 23\SAP2000v1.dll')

    def elapsed_time(self):
        return int(time.time() - self.t0)

    def launch(self):
        print(f'S2K LAUNCHER: Using API in {self.client} mode.')
        if self.path is None:
            print(f'S2K LAUNCHER: Warning, no destination path was passed. Do not attempt to save.')
        else:
            if not self.path.exists():
                print(f'S2K LAUNCHER: Creating directory {self.path}')
                self.path.mkdir(parents=True, exist_ok=True)

        if self.attach:
            self.s2k_object = self.attach_to_instance_com() if self.client == 'COM' else self.attach_to_instance_net()
            print(f'S2K LAUNCHER: Attached to instance')
        else:
            self.s2k_helper = self.get_helper_com() if self.client == 'COM' else self.get_helper_net()
            if self.program_path:
                self.s2k_object = \
                    self.launch_from_given_path_com() if self.client == 'COM' else self.launch_from_given_path_net()
            else:
                self.s2k_object = \
                    self.launch_from_system_path_com() if self.client == 'COM' else self.launch_from_system_path_net()
            print(f'S2K LAUNCHER: Application successfully initialized.')
            self.s2k_object.ApplicationStart()
            print(f'S2K LAUNCHER: Application successfully started in {self.elapsed_time()} sec.')

    def relaunch(self):
        self.s2k_model = None
        self.s2k_object = None
        print(f'S2K LAUNCHER: Using API in {self.client} mode.')
        self.launch()
        self.setup_model()
        self.setup_project()

    def get_helper_com(self):
        import comtypes.client
        return comtypes.client.CreateObject('SAP2000v1.Helper').QueryInterface(comtypes.gen.SAP2000v1.cHelper)

    def get_helper_net(self):
        from SAP2000v1 import cHelper, Helper
        return cHelper(Helper())

    def attach_to_instance_com(self):
        try:
            import comtypes.client
            return comtypes.client.GetActiveObject('CSI.SAP2000.API.SapObject')
        except:
            print(f'S2K LAUNCHER: WARNING, no running instance of the program found or failed to attach (COM).')
            sys.exit(-1)

    def attach_to_instance_net(self):
        try:
            from SAP2000v1 import cOAPI
            return cOAPI(self.s2k_helper.GetObject('CSI.SAP2000.API.SAPObject'))
        except:
            print(f'S2K LAUNCHER: WARNING, no running instance of the program found or failed to attach (.NET).')
            sys.exit(-1)

    def launch_from_given_path_com(self):
        try:
            return self.s2k_helper.CreateObject(self.program_path)
        except:
            print(f'S2K LAUNCHER: WARNING, cannot start a new instance of the program from {self.program_path}')
            sys.exit(-1)

    def launch_from_given_path_net(self):
        try:
            from SAP2000v1 import cOAPI
            return cOAPI(self.s2k_helper.CreateObject(self.program_path))
        except:
            print(f'S2K LAUNCHER: WARNING, cannot start a new instance of the program from {self.program_path}')
            sys.exit(-1)

    def launch_from_system_path_com(self):
        try:
            return self.s2k_helper.CreateObjectProgID('CSI.SAP2000.API.SapObject')
        except:
            print('S2K LAUNCHER: WARNING, cannot start a new instance of the program.')
            sys.exit(-1)

    def launch_from_system_path_net(self):
        if self.remote_computer is not None:
            return self.launch_from_system_path_net_remote()
        else:
            try:
                from SAP2000v1 import cOAPI
                return cOAPI(self.s2k_helper.CreateObjectProgID('CSI.SAP2000.API.SapObject'))
            except:
                print('S2K LAUNCHER: WARNING, cannot start a new instance of the program.')
                sys.exit(-1)

    def launch_from_system_path_net_remote(self):
        try:
            from SAP2000v1 import cOAPI
            return cOAPI(self.s2k_helper.CreateObjectProgIDHost(self.remote_computer, 'CSI.SAP2000.API.SAPObject'))
        except:
            print('S2K LAUNCHER: WARNING, cannot connect with remote computer.')
            sys.exit(-1)

    def setup_model(self):
        self.s2k_model = self.s2k_object.SapModel
        assert self().InitializeNewModel() == 0
        print(f'S2K LAUNCHER: SAP2000 OAPI model is up and running...')
        if self.client == 'COM':
            external_version, internal_version, ret = self().GetVersion(); assert ret == 0
        else:
            ret, external_version, internal_version = self().GetVersion('', 0); assert ret == 0
        print(f'S2K LAUNCHER: Running software version v{external_version}')

    def check_filename_extension(self):
        if self.filename is None:
            self.filename = 'default.sdb'
        else:
            if self.filename.split('.')[-1] != 'sdb':
                self.filename += '.sdb'

    def setup(self):
        self.check_filename_extension()
        self.launch()
        self.setup_model()


if __name__ == '__main__':
    s2k_model = S2KLauncherHandler()
    s2k_model.setup()
    print('S2K Model launched!')
