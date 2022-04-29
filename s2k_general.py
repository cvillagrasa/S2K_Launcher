class S2KGeneralHandler:
    def __init__(self, **kwargs):
        self.project_info = kwargs.get('project_info', {})
        self.s2k_model = None
        self.s2k_object = None
        self.path = None
        self.filename = None
        self.node = kwargs.get('node')

    def __call__(self):
        return self.s2k_model

    def setup_project(self):
        assert self().File.NewBlank() == 0
        kN_m_C = 6
        assert self().SetPresentUnits(kN_m_C) == 0

        for element, value in self.project_info.items():
            if element is not None:
                assert self().SetProjectInfo(element, value) == 0

        self.delete_all_materials()

    def refresh_view(self):
        assert self().View.RefreshView(0, False) == 0

    def save(self):
        assert self().File.Save(self.path / self.filename) == 0

    def close(self, save=False):
        assert self.s2k_object.ApplicationExit(save) == 0
        self.s2k_model = None
        self.s2k_object = None

    def is_closed(self):
        closed = False
        try:
            program, version, license, ret = self().GetProgramInfo()
        except:
            closed = True
        return closed

    def make_groups(self, groups):
        if not isinstance(groups, list):
            groups = [groups]
        for group in groups:
            assert self().GroupDef.SetGroup(group) == 0
