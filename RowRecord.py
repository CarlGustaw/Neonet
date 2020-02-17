class RowRecord:
    id_dkf = ""

    def __init__(self, id_dkf):
        self.id_dkf = id_dkf

    def show_row_object(self):
        print("ID-DKF: ", self.id_dkf)

    def get_dkf(self):
        return self.id_dkf
