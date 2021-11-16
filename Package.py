import sys
import os
from PyQt5 import QtWidgets # For GUI Window
import GUI_Package #GUI file convertation
from docxtpl import DocxTemplate
import openpyxl
from datetime import datetime # To know when Register file was modified last time


# Window with result
class DialogApp(QtWidgets.QDialog, GUI_Package.Ui_Dialog):
    def __init__(self, result_message, res_fname, er_ocur):
        super().__init__()
        self.setupUi(self)
        self.lblResultMessage.setText(result_message)
        self.res_fname = res_fname
        # If there was an error, button will have other purpose
        if er_ocur:
            self.btnOk.setText("OK")
            self.btnOk.clicked.connect(self.hide)
        else:
            self.btnOk.clicked.connect(self.open_folder)

    # open folder with result if there weren't any errors
    def open_folder(self):
        os.startfile(self.res_fname)
        self.hide()


# Start Window
class MainApp(QtWidgets.QMainWindow, GUI_Package.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  
        # Checking if there Register file with default name and file location
        if os.path.isfile(os.path.dirname(os.getcwd()) + "\Реестр.xlsx"):
            self.reg_name = os.path.dirname(os.getcwd()) + "\Реестр.xlsx"
        # Using back up Register file
        else:
            self.reg_name = "Реестр_default.xlsx"
        # Initial setting
        self.res_changed = False
        self.lineEditRegister.setText(os.path.abspath(self.reg_name))  # Initial setting
        self.btnRegister.clicked.connect(self.browse_Register)  # Do browse_Register if button is clicked
        # Checking if there template folder with default name and location
        if os.path.isfile(os.path.dirname(os.getcwd()) + "\Package_template.docx"):
            self.tmpl_name = os.path.dirname(os.getcwd()) + "\Package_template.docx"
        # Using back up Register file
        else:
            self.tmpl_name = "Package_template_default.docx"
            # Initial setting
        self.lineEditTemplate.setText(os.path.abspath(self.tmpl_name))  # Initial setting
        self.btnTemplate.clicked.connect(self.browse_Template)  # Do browse_Template if button is clicked
        # Finding out when selected Register file was modified last time
        reg_mtime = datetime.fromtimestamp(os.stat(self.reg_name).st_mtime).strftime("%d.%m.%Y_%H.%M")
        # Initial setting
        self.lineEditResult.setText(os.path.abspath(os.path.dirname(os.getcwd()) + "\\Упаковочные_листы_" + reg_mtime))
        self.btnResult.clicked.connect(self.browse_Result)  # Do browse_Result if button is clicked
        self.btnStart.clicked.connect(self.process)  # Do main body of utility

    # selecting new Register file with usage of dialog window
    def browse_Register(self):
        ch_file = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите файл Реестр", self.reg_name, " *.xls *.xlsx")
        if ch_file[0]:  # checking that User selected something
            self.lineEditRegister.setText(ch_file[0])
            # If User hasn't selected specific folder for result saving, modify future folder name
            if not self.res_changed:
                reg_mtime = datetime.fromtimestamp(os.stat(ch_file[0]).st_mtime).strftime("%d.%m.%Y_%H.%M")
                self.lineEditResult.setText(
                    os.path.abspath(
                        os.path.dirname(
                            os.getcwd()
                        ) +
                        "\\Упаковочные_листы_" +
                        reg_mtime
                    )
                )

    # selecting new package template with usage of dialog window
    def browse_Template(self):
        ch_file = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Выберите шаблон упаковочного листа",
            os.path.dirname(self.lineEditTemplate.text()),
            " *.doc *.docx"
        )
        if ch_file[0]:  # checking that User selected something
            self.lineEditTemplate.setText(ch_file[0])

    # selecting new folder for results with usage of dialog window
    def browse_Result(self):
        ch_fold = QtWidgets.QFileDialog.getExistingDirectory(
            self,
            "Выберите папку, куда сохранить результаты работы",
            os.path.dirname(self.lineEditResult.text())
        )
        if ch_fold:  # checking that User selected something
            self.lineEditResult.setText(ch_fold)
            self.res_changed = True  # Stop modifying name of folder with result when new Register file is selected

    # main body of utility
    def process(self):
        # Reset
        self.res_changed = False
        # If selected Register file doesn't exist prepare message for User
        if os.path.isfile(self.lineEditRegister.text()) == False:
            result_message = "Работа утилиты прервана, т.к. указанного файл-Реестра не существует"
            res_fname = ""
            er_ocur = True
        # If selected Template doesn't exist prepare message for User
        elif os.path.exists(self.lineEditTemplate.text()) == False:
            result_message = "Работа утилиты прервана, т.к. указанного шаблона упаковочного листа не существует"
            er_ocur = True
            res_fname = ""
        # If selected Register file and package template do exist
        else:
            # Reset
            er_ocur = False
            unpacked_eq = False
            # open Register file
            reg_name = self.lineEditRegister.text()
            wb = openpyxl.load_workbook(reg_name, data_only=True)
            # reading time of last modification of register file
            # reading raw data in a row way from register. values_only = 1 is for reading data, not formules.
            # We need List to be able change values. (it is tuple by default)
            raw_data = list(map(list, wb.active.iter_rows(values_only=True)))
            # To avoid writing "None" in package documents there is a substitution for spaces and None
            for i in range(len(raw_data)):
                for j in range(len(raw_data[i])):
                    if raw_data[i][j] is None or (set(str(raw_data[i][j]))) == {' '}:
                        raw_data[i][j] = ''
            # first row with meanings in register file will be used for tags in word document
            label_tuple = tuple(map(str, raw_data[0]))
            # creating set for package's id
            pack_nums = set()
            for i in range(1, len(raw_data)):
                # filling a set of package's id
                pack_nums.add(raw_data[i][0])
                # modifying raw data for further convinient usage
                raw_data[i] = dict(zip(label_tuple, raw_data[i]))
                # Three columns (B : D) is used for table filling in word template. They need to be put in dictionary
                raw_data[i]["pack_cont"] = dict(map(lambda x: (x, raw_data[i][x]), label_tuple[1:4]))
            # creating dictionary for tags replacement for each package
            cont_dict = {i: None for i in pack_nums}
            for i in range(1, len(raw_data)):
                # when package's id appears for the first time
                if cont_dict[raw_data[i][label_tuple[0]]] is None:
                    # filling tags information not including table tags
                    cont_dict[raw_data[i][label_tuple[0]]] = raw_data[i]
                    # creating list for tags information consisted from table tags
                    cont_dict[raw_data[i][label_tuple[0]]]["pack_cont"] = [
                        cont_dict[raw_data[i][label_tuple[0]]]["pack_cont"]]
                # when package id appears NOT for the first time
                else:
                    # adding information in  table tags list
                    cont_dict[raw_data[i][label_tuple[0]]]["pack_cont"].append(raw_data[i]["pack_cont"])
            # creating folder with result
            res_fname = self.lineEditResult.text()
            os.makedirs(res_fname, exist_ok=True)
            for i in cont_dict:
                #   choosing template with tags.
                doc = DocxTemplate(self.lineEditTemplate.text())
                # creating dictionary for tags replacement for current package
                # Useless. Just for understanding
                context = cont_dict[i]
                #   tags replacement
                doc.render(context, autoescape=True)
                # if package's id is empty
                if i == "":
                    doc.save(res_fname + "\\" + 'Нераспределенное оборудование.docx')
                    unpacked_eq = True
                # if package's id isn't empty
                else:
                    doc.save(res_fname + "\\" + 'Упаковочный лист ' + str(i) + '.docx')
            # Message for User generating
            # Show the name of folder where results were saved
            result_message = "Упаковочные листы заполнены и сохранены в папку " + res_fname
            # When some equipment have  package's id
            if unpacked_eq:
                result_message += "\n\n\nЕсть оборудование без номера коробки." \
                                  " Подробная информация приведена в файле: Нераспределенное оборудование"
        # call for Window with result
        window = DialogApp(result_message, res_fname, er_ocur)
        window.show()
        window.exec_()

def main():
    # call for start Window
    app = QtWidgets.QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
