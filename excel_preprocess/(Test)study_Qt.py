from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QLabel

class MyWindows(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('GEpmTool')
        self.resize(800,600)

        '''btn definit'''
        btn_generate = QPushButton('Generate',self)
        btn_generate.setGeometry(330, 500, 100, 32)

        btn_outfielder = QPushButton('Output Folder',self)
        btn_outfielder.setGeometry(450, 500, 100, 32)

        btn_quit = QPushButton('Quit',self)
        btn_quit.setGeometry(570, 500, 100, 32)

        '''edit label definit'''
        label_name = QLabel('PM engineer',self)
        label_name.setGeometry(330,30,81,16)


        '''edit buff definit'''
        edit_name = QLineEdit(self)
        edit_name.setGeometry(330,50,113,21)
        edit_name.setPlaceholderText('your name')

        edit_tel = QLineEdit(self)
        edit_tel.setGeometry(330,100,113,21)
        edit_tel.setPlaceholderText('phone number')

        edit_doc_path = QLineEdit(self)
        edit_doc_path.setGeometry(330,170,432,21)
        edit_doc_path.setPlaceholderText('/Volumes/SSD 1TB/GEhealthcare/Doc/report_demo.xlsx')

        edit_tel = QLineEdit(self)
        edit_tel.setGeometry(330,220,432,21)
        edit_tel.setPlaceholderText('/Volumes/SSD 1TB/GEhealthcare/202508/Output')

if __name__ == '__main__':
    app = QApplication([])
    window = MyWindows()
    window.show()
    app.exec()


