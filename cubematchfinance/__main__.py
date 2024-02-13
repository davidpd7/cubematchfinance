import sys

from PyQt6.QtWidgets import QApplication

from cubematchfinance.view import View
from cubematchfinance.controller import Controller
from cubematchfinance.models import Model

def main(args = None):

    app = QApplication(sys.argv)
    view = View()
    view.show()
    model = Model()
    controller = Controller(view, model)

    sys.exit(app.exec())

if __name__ == '__main__':
    main()
    
    