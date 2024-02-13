import os
import inspect

from PyQt6 import QtGui
from PyQt6.QtCore import QUrl, QSize, Qt
from PyQt6.QtWidgets import (QMainWindow,QWidget, QHBoxLayout, 
                             QVBoxLayout, QTabWidget, QLabel,
                             QPushButton)

from PyQt6.QtGui import QDesktopServices

from cubematchfinance.assets.config.config import cfg_item
import cubematchfinance.entities.tabs as tabs

class View(QMainWindow):
     
    def __init__(self):

        super().__init__()
        
        self.setWindowIcon(QtGui.QIcon(os.path.join(*cfg_item("app","icon_path"))))
        self.setWindowTitle(cfg_item("app","title"))
        self.setGeometry(*cfg_item("app", "geometry"))
        
        self.__central_widget = QWidget(self)
        self.setCentralWidget(self.__central_widget)

        self.__vlayout = QHBoxLayout(self.__central_widget)
        self.__central_widget.setLayout(self.__vlayout)
        
        self.__render()
        
    def __render(self):
         
        self.__create_and_add_tabs()
        self.__add_link_buttons()

    def __create_and_add_tabs(self):
        
        tabs_list = inspect.getmembers(tabs)
        self.tab_instances = {} 
        __tabs = QTabWidget(parent = self.__central_widget)
        self.__vlayout.addWidget(__tabs)

        for name_object, object in tabs_list:
            if inspect.isclass(object) and "TabApp" in name_object:
                tab_instance = object()
                self.tab_instances[name_object] = tab_instance  
                __tabs.addTab(tab_instance, tab_instance.get_name())
                
    def __create_link_buttons(self, button, description, url):
            
        button = QPushButton(parent = self.__central_widget)
        label = QLabel(description,parent = self.__central_widget)
        icon_size = QSize(*cfg_item("view", "icon_size"))
        button_size = QSize(*cfg_item("view", "button_size"))
        style = self.__css_style(cfg_item('view','button_style'))
        label.setStyleSheet(self.__css_style(cfg_item("view","label_style")))
        button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(url)))
        button.setIconSize(icon_size)
        button.setFixedSize(button_size)
        button.setStyleSheet(style)
        
        return button, label
    
    def __add_link_buttons(self):
            
            vbox = QVBoxLayout()
            
            button_names = cfg_item("view","push_buttons")
            self.__vlayout.addLayout(vbox)

            for name in button_names:
                icon_path = os.path.join(*cfg_item("view","push_buttons", name, "icon_path"))
                url = cfg_item("view","push_buttons", name, "url")
                description = cfg_item("view","push_buttons", name, "name")
                button, label = self.__create_link_buttons(icon_path, description, url)
                button.setIcon(QtGui.QIcon(icon_path))
                vbox.addWidget(button, alignment=Qt.AlignmentFlag.AlignCenter)
                vbox.addWidget(label, alignment=Qt.AlignmentFlag.AlignCenter)
    
    def __css_style(self, styles_data):
        css_style = ""
        for key, value in styles_data.items():
            css_style += f"{key}: {value}; "
        return css_style
        
    def get_pushbuttons(self):
        pushbuttons = {}
        for name, tab_instance in self.tab_instances.items():
            try:
                pushbuttons[name] = tab_instance.get_pushbuttons()
            except AttributeError:
                pass
        return pushbuttons

    def get_checkbuttons(self):
        checkbuttons = {}
        for name, tab_instance in self.tab_instances.items():
            try:
                tab_instance.get_checkbuttons()
                checkbuttons[name] = tab_instance.get_checkbuttons()
            except AttributeError:
                pass
        return checkbuttons
    
    def get_tables(self):
        tables = {}
        for name, tab_instance in self.tab_instances.items():
            try:
                tab_instance.get_tables()
                tables[name] = tab_instance.get_tables()
            except AttributeError:
                pass
        return tables