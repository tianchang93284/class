import sys
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QRadioButton, QButtonGroup, QLabel, QFileDialog
from PyQt5.QtGui import QPainter, QImage, QPixmap, QPen
from PyQt5.QtCore import Qt, QPoint
import os

class DrawingDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("选择和绘制")
        self.setFixedSize(470, 380)

        # 布局和选择项按钮
        layout = QVBoxLayout()
        button_layout = QHBoxLayout()

        # 创建三个选择按钮
        self.check_button = QRadioButton("对勾")
        self.half_check_button = QRadioButton("半对勾")
        self.cross_button = QRadioButton("叉")

        # 将按钮添加到按钮组，确保只有一个选项可以被选择
        self.button_group = QButtonGroup()
        self.button_group.addButton(self.check_button)
        self.button_group.addButton(self.half_check_button)
        self.button_group.addButton(self.cross_button)

        # 将按钮添加到布局
        button_layout.addWidget(self.check_button)
        button_layout.addWidget(self.half_check_button)
        button_layout.addWidget(self.cross_button)

        # 添加一个绘图区
        self.canvas = QLabel(self)
        self.canvas.setFixedSize(500, 420)
        self.canvas.setStyleSheet("background-color: transparent;")

        # 设置画布的图片，使用 ARGB32 格式，背景为透明
        self.image = QImage(self.canvas.size(), QImage.Format_ARGB32)
        self.image.fill(Qt.transparent)  # 背景设置为透明

        #刷新按钮
        refresh_button = QPushButton("刷新")
        refresh_button.clicked.connect(self.refesh_canvas)

        # 保存按钮
        save_button = QPushButton("保存")
        save_button.clicked.connect(self.save_image)

        # 添加布局
        layout.addLayout(button_layout)
        layout.addWidget(self.canvas)
        layout.addWidget(refresh_button)
        layout.addWidget(save_button)
        self.setLayout(layout)

        # 绘制状态
        self.drawing = False
        self.last_point = QPoint()

    def mousePressEvent(self, event):
        # 检查是否在画布区域内点击
        if event.button() == Qt.LeftButton and self.canvas.geometry().contains(event.pos()):
            self.drawing = True
            self.last_point = event.pos() - self.canvas.pos()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton and self.drawing:
            painter = QPainter(self.image)
            pen = QPen(Qt.red, 5, Qt.SolidLine)
            painter.setPen(pen)
            current_point = event.pos() - self.canvas.pos()
            painter.drawLine(self.last_point, current_point)
            self.last_point = current_point
            self.update_canvas()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drawing = False

    def update_canvas(self):
        # 更新画布上的图片显示
        self.canvas.setPixmap(QPixmap.fromImage(self.image))

    def refesh_canvas(self ):
        self.image.fill(Qt.transparent)
        self.canvas.clear()

    def save_image(self):
        # 根据check_button，half_check_button，或cross_button的状态，将图片保存在gou文件夹，命名分别为gou,semigou,x
        goupath = 'gou'
        num_str = 0
        prefix = None
        for root, dirs, files in os.walk(goupath):
            for base_name in files:
                file_path = os.path.join(root, base_name)
                if self.check_button.isChecked():
                    prefix = "gou"
                    if base_name[:3] != prefix:
                        continue
                    num_str = max(num_str,int(base_name[3:base_name.find(".png")]))
                elif self.half_check_button.isChecked():
                    prefix = "semigou"
                    if base_name[:7] != prefix:
                        continue
                    num_str = max(num_str,int(base_name[7:base_name.find(".png")]))
                elif self.cross_button.isChecked():
                    prefix = "x"
                    if base_name[:1] != prefix:
                        continue
                    num_str = max(num_str,int(base_name[1:base_name.find(".png")]))
        num_str += 1
        self.image.save("gou/" +prefix+ str(num_str) + ".png")
