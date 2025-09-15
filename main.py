import warnings
import os.path
import time
from create_epub_from_xlsx import translated_xlsx_to_epub, clear_files, epub_to_xlsx
import sys
import traceback
from PyQt6 import QtWidgets, QtCore
from PyQt6.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QTextEdit,
    QMessageBox, QProgressDialog
)

warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)


class EpubToXlsxWorker(QtCore.QThread):
    progress = QtCore.pyqtSignal(int, int, str)   # (current, total, message)
    finished = QtCore.pyqtSignal(bool, str)       # (ok, message)
    exception = QtCore.pyqtSignal(str)            # traceback text

    def __init__(self, directory: str = "./to_translate", parent=None):
        super().__init__(parent)
        self.directory = directory

    def run(self):
        try:
            files = []
            if os.path.isdir(self.directory):
                files = sorted(os.listdir(self.directory))
            total = len(files)
            index = 0

            for i, file in enumerate(files, start=1):
                path = os.path.join(self.directory, file)
                if not os.path.isfile(path):
                    continue
                self.progress.emit(i, total, f"[{i}/{total}] EPUB → XLSX: {file}")
                try:
                    epub_to_xlsx(path, index)
                except Exception as e:
                    tb = "".join(traceback.format_exception(type(e), e, e.__traceback__))
                    self.exception.emit(tb)
                index += 1

            self.finished.emit(True, "EPUB → XLSX: Готово")
        except Exception as e:
            tb = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.exception.emit(tb)
            self.finished.emit(False, f"Ошибка: {e}")

class XlsxToEpubWorker(QtCore.QThread):
    progress = QtCore.pyqtSignal(int, int, str)   # (current, total, message)
    finished = QtCore.pyqtSignal(bool, str)       # (ok, message)
    exception = QtCore.pyqtSignal(str)            # traceback text

    """
    Поведение полностью повторяет ваш show_dialog2:
      - Перебираем ./to_epub
      - Уникализируем по префиксу до первого '-': ids_list
      - Увеличиваем index при каждой новой id
      - Обрабатываем только файлы, где в первой части после '_' значение вида '*-0' (т.е. split('_')[0].split('-')[1] == '0')
      - translated_xlsx_to_epub(xlsx_path, book_name, index)
      - time.sleep(0.1)
    Прогресс считает только те файлы, которые реально будут обработаны (соответствуют условию и уникальной id).
    """

    def __init__(self, directory: str = "./to_epub", parent=None):
        super().__init__(parent)
        self.directory = directory

    def _enumerate_targets(self):
        if not os.path.isdir(self.directory):
            return []
        files = sorted(os.listdir(self.directory))
        ids_list = []
        targets = []
        for file in files:
            path = os.path.join(self.directory, file)
            if not os.path.isfile(path):
                continue
            # уникализация по первой части до '-'
            id_prefix = file.split('-')[0]
            if id_prefix not in ids_list:
                ids_list.append(id_prefix)
                # условие на суффикс "-0" в первом блоке до '_'
                first_block = file.split('_')[0]
                parts = first_block.split('-')
                if len(parts) > 1 and parts[1] == '0':
                    targets.append(file)
        return targets

    def run(self):
        try:
            targets = self._enumerate_targets()
            total = len(targets)

            # index по вашей логике: начинается с 1, увеличивается на каждой новой id
            # Так как мы уже уникализировали, index = порядковый номер цели
            for i, file in enumerate(targets, start=1):
                path = os.path.join(self.directory, file)
                book_name = ".".join(file.split(".")[:-1])  # как у вас
                self.progress.emit(i, total, f"[{i}/{total}] XLSX → EPUB: {file}")

                try:
                    translated_xlsx_to_epub(path, book_name, i)  # i == index
                except Exception as e:
                    tb = "".join(traceback.format_exception(type(e), e, e.__traceback__))
                    # ровно как у вас: печатаем, но продолжаем
                    print(f"[ERROR] translated_xlsx_to_epub failed: {e}")
                    self.exception.emit(tb)
                time.sleep(0.1)

            self.finished.emit(True, "XLSX → EPUB: Готово")
        except Exception as e:
            tb = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.exception.emit(tb)
            self.finished.emit(False, f"Ошибка: {e}")

class EpubTranslatorForm(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Перевод EPUB ↔ XLSX (с прогрессом)")
        self.resize(700, 360)

        # кнопки
        self.btn_epub_to_xlsx = QPushButton("EPUB → Temp XLSX (обработать ./to_translate)")
        self.btn_xlsx_to_epub = QPushButton("RUS XLSX → RUS EPUB (обработать ./to_epub)")
        self.btn_clear = QPushButton("Очистить папки (result, temp, to_epub, to_translate)")

        # логи
        self.log = QTextEdit()
        self.log.setReadOnly(True)

        # раскладка
        row = QHBoxLayout()
        row.addWidget(self.btn_epub_to_xlsx)
        row.addWidget(self.btn_xlsx_to_epub)

        root = QVBoxLayout(self)
        root.addLayout(row)
        root.addWidget(self.btn_clear)
        root.addWidget(self.log, stretch=1)

        # сигналы
        self.btn_epub_to_xlsx.clicked.connect(self.run_epub_to_xlsx)
        self.btn_xlsx_to_epub.clicked.connect(self.run_xlsx_to_epub)
        self.btn_clear.clicked.connect(self.run_clear)

        # прогресс
        self.progress: QProgressDialog | None = None
        self.worker: QtCore.QThread | None = None

    # ---------- прогресс ----------
    def _start_progress(self, title: str, text: str):
        self.progress = QProgressDialog(text, None, 0, 100, self)
        self.progress.setWindowTitle(title)
        self.progress.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.progress.setAutoClose(False)
        self.progress.setAutoReset(False)
        self.progress.setMinimumDuration(0)
        self.progress.setValue(0)
        self.progress.show()

    def _stop_progress(self):
        if self.progress:
            self.progress.hide()
            self.progress.deleteLater()
            self.progress = None

    def _bind_worker(self, worker: QtCore.QThread, title: str):
        self.worker = worker

        def on_progress(current: int, total: int, message: str):
            self.append_log(message)
            if self.progress:
                if total <= 0:
                    self.progress.setRange(0, 0)  # «бесконечный»
                else:
                    self.progress.setRange(0, total)
                    self.progress.setValue(current)

                self.progress.setLabelText(f"{title}\n{message}")

        def on_exception(tb_text: str):
            print("[ERROR]\n" + tb_text)
            self.append_log(tb_text)

        def on_finished(ok: bool, msg: str):
            self._stop_progress()
            QMessageBox.information(self, "Результат", "Успех" if ok else f"Ошибка: {msg}")
            self.worker = None

        worker.progress.connect(on_progress)
        worker.exception.connect(on_exception)
        worker.finished.connect(on_finished)

    def append_log(self, text: str):
        self.log.append(text)

    # ---------- действия ----------
    def run_epub_to_xlsx(self):
        directory = "./to_translate"
        if not os.path.isdir(directory):
            QMessageBox.warning(self, "Внимание", f"Папка не найдена: {directory}")
            return
        worker = EpubToXlsxWorker(directory)
        self._start_progress("EPUB → XLSX", f"Обработка папки: {directory}")
        self._bind_worker(worker, "EPUB → XLSX")
        worker.start()

    def run_xlsx_to_epub(self):
        directory = "./to_epub"
        if not os.path.isdir(directory):
            QMessageBox.warning(self, "Внимание", f"Папка не найдена: {directory}")
            return
        worker = XlsxToEpubWorker(directory)
        self._start_progress("XLSX → EPUB", f"Обработка папки: {directory}")
        self._bind_worker(worker, "XLSX → EPUB")
        worker.start()

    def run_clear(self):
        reply = QMessageBox.question(
            self, "Подтверждение",
            "Очистить папки result, temp, to_epub, to_translate?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # чистим синхронно, это быстро; если хотите – перевести в поток по аналогии
        try:
            self._start_progress("Очистка", "Удаление файлов...")
            QtWidgets.QApplication.processEvents()
            clear_files()
            self._stop_progress()
            QMessageBox.information(self, "Готово", "Папки очищены")
            self.append_log("[clear] Папки очищены")
        except Exception as e:
            self._stop_progress()
            tb = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            print("[ERROR]\n" + tb)
            self.append_log(tb)
            QMessageBox.critical(self, "Ошибка", f"Не удалось очистить папки:\n{e}")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = QtWidgets.QWidget()
    layout = QVBoxLayout(w)
    ui = EpubTranslatorForm()
    layout.addWidget(ui)
    w.setWindowTitle("Перевод EPUB ↔ XLSX (с прогрессом)")
    w.resize(760, 420)
    w.show()
    try:
        sys.exit(app.exec())
    except Exception as e:
        print("[FATAL]", e)
        traceback.print_exc()
        raise