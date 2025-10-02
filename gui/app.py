from PyQt5 import QtWidgets
import sqlite3
import subprocess
import sys

DB_PATH = "dados.db"


class MainWindow(QtWidgets.QWidget):
    """Janela principal da aplicação."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ferramenta de Integração")
        self._build_ui()
        self.load_data()

    def _build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        form_layout = QtWidgets.QFormLayout()
        self.user_edit = QtWidgets.QLineEdit()
        self.pass_edit = QtWidgets.QLineEdit()
        self.pass_edit.setEchoMode(QtWidgets.QLineEdit.Password)
        form_layout.addRow("Usuário:", self.user_edit)
        form_layout.addRow("Senha:", self.pass_edit)
        layout.addLayout(form_layout)

        button_layout = QtWidgets.QHBoxLayout()
        self.deps_btn = QtWidgets.QPushButton("Verificar dependências")
        self.tests_btn = QtWidgets.QPushButton("Rodar testes")
        self.gmail_btn = QtWidgets.QPushButton("Processar Gmail")
        button_layout.addWidget(self.deps_btn)
        button_layout.addWidget(self.tests_btn)
        button_layout.addWidget(self.gmail_btn)
        layout.addLayout(button_layout)

        self.tabs = QtWidgets.QTabWidget()
        self.geral_table = QtWidgets.QTableWidget()
        self.cadastros_table = QtWidgets.QTableWidget()
        self.tabs.addTab(self.geral_table, "Geral")
        self.tabs.addTab(self.cadastros_table, "Cadastros")
        layout.addWidget(self.tabs)

        self.deps_btn.clicked.connect(self.verify_dependencies)
        self.tests_btn.clicked.connect(self.run_tests)
        self.gmail_btn.clicked.connect(self.process_gmail)

    def verify_dependencies(self):
        missing = []
        try:
            import PyQt5  # noqa: F401
        except ImportError:
            missing.append("PyQt5")
        if missing:
            QtWidgets.QMessageBox.warning(
                self, "Dependências", f"Dependências faltando: {', '.join(missing)}"
            )
        else:
            QtWidgets.QMessageBox.information(
                self, "Dependências", "Todas as dependências estão instaladas."
            )

    def run_tests(self):
        proc = subprocess.run(["pytest"], capture_output=True, text=True)
        QtWidgets.QMessageBox.information(
            self, "Testes", proc.stdout or proc.stderr
        )

    def process_gmail(self):
        QtWidgets.QMessageBox.information(
            self, "Gmail", "Processamento do Gmail não implementado."
        )

    def load_data(self):
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            for table_name, widget in [
                ("Geral", self.geral_table),
                ("Cadastros", self.cadastros_table),
            ]:
                cur.execute(f"SELECT * FROM {table_name}")
                rows = cur.fetchall()
                headers = [desc[0] for desc in cur.description]
                widget.setRowCount(len(rows))
                widget.setColumnCount(len(headers))
                widget.setHorizontalHeaderLabels(headers)
                for row_idx, row in enumerate(rows):
                    for col_idx, value in enumerate(row):
                        item = QtWidgets.QTableWidgetItem(str(value))
                        widget.setItem(row_idx, col_idx, item)
            conn.close()
        except sqlite3.Error as exc:
            QtWidgets.QMessageBox.warning(self, "Banco de dados", str(exc))


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
