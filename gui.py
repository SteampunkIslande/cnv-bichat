import json
import os
import sys
from pathlib import Path

import PySide6.QtCore as qc
import PySide6.QtGui as qg
import PySide6.QtWidgets as qw


def get_user_prefs_file():
    return (
        Path(
            qc.QStandardPaths().writableLocation(
                qc.QStandardPaths.StandardLocation.AppDataLocation
            )
        )
        / "config.json"
    ).resolve()


def save_user_prefs(prefs: dict):

    user_prefs = get_user_prefs_file()
    print("Writing user prefs to", user_prefs)
    if not user_prefs.parent.exists():
        user_prefs.parent.mkdir(parents=True, exist_ok=True)
    old_prefs = {}
    if user_prefs.exists():
        with open(user_prefs, "r", encoding="utf-8") as f:
            old_prefs = json.load(f)

    old_prefs.update(prefs)

    with open(user_prefs, "w", encoding="utf-8") as f:
        json.dump(old_prefs, f)


def load_user_prefs():
    user_prefs = get_user_prefs_file()
    print("Loading user prefs from", user_prefs)
    prefs = {}
    if user_prefs.exists():
        with open(user_prefs, "r", encoding="utf-8") as f:
            prefs = json.load(f)
    return prefs


class MainWindow(qw.QMainWindow):

    def __init__(self, parent: qw.QWidget = None):

        # Simple GUI with three buttons:
        # 1) Select reference directory (saved in user config)
        # 2) Select working directory (saved in user config)
        # 3) Run the script with the selected directories (should prompt for ZIP file with openFileDialog and ask for a directory to save the output)
        super().__init__(parent)
        self.setWindowTitle("CNV Script")

        self.setFixedSize(300, 150)

        self._central_widget = qw.QWidget()

        self._main_layout = qw.QVBoxLayout(self._central_widget)

        self._run_button = qw.QPushButton("LANCER LE SCRIPT")
        self._run_button.setSizePolicy(
            qw.QSizePolicy.Policy.Expanding, qw.QSizePolicy.Policy.Expanding
        )
        self._run_button.setFont(qg.QFont("Arial", 20))
        self._run_button.clicked.connect(self.run_cnvcall_script)

        self._progressbar = qw.QProgressBar(self)
        self._progressbar.setRange(0, 0)

        # Add menu bar entry to open a Preferences dialog
        self._menu_bar = self.menuBar()
        self._file_menu = self._menu_bar.addMenu("&Fichier")

        self._select_refdir_action = qg.QAction("BED de référence...", self)
        self._select_refdir_action.triggered.connect(self.select_refbed)

        self._select_design_bed_action = qg.QAction("BED de design...", self)
        self._select_design_bed_action.triggered.connect(self.select_design_bed)

        self._select_workdir_action = qg.QAction("Répertoire de travail...", self)
        self._select_workdir_action.triggered.connect(self.select_workdir)

        # Simply opens the json file in the default text editor
        self.advanced_user_prefs = qg.QAction("Préférences avancées...", self)
        self.advanced_user_prefs.triggered.connect(
            lambda: qg.QDesktopServices.openUrl(
                qc.QUrl.fromLocalFile(get_user_prefs_file())
            )
        )

        self._file_menu.addAction(self._select_refdir_action)
        self._file_menu.addAction(self._select_workdir_action)
        self._file_menu.addAction(self.advanced_user_prefs)

        self._main_layout.addWidget(self._run_button)
        self._main_layout.addWidget(self._progressbar)

        self._design_bed = None
        self._ref_bed = None
        self._workdir = None
        self._last_file = None

        prefs = load_user_prefs()
        if "designbed" in prefs:
            self._design_bed = prefs["designbed"]
        if "refbed" in prefs:
            self._ref_bed = prefs["refbed"]
        if "workdir" in prefs:
            self._workdir = prefs["workdir"]
        if "last_file" in prefs:
            self._last_file = prefs["last_file"]

        self.setup_normal_mode()

        self.setCentralWidget(self._central_widget)

        self._process = qc.QProcess(self)

    def select_refbed(self):
        refbed, _ = qw.QFileDialog.getOpenFileName(
            self,
            "Choisissez le BED de référence (témoin CNV)",
            filter="Fichiers BED (*.bed)",
        )
        if refbed and os.path.isfile(refbed):
            self._ref_bed = refbed
            save_user_prefs({"refbed": refbed})
        else:
            qw.QMessageBox.information(
                self,
                "Erreur",
                f"Répertoire de référence invalide. L'emplacement du dossier de référence n'a pas été modifié. ({self._ref_bed})",
            )

    def select_design_bed(self):
        designbed, _ = qw.QFileDialog.getOpenFileName(
            self,
            "Choisissez le BED de design",
            filter="Fichiers BED (*.bed)",
        )
        if designbed and os.path.isfile(designbed):
            self._design_bed = designbed
            save_user_prefs({"designbed": designbed})
        else:
            qw.QMessageBox.information(
                self,
                "Erreur",
                f"Répertoire de design invalide. L'emplacement du dossier de design n'a pas été modifié. ({self._design_bed})",
            )

    def select_workdir(self):
        workdir = qw.QFileDialog.getExistingDirectory(
            self, "Choisissez le répertoire de travail"
        )
        if workdir and os.path.isdir(workdir):
            self._workdir = workdir
            save_user_prefs({"workdir": workdir})
        else:
            qw.QMessageBox.information(
                self,
                "Erreur",
                f"Répertoire de travail invalide. L'emplacement du dossier de travail n'a pas été modifié. ({self._workdir})",
            )

    def run_cnvcall_script(self):
        if self._process.state() == qc.QProcess.ProcessState.Running:
            qw.QMessageBox.information(
                self,
                "What?",
                "How did you manage to run this method while the process is running? The button is supposed to be hidden!!!",
            )
            print("Congrats ;)")
            return
        if self._ref_bed is None:
            self.select_refbed()
        if self._workdir is None:
            self.select_workdir()
        if self._design_bed is None:
            self.select_design_bed()
        if self._ref_bed is None or self._workdir is None or self._design_bed is None:
            return
        # Prompt for an existing ZIP file
        input_files, _ = qw.QFileDialog.getOpenFileNames(
            self,
            caption="Choisissez le ou les fichiers d'entrée",
            dir=self._last_file or qc.QDir().homePath(),
            filter="Fichiers ZIP (*.zip);;Fichiers BED (*.bed)",
        )
        if not input_files:
            return
        self._last_file = input_files[0]
        save_user_prefs({"last_file": self._last_file})

        self.run_name, _ = qw.QInputDialog.getText(
            self, "Nom du run", "Veuillez entrer le nom du run"
        )
        if not self.run_name:
            return

        arguments = [
            "--reference-coverage-bed",
            self._ref_bed,
            "--design-bed",
            self._design_bed,
            "--workdir",
            os.path.join(self._workdir, self.run_name),
            *input_files,
        ]

        if "__compiled__" in globals():
            self._process.setProgram(sys.argv[0])
        else:
            self._process.setProgram(sys.executable)
            arguments.insert(0, sys.argv[0])
        self._process.setArguments(arguments)
        self._progressbar.setRange(0, 0)

        qw.QMessageBox.information(
            self,
            "Running script",
            "Starting process {0} with arguments {1}".format(
                self._process.program(), " ".join(self._process.arguments())
            ),
        )

        self._process.start()

        if self._process.waitForStarted():
            self.setup_wait_mode()
            self._process.finished.connect(self.on_worker_finished)
        else:
            qw.QMessageBox.critical(self, "Erreur", "Impossible de lancer le script")

    def closeEvent(self, event: qg.QCloseEvent):
        if self._process.state() == qc.QProcess.ProcessState.Running:
            if (
                qw.QMessageBox.question(
                    self,
                    "Arrêter le script ?",
                    "Le script est en cours d'exécution. Voulez-vous vraiment l'arrêter ?",
                    qw.QMessageBox.StandardButton.Yes
                    | qw.QMessageBox.StandardButton.No,
                )
                == qw.QMessageBox.StandardButton.Yes
            ):
                self._process.kill()
                self._process.waitForFinished()
                event.accept()
            else:
                event.ignore()

    def on_worker_finished(self, returncode: int):
        if returncode != 0:
            qw.QMessageBox.critical(self, "Erreur", "Le script a échoué")
        else:
            qw.QMessageBox.information(
                self, "Succès !", "Le script s'est terminé avec succès !"
            )
            qg.QDesktopServices.openUrl(
                qc.QUrl.fromLocalFile(os.path.join(self._workdir, self.run_name))
            )
        self.setup_normal_mode()

    def setup_wait_mode(self):
        self._run_button.hide()
        self._progressbar.show()

    def setup_normal_mode(self):
        self._run_button.show()
        self._progressbar.hide()


def main_gui():
    app = qw.QApplication(sys.argv)
    app.setApplicationName("CNVScript")
    window = MainWindow()
    window.show()
    return app.exec()
