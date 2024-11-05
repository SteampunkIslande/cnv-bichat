#!/usr/bin/env python3
"""
Created on Fri Sept 27 2024
Adaptation à Linux et passage de python2 à python3 du script de Karim Diallo

Ajout d'une interface en ligne de commande.
Suppression du code neutralisé par commentaires.
Légère refactorisation du code.

"""

import argparse
import os
import sys

reference_files = [
    "listeCorrespondancePositionAmpliconGene.xlsx",
    "GENEXUS_fichierOrdonneRegionStartGene_PanelAPHP.xlsx",
    "fichierOrdonneRegionStartGene_PanelAPHP.xlsx",
    "Moyenne_NormalizedRead_count_TemoinsPorphyriesGENEXUS.xlsx",
]

import json
import os
from pathlib import Path

from gui import main_gui


def main():

    import argparse
    import sys

    if len(sys.argv) == 1:
        return main_gui()

    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument(
        "input_files",
        help="Either a zip file as outputted by the Genexus coverage plugin, or a list of amplicon coverage files",
        nargs="+",
        type=Path,
    )
    parser.add_argument(
        "--duplication-threshold",
        help="Threshold above which a copy number is considered at least a duplication",
        default=1.76,
        type=float,
    )
    parser.add_argument(
        "--deletion-threshold",
        help="Threshold below which a copy number is considered at least a deletion",
        default=0.5,
        type=float,
    )
    parser.add_argument(
        "--workdir",
        help="Working directory. Defaults to current path",
        default=os.getcwd(),
        type=Path,
    )
    parser.add_argument("--design-bed", help="Design bed file", type=Path)
    parser.add_argument(
        "--reference-coverage-bed", help="Reference coverage bed file", type=Path
    )
    parser.add_argument(
        "--keep-normalized-depth",
        help="Keep normalized depth",
        action="store_true",
    )
    args = parser.parse_args()

    entry_point(
        args.input_files,
        args.workdir,
        args.design_bed,
        args.duplication_threshold,
        args.deletion_threshold,
    )

    return 0


def tsv_to_json(
    input_filename: str | os.PathLike,
    output_filename: str | os.PathLike,
    key_column_name: str,
    delimiter: str = "\t",
):
    """Transforms a tsv file into a json file, where the key is the value of the key_column_name column.
    input_filename should be a tsv file with a header. Delimiter can be changed.
    Output keys are the values of the key_column_name column, and the values are dictionaries with the column names as keys.

    Args:
        input_filename (str | os.PathLike): Path to input tsv file (with header)
        output_filename (str | os.PathLike): Path to output json file
        key_column_name (str): Name of the column to use as key
        delimiter (str, optional): Input file delimiter. Defaults to "\t".
    """
    import csv
    import json

    with open(input_filename) as input_file:
        reader = csv.DictReader(input_file, delimiter=delimiter)
        d = {
            field[key_column_name]: {
                k: v for k, v in field.items() if k != key_column_name
            }
            for field in reader
        }

    with open(output_filename, "w") as output_file:
        json.dump(d, output_file)


def extract_amplicon_files_from_zip(
    input_zip: str | os.PathLike, workdir: str | os.PathLike
):
    """Writes all amplicon coverage files from the input zip to the workdir.
    Returns a list of written files paths (relative to workdir).

    Args:
        input_zip (str | os.PathLike): Zip file to extract amplicon coverage files from
        workdir (str | os.PathLike): Directory to write the amplicon coverage files to (will try to create if it doesn't exist)
    """
    if not os.path.isdir(workdir):
        os.makedirs(workdir, exist_ok=True)
    written_files = []
    import zipfile

    with zipfile.ZipFile(input_zip, "r") as zip_ref:
        for dirname in zipfile.Path(zip_ref).iterdir():
            if dirname.is_dir():
                samplename = dirname.name
                for filename in dirname.iterdir():
                    if filename.name.endswith(".amplicon.cov.xls"):
                        with open(
                            os.path.join(workdir, samplename + ".bed"),
                            "wb",
                        ) as output_file:
                            with filename.open("rb") as input_file:
                                for line in input_file:
                                    output_file.write(line)
                                written_files.append(samplename + ".bed")
    return written_files


def entry_point(
    input_files, workdir, design_bed, duplication_threshold, deletion_threshold
):
    os.makedirs(workdir, exist_ok=True)
    for input_file in input_files:
        input_file = Path(input_file)
        if input_file.suffix == ".zip":
            input_bed_filenames = [
                workdir / "extracted" / p
                for p in extract_amplicon_files_from_zip(
                    input_file, os.path.join(workdir, "extracted")
                )
            ]
            aggregate_samples_coverage(
                input_bed_filenames, design_bed, workdir / "all_samples_coverage.tsv"
            )


def cnv_call(input_bed_filenames, design_bed_filename, reference_coverage_bed_filename):
    pass


def aggregate_samples_coverage(
    input_bed_filenames: list[Path], design_bed_filename: Path, output_filename: Path
):
    input_bed_filenames = [Path(p) for p in input_bed_filenames]
    for input_bed_filename in input_bed_filenames:
        tsv_to_json(
            input_bed_filename, input_bed_filename.with_suffix(".json"), "region_id"
        )
    if not os.path.exists(design_bed_filename.with_suffix(".json")):
        # First time we see this design bed file
        tsv_to_json(
            design_bed_filename, design_bed_filename.with_suffix(".json"), "region_id"
        )

    aggregate: dict[str, dict[str, str]] = {}

    with open(design_bed_filename.with_suffix(".json")) as design_bed_file:
        design_bed = json.load(design_bed_file)

    patient_ids = [(p.stem, p) for p in input_bed_filenames]

    for region_id in design_bed:
        if region_id not in aggregate:
            aggregate[region_id] = {}
        aggregate[region_id]["gene"] = design_bed[region_id]["GENE"]
        aggregate[region_id][
            "locus"
        ] = f'{design_bed[region_id]["contig_id"]}:{design_bed[region_id]["contig_srt"]}-{design_bed[region_id]["contig_end"]}'
        for patient_id, patient_bed_path in patient_ids:
            with open(patient_bed_path.with_suffix(".json")) as input_bed_file:
                input_bed: dict[str, dict[str, str]] = json.load(input_bed_file)
            if region_id in input_bed:
                aggregate[region_id][patient_id] = input_bed[region_id]["total_reads"]
            else:
                aggregate[region_id][patient_id] = 0

    # Output aggregate to output_filename (as TSV)
    with open(output_filename, "w") as output_file:
        output_file.write("Gene\tregion_id")
        for patient_id, _ in patient_ids:
            output_file.write(f"\t{patient_id}")
        output_file.write("\n")
        for region_id, region_data in aggregate.items():
            output_file.write(f"{region_data['gene']}\t{region_id}")
            for patient_id, _ in patient_ids:
                output_file.write(f"\t{region_data.get(patient_id, 0)}")
            output_file.write("\n")


def cnv_script_karim(
    input_zip, workdir, refdir, duplication_threshold=1.76, deletion_threshold=0.5
):
    """
    Created on Mon Apr 10 20:22:44 2020

    Script CNV Somatique

    @author: Karim Diallo
    """
    import glob
    import os
    import re
    import shutil
    import time
    import zipfile

    import matplotlib.pyplot as plt
    import numpy
    import xlrd
    import xlsxwriter

    # Used to save the file as excel workbook
    # Need to install this library
    import xlwt

    ##############################################################################
    tmpsDebut = time.time()
    ##############################################################################
    print("\n************ Script CNV secteur SO GENEXUS panel AP-HP V1 *********\n\n")
    #####################################################################
    #####################################################################
    ######################################################################################## LECTURE Tous les fichiers coverage de tous les patients
    #######################################################################################  CREATION d'un fichier entree intermediaire
    ##### LECTURE DU FICHIER ZIP pour extraire les répertoires contenant les fichiers

    ###### NOM DU RUN DANS FICHIER DE SORTIE
    nameRUN = os.path.basename(input_zip).split(".")[0]

    resultats_amp_dir = os.path.join(workdir, "resultats_AMP")
    os.makedirs(resultats_amp_dir, exist_ok=True)

    resultats_gene_dir = os.path.join(workdir, "resultats_Gene")
    os.makedirs(resultats_gene_dir, exist_ok=True)

    filePatientALL = workdir + "/fichierEntreCNV_ALLpatients.xlsx"

    fichierNomGenes = refdir + "/listeCorrespondancePositionAmpliconGene.xlsx"
    if not os.path.exists(fichierNomGenes):
        raise FileNotFoundError(fichierNomGenes)

    fichierListeOrdonnee = (
        refdir + "/GENEXUS_fichierOrdonneRegionStartGene_PanelAPHP.xlsx"
    )
    if not os.path.exists(fichierListeOrdonnee):
        raise FileNotFoundError(fichierListeOrdonnee)

    fichierOrdonnee = refdir + "/fichierOrdonneRegionStartGene_PanelAPHP.xlsx"
    if not os.path.exists(fichierOrdonnee):
        raise FileNotFoundError(fichierOrdonnee)

    fichierTemoin = (
        refdir + "/Moyenne_NormalizedRead_count_TemoinsPorphyriesGENEXUS.xlsx"
    )
    if not os.path.exists(fichierTemoin):
        raise FileNotFoundError(fichierTemoin)

    ####### Extraction dans le dossier Resultats des fichiers/dossiers contenus dans l'archive
    with zipfile.ZipFile(input_zip, "r") as zip_ref:
        zip_ref.extractall(os.path.join(workdir, "rawData_extractCNV"))

    ###### Obtenir la liste des répertoireset DONC le NOM des patients
    listeRepertoirePatients = os.listdir(os.path.join(workdir, "rawData_extractCNV"))

    # print(listeRepertoirePatients)
    # for i, repName in enumerate(listeRepertoirePatients):
    #     full_dirname_before = os.path.join(workdir, "rawData_extractCNV", repName)
    #     full_dirname_after = os.path.join(
    #         workdir,
    #         "rawData_extractCNV",
    #         re.sub(r"^(AssayDev_\d*-?)?([^_]+)(_.+)?", r"\2", repName),
    #     )
    #     os.rename(full_dirname_before, full_dirname_after)

    #     listeRepertoirePatients

    # # J'ai (peut-être) renommé des dossiers, donc il faut refaire la liste des patients
    # listeRepertoirePatients = os.listdir(os.path.join(workdir, "rawData_extractCNV"))

    ##### lecture liste repertoire des patients
    print("Les fichiers extraits... \n")

    for i, listRepPatient in enumerate(listeRepertoirePatients):
        ###### Obtenir liste des fichier dans le répertoire COURANT

        listeFichierP_repertoire = glob.glob(
            os.path.join(
                workdir, "rawData_extractCNV", listRepPatient, "*.amplicon.cov.xls"
            )
        )

        file_from = os.path.join(
            workdir,
            "rawData_extractCNV",
            listRepPatient,
            os.path.basename(listeFichierP_repertoire[0]),
        )

        file_to = os.path.join(
            workdir,
            "rawData_extractCNV",
            listRepPatient + "." + os.path.basename(listeFichierP_repertoire[0]),
        )

        ######### Je deplace fichier coverage AMPLICON vers repo rawData et je renomme avec Patient ID
        shutil.copy(file_from, file_to)

        listeRepertoirePatients

    #######################ICI fin du script extraction fichier .ZIP

    #########################################################################################################################################
    ###########################################################################################################################################
    ### LECTURE all XLS files
    xlsFilename = glob.glob(f"{workdir+'/rawData_extractCNV'}/*.xls")

    print("\nNB. fichier DONNEES BRUTES :: ", len(xlsFilename))

    ####### LECTURE DE MES FICHIERS CORROMPUS POUR "DE-CORROMPRE"
    pID = 0
    for i, chqFile in enumerate(xlsFilename):
        ####-------------------------------------------------------------------------------------------------------------------------------------------
        # Opening the file using 'utf-16' encoding
        file1 = open(chqFile, "r")
        data = file1.readlines()

        # Creating a workbook object
        xldoc = xlwt.Workbook()
        # Adding a sheet to the workbook object
        nomEchantillon = listeRepertoirePatients[pID]
        sheet: xlwt.Worksheet = xldoc.add_sheet(
            nomEchantillon, cell_overwrite_ok=True
        )  ### Nom de la feuille Excel
        # Iterating and saving the data to sheet
        for i, row in enumerate(data):
            # Two things are done here
            # Removing the '\n' which comes while reading the file using open
            # Getting the values after splitting using '\t'

            for j, val in enumerate(row.replace("\n", "").split("\t")):
                sheet.write(i, j, val)
        pID += 1

        xldoc.save(chqFile)

    ##### LECTURE FICHIERS XLSX four faire mon fichier global d'entre script CNV
    fichierEntreCNV: xlsxwriter.Workbook = xlsxwriter.Workbook(filePatientALL)
    worksheetEntreCNV = fichierEntreCNV.add_worksheet("matrix_coverageALL")

    #### Ecrire colonne zero et une
    worksheetEntreCNV.write(0, 0, "Gene")
    worksheetEntreCNV.write(0, 1, "region_id")

    ###### LIRE TOUS LES FICHIERS PATIENTS xlsx
    filenameXLSX = glob.glob(f"{workdir}/rawData_extractCNV/*.xls")

    ##### FAIRE LA LISTE ORDONNEE AMPLICON les dictionnaires AMPLI-CHR et GENE

    workbook_listOrdonnee: xlrd.Book = xlrd.open_workbook(fichierListeOrdonnee)
    sheet_listOrdonnee = workbook_listOrdonnee.sheet_by_index(0)

    listOrdonneeAMP = []
    dictAmpliGene = {}
    dictAmpliLocus = {}
    dictAmpliChr = {}

    for l_ordonnee in range(1, sheet_listOrdonnee.nrows):
        listOrdonneeAMP.append(
            str(sheet_listOrdonnee.cell_value(l_ordonnee, 3)).replace(".0", "")
        )
        dictAmpliChr[
            str(sheet_listOrdonnee.cell_value(l_ordonnee, 3)).replace(".0", "")
        ] = sheet_listOrdonnee.cell_value(l_ordonnee, 0)
        dictAmpliGene[
            str(sheet_listOrdonnee.cell_value(l_ordonnee, 3)).replace(".0", "")
        ] = sheet_listOrdonnee.cell_value(l_ordonnee, 5)
        dictAmpliLocus[
            str(sheet_listOrdonnee.cell_value(l_ordonnee, 3)).replace(".0", "")
        ] = (
            sheet_listOrdonnee.cell_value(l_ordonnee, 0)
            + ":"
            + str(int(sheet_listOrdonnee.cell_value(l_ordonnee, 1)))
            + "-"
            + str(int(sheet_listOrdonnee.cell_value(l_ordonnee, 2)))
        )

    print("NB. amplicon PANEL ::: ", len(listOrdonneeAMP))

    #
    ##### lire chaque fichier dans la liste et FAIRE la matrix de coverage
    compteNbPatient = 2  ###☺ Nombre patient liste patient Attention EGAL 2 car il y a 3 colonnes "gene, region_id, et locus" quand j'écris dans le matrix
    for i, monFichierMatrix in enumerate(filenameXLSX):
        compteNbPatient += 1
        print(
            "Patient :: ",
            os.path.basename(monFichierMatrix).split(".")[0],
        )  ####SPLIT "_IonDual" pour récuprer le NOM du PATIENT ")
        monID_patient = os.path.basename(monFichierMatrix).split(".")[0]

        #### Lecture XLSX file
        myXLSX_workbook: xlrd.Book = xlrd.open_workbook(monFichierMatrix)
        myXLSXsheet_1 = myXLSX_workbook.sheet_by_index(0)

        ####### REMPLIR MA MATRICE FILE
        worksheetEntreCNV.write(0, compteNbPatient, monID_patient)
        #### dictionnaire du patient EN cours
        dictPatientEnCours = {}
        for ligneXLSX in range(1, myXLSXsheet_1.nrows):
            dictPatientEnCours[myXLSXsheet_1.cell_value(ligneXLSX, 3)] = (
                myXLSXsheet_1.cell_value(ligneXLSX, 9)
            )  #### CREATION dict AMPLI-READS

        lesLignes = 1
        for ampli_LO in listOrdonneeAMP:
            worksheetEntreCNV.write(
                lesLignes, 0, dictAmpliGene[ampli_LO]
            )  ################ Remplir colonne GENE (colonne zéro donc)
            worksheetEntreCNV.write(
                lesLignes, 1, ampli_LO
            )  ################ Remplir colonne AMPLICON
            worksheetEntreCNV.write_formula(
                lesLignes,
                2,
                f"""=HYPERLINK("localhost:60151/goto?locus={dictAmpliLocus[ampli_LO]}","{dictAmpliLocus[ampli_LO]}")""",
            )  #####○ Ajouter un lien vers IGV
            worksheetEntreCNV.write(
                lesLignes, compteNbPatient, int(dictPatientEnCours[ampli_LO])
            )  #####○ Remplir Total reads
            lesLignes += 1

    fichierEntreCNV.close()

    #######################################################################################################################################
    ########################################################################################################################################
    ######################################################################################################################################### CNV CNV CNV
    ##########################################################################################################################################
    ############################################################################################################################################

    ####### A partir d'ici PRENDRE LE FICHIER GLOBAL DES COUVERTURES pour FILE Patients CNV du dessus

    f_ALL: xlrd.Book = xlrd.open_workbook(filePatientALL)
    feui_ALL = f_ALL.sheet_by_index(0)

    ##Nom fichier SORTIE
    workbook = xlsxwriter.Workbook(
        os.path.join(resultats_amp_dir, "Resultat_Ratio_" + nameRUN + ".xlsx")
    )

    #####################################################################

    def somme_colonnePatientX(fichierP, no_col):
        fich_col: xlrd.Book = xlrd.open_workbook(fichierP)
        feuil_col = fich_col.sheet_by_index(0)
        somme_colPatX = 0
        for i in range(1, feuil_col.nrows):
            somme_colPatX += feuil_col.cell_value(i, no_col)
        return somme_colPatX

    def moy_Norm_TemoinsPorphy(fichierTemoin):
        fichier1: xlrd.Book = xlrd.open_workbook(fichierTemoin)
        feuil1 = fichier1.sheet_by_index(0)
        dict_moy_Norm_TemoinsPorphy = {}
        for li in range(1, feuil1.nrows):
            dict_moy_Norm_TemoinsPorphy[
                str(feuil1.cell_value(li, 0)).replace(".0", "")
            ] = feuil1.cell_value(li, 1)

        return dict_moy_Norm_TemoinsPorphy

    # calcul function temoin porphyrie
    dict_moy_Norm_TemoinsPorphy = moy_Norm_TemoinsPorphy(fichierTemoin)

    def moy_Norm_dict_patient(filePat, numPat):
        fichier2: xlrd.Book = xlrd.open_workbook(filePat)
        feuil2 = fichier2.sheet_by_index(0)
        dictPatient = {}
        for li in range(1, feuil2.nrows):
            dictPatient[str(feuil2.cell_value(li, 1)).replace(".0", "")] = (
                feuil2.cell_value(li, numPat)
                / (
                    somme_colonnePatientX(filePat, numPat)
                    - feuil2.cell_value(li, numPat)
                )
            )

        return dictPatient

    ####Amplicon ordonnée
    fichier1: xlrd.Book = xlrd.open_workbook(fichierOrdonnee)
    feuil1 = fichier1.sheet_by_index(0)
    listeOrdonneeAmpli = []
    for li in range(1, feuil1.nrows):
        listeOrdonneeAmpli.append(str(feuil1.cell_value(li, 1)).replace(".0", ""))

    def dictChrm(filePati):
        fichier2: xlrd.Book = xlrd.open_workbook(filePati)
        feuil2 = fichier2.sheet_by_index(0)
        dictChromosome = {}
        for li in range(1, feuil2.nrows):
            dictChromosome[str(feuil2.cell_value(li, 1)).replace(".0", "")] = (
                feuil2.cell_value(li, 0)
            )
        return dictChromosome

    dictChromosome = dictChrm(fichierOrdonnee)

    ##Ecrire fichier de sortie normalisation
    worksheet = workbook.add_worksheet()
    # Iterate over the data and write it out row by row.
    #########
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({"bold": True})
    worksheet.write("A1", "Chr", bold)
    worksheet.write("B1", "AmpliconID", bold)

    cell_formatRED = workbook.add_format()

    cell_formatRED.set_font_color("red")

    cell_formatBLUE = workbook.add_format()

    cell_formatBLUE.set_font_color("blue")

    numpatient = 3
    varPat = 3
    for i in range(3, feui_ALL.ncols):

        #####calcul dictRatio
        dictRatio = {}

        moy_Norm_dict_patientEC = moy_Norm_dict_patient(
            filePatientALL, numpatient
        )  ###APPEL DE LA FONCTION patient all
        for kle, val in moy_Norm_dict_patientEC.items():
            dictRatio[kle] = round(
                moy_Norm_dict_patientEC[kle] / dict_moy_Norm_TemoinsPorphy[kle], 3
            )

        ## Add a number format for cells with xxx.
        row = 1
        col = 0
        for item in listeOrdonneeAmpli:
            if dictRatio[item] < deletion_threshold:
                worksheet.write(row, col, dictChromosome[item])
                worksheet.write(row, col + 1, item)
                worksheet.write(row, col + varPat, dictRatio[item], cell_formatRED)
                row += 1
            elif dictRatio[item] >= duplication_threshold:
                worksheet.write(row, col, dictChromosome[item])
                worksheet.write(row, col + 1, item)
                worksheet.write(row, col + varPat, dictRatio[item], cell_formatBLUE)
                row += 1
            else:
                worksheet.write(row, col, dictChromosome[item])
                worksheet.write(row, col + 1, item)
                worksheet.write(row, col + varPat, dictRatio[item])
                row += 1

        worksheet.write(0, varPat, feui_ALL.cell_value(0, varPat))

        ######
        ####### Les seuils 1.25dup & 0.7 dél  !!!

        ####################################
        with open(
            os.path.join(
                resultats_amp_dir, "Fichier_Anomalies des patients_" + nameRUN + ".txt"
            ),
            "a",
        ) as mon_fichier:
            mon_fichier.write(
                "********Résultats récap...patient : {} \n".format(
                    feui_ALL.cell_value(0, varPat)
                )
            )
            for cle in listeOrdonneeAmpli:
                if dictRatio[cle] < deletion_threshold:
                    mon_fichier.write(
                        "{}; Délétion sur amplicon : {} son ratio = {} \n".format(
                            dictChromosome[cle], cle, dictRatio[cle]
                        )
                    )
            for cle in listeOrdonneeAmpli:
                if dictRatio[cle] >= duplication_threshold:
                    mon_fichier.write(
                        "{}; Duplication sur amplicon : {} son ratio = {} \n".format(
                            dictChromosome[cle], cle, dictRatio[cle]
                        )
                    )
            mon_fichier.write("\n-----------------\n")

        ## Create  data
        seq1y = []
        seq2y = []
        seq3y = []
        ind1 = []
        ind2 = []
        ind3 = []
        var = 0
        for exID in listeOrdonneeAmpli:
            var += 1
            for i, j in dictRatio.items():
                if exID == i and j < deletion_threshold:
                    seq1y.append(j)
                    ind1.append(var)
                elif exID == i and j >= duplication_threshold:
                    seq2y.append(j)
                    ind2.append(var)
                elif exID == i:
                    seq3y.append(j)
                    ind3.append(var)

        g1 = (ind1, seq1y)
        g2 = (ind2, seq2y)
        g3 = (ind3, seq3y)
        #
        data = (g1, g2, g3)
        colors = ("red", "blue", "green")
        groups = ("Deletion", "Duplication", "Normal")

        # Create plot
        fig = plt.figure()
        ax = fig.add_subplot(1, 1, 1)

        for data, color, group in zip(data, colors, groups):
            x, y = data
            ax.scatter(x, y, alpha=0.8, c=color, edgecolors="none", s=30, label=group)

        plt.xlabel("Amplicon-ID")
        plt.ylabel("Ratio")

        ax.plot(
            [0, 340],
            [deletion_threshold, deletion_threshold],
            color="black",
            linestyle="solid",
        )
        ax.plot(
            [0, 340],
            [duplication_threshold, duplication_threshold],
            color="black",
            linestyle="solid",
        )
        ax.plot([0, 340], [2.4, 2.4], color="black", linestyle="dashdot")
        kurs = "%s.png" % feui_ALL.cell_value(0, varPat)
        plt.title(kurs)

        plt.savefig(os.path.join(resultats_amp_dir, kurs), format="png")

        fig.clf()
        plt.close()

        numpatient += 1
        varPat += 1

    workbook.close()
    #
    #################################################################################################################################
    #################################################################################################################################
    #################################################################################################################################
    #############################################  MOYENNE RATIO
    fileMoyenneRatio = os.path.join(
        resultats_amp_dir, "Resultat_Ratio_" + nameRUN + ".xlsx"
    )
    rbFileMoyRatio: xlrd.Book = xlrd.open_workbook(fileMoyenneRatio)
    feui_FMR = rbFileMoyRatio.sheet_by_index(0)

    open_fichierNomGenes: xlrd.Book = xlrd.open_workbook(fichierNomGenes)
    feuille_fichierNomGenes = open_fichierNomGenes.sheet_by_index(0)

    listeNomGene0 = []
    for liFNG in range(1, feuille_fichierNomGenes.nrows):
        listeNomGene0.append(
            str(feuille_fichierNomGenes.cell_value(liFNG, 4)).replace(".0", "")
        )

    listeNomGene = set(listeNomGene0)

    ##### LES GENES IDENTITOVIGILANCES
    listeGeneNonInterets = [
        "PENTA",
        "224830378",
        "224869380",
        "224862488",
        "224879644",
        "224829586",
        "224851112",
        "TH01",
        "224825605",
        "224824531",
        "AMEX",
        "AMEY",
        "D1MS201754411",
        "MON27",
        "BAT26",
        "D2MS62063094",
        "NR24",
        "BAT25",
        "D5MS172421761",
        "D6MS142691951",
        "D7MS1787520",
        "D7MS74608741",
        "D11MS106695515",
        "D13MS31722621",
        "NR21",
        "D15MS45897772",
        "D16MS18882660",
        "D17MS19314918",
    ]

    ######fichier excel de sortie
    workbook_RG = xlsxwriter.Workbook(
        os.path.join(resultats_gene_dir, "Resultat_MeanRatioGene_" + nameRUN + ".xlsx")
    )
    worksheet_rg = workbook_RG.add_worksheet("Reustats_moyenneRatioParGene")
    bold = workbook_RG.add_format({"bold": True})

    cell_formatRED = workbook_RG.add_format()

    cell_formatRED.set_font_color("red")

    cell_formatBLUE = workbook_RG.add_format()

    cell_formatBLUE.set_font_color("blue")

    ####remplir 1 ligne titre des colonnes
    for erLigne in range(0, feui_FMR.ncols):
        worksheet_rg.write(0, erLigne, feui_FMR.cell_value(0, erLigne), bold)
        worksheet_rg.write(0, 1, "Gene", bold)

    ########☻parcours les colonnes de mon fichier par patient
    numPat = 3
    numpatient = 3
    for patientNum in range(3, feui_FMR.ncols):
        dictRatioParGene = {}
        dictChromoso = {}
        for gene in listeNomGene:
            listeGene = []
            for ifmr in range(1, feui_FMR.nrows):
                if str(feui_FMR.cell_value(ifmr, 1)).replace(".0", "").find(gene) != -1:
                    listeGene.append(feui_FMR.cell_value(ifmr, numPat))
                    dictChromoso[gene] = feui_FMR.cell_value(ifmr, 0)
            dictRatioParGene[gene] = numpy.mean(listeGene)

        #######Ecriture fichier des ratios et fichier anomalie

        with open(
            os.path.join(
                resultats_gene_dir, "Fichier_Anomalies des patients_" + nameRUN + ".txt"
            ),
            "a",
        ) as my_fichier:
            my_fichier.write(
                "********Résultats récap...patient : {} \n".format(
                    feui_FMR.cell_value(0, numPat)
                )
            )

            ####
            for geneName in listeNomGene:
                if (
                    dictRatioParGene[geneName] < deletion_threshold
                    and geneName not in listeGeneNonInterets
                ):
                    my_fichier.write(
                        "{} - Délétion gene : {} avec ratio = {} \n".format(
                            dictChromoso[geneName], geneName, dictRatioParGene[geneName]
                        )
                    )

            ###
            for geneName in listeNomGene:
                if (
                    dictRatioParGene[geneName] >= duplication_threshold
                    and geneName not in listeGeneNonInterets
                ):
                    my_fichier.write(
                        "{} - Duplication gene : {} avec ratio = {} \n".format(
                            dictChromoso[geneName], geneName, dictRatioParGene[geneName]
                        )
                    )

            ligne = 1
            for geneName in listeNomGene:
                if (
                    dictRatioParGene[geneName] < deletion_threshold
                    and geneName not in listeGeneNonInterets
                ):
                    worksheet_rg.write(ligne, 0, dictChromoso[geneName])
                    worksheet_rg.write(ligne, 1, geneName)
                    worksheet_rg.write(
                        ligne, numPat, dictRatioParGene[geneName], cell_formatRED
                    )

                elif (
                    dictRatioParGene[geneName] >= duplication_threshold
                    and geneName not in listeGeneNonInterets
                ):
                    worksheet_rg.write(ligne, 0, dictChromoso[geneName])
                    worksheet_rg.write(ligne, 1, geneName)
                    worksheet_rg.write(
                        ligne, numPat, dictRatioParGene[geneName], cell_formatBLUE
                    )

                else:
                    worksheet_rg.write(ligne, 0, dictChromoso[geneName])
                    worksheet_rg.write(ligne, 1, geneName)
                    worksheet_rg.write(ligne, numPat, dictRatioParGene[geneName])

                ligne += 1
            my_fichier.write("\n-----------------\n")

            numPat += 1

        #####GRAPHIQUE made
        # Create  data
        seq1y = []
        seq2y = []
        seq3y = []
        ind1 = []
        ind2 = []
        ind3 = []
        abcisse = 1
        dictSticks = {}
        for geneNom in listeNomGene:
            if (
                dictRatioParGene[geneNom] < deletion_threshold
                and geneNom not in listeGeneNonInterets
            ):
                seq1y.append(dictRatioParGene[geneNom])
                ind1.append(abcisse)
                dictSticks[geneNom] = abcisse
            elif (
                dictRatioParGene[geneNom] >= duplication_threshold
                and geneNom not in listeGeneNonInterets
            ):
                seq2y.append(dictRatioParGene[geneNom])
                ind2.append(abcisse)
                dictSticks[geneNom] = abcisse
            elif (
                dictRatioParGene[geneNom] > deletion_threshold
                and dictRatioParGene[geneNom] < duplication_threshold
                and geneNom not in listeGeneNonInterets
            ):
                seq3y.append(dictRatioParGene[geneNom])
                ind3.append(abcisse)
            abcisse += 1

        g1 = (ind1, seq1y)
        g2 = (ind2, seq2y)
        g3 = (ind3, seq3y)

        data = (g1, g2, g3)
        colors = ("red", "blue", "green")
        groups = ("Deletion", "Duplication", "Normal")

        # Create plot
        fig = plt.figure()
        ax = fig.add_subplot(1, 1, 1)

        for data, color, group in zip(data, colors, groups):
            x, y = data
            ax.scatter(x, y, alpha=0.8, c=color, edgecolors="none", s=30, label=group)

        plt.xlabel("Gene")
        plt.ylabel("Ratio mean")

        ax.plot(
            [0, 80],
            [deletion_threshold, deletion_threshold],
            color="black",
            linestyle="solid",
        )
        ax.plot(
            [0, 80],
            [duplication_threshold, duplication_threshold],
            color="black",
            linestyle="solid",
        )
        ax.plot([0, 80], [2.4, 2.4], color="black", linestyle="dashdot")
        kurs = "%s.png" % feui_FMR.cell_value(0, numpatient)

        plt.title(kurs)
        plt.xticks(
            list(dictSticks.values()),
            list(dictSticks.keys()),
            rotation="vertical",
            fontsize=8,
        )

        plt.savefig(os.path.join(resultats_gene_dir, kurs), format="png")
        fig.clf()
        plt.close()

        numpatient += 1

    workbook_RG.close()

    print("\nSuccessful....OK\n")
    tmpsFin = time.time() - tmpsDebut
    print("Temps d'execution en s = %f" % tmpsFin)


if __name__ == "__main__":
    sys.exit(main())
