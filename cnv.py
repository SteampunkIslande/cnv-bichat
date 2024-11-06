#!/usr/bin/env python3
"""
Created on Fri Sept 27 2024
Adaptation à Linux et passage de python2 à python3 du script de Karim Diallo

Ajout d'une interface en ligne de commande.
Suppression du code neutralisé par commentaires.
Légère refactorisation du code.

Mise à jour le 5 novembre 2024

Refactorisation du code pour le rendre plus modulaire et plus facile à tester.

"""

import os
from pathlib import Path

import duckdb as db

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
        "--is-reference-run",
        help="Outputs result as a reference BED file",
        action="store_true",
    )
    args = parser.parse_args()

    return call(
        args.input_files,
        args.workdir,
        args.design_bed,
        args.reference_coverage_bed,
        args.duplication_threshold,
        args.deletion_threshold,
        args.is_reference_run,
    )


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


def call(
    input_files: list[Path],
    workdir: Path,
    design_bed: Path,
    reference_coverage_bed_filename: Path,
    duplication_threshold,
    deletion_threshold,
    is_reference_run=False,
):

    if not design_bed.exists():
        raise FileNotFoundError(design_bed)

    if is_reference_run:
        if reference_coverage_bed_filename.exists():
            raise FileExistsError(reference_coverage_bed_filename)
    else:
        if not reference_coverage_bed_filename:
            raise ValueError(
                "Reference coverage bed file is required for non-reference run"
            )
        if not reference_coverage_bed_filename.exists():
            raise FileNotFoundError(reference_coverage_bed_filename)

    # Avoid re-running the same analysis
    if workdir.exists():
        print(f"Output directory {workdir} already exists. Skipping analysis.")
        return 1

    zip_files = [f for f in input_files if f.suffix == ".zip"]

    # Any other file is considered an amplicon coverage file in bed format
    bed_files = [f for f in input_files if f.suffix != ".zip"]

    for input_zip_file in zip_files:
        input_zip_file = Path(input_zip_file)
        input_bed_filenames = [
            workdir / "extracted" / p
            for p in extract_amplicon_files_from_zip(
                input_zip_file, os.path.join(workdir, "extracted")
            )
        ]
        cnv_call(
            input_bed_filenames,
            design_bed,
            workdir,
            reference_coverage_bed_filename,
            duplication_threshold,
            deletion_threshold,
            is_reference_run,
        )

    if bed_files:

        cnv_call(
            bed_files,
            design_bed,
            workdir,
            reference_coverage_bed_filename,
            duplication_threshold,
            deletion_threshold,
            is_reference_run,
        )


def build_reference_coverage_table(
    normalized_depth_table: db.DuckDBPyRelation,
    design_bed_table: db.DuckDBPyRelation,
):
    return db.sql(
        "SELECT design_bed_table.region_id,avg(normalized_depth_table.normalized_depth),index AS avg_normalized_depth FROM normalized_depth_table JOIN design_bed_table ON design_bed_table.region_id=normalized_depth_table.region_id GROUP BY region_id"
    )


def normalize_coverage_table(
    coverage_table: db.DuckDBPyRelation,
    total_read_counts_per_sample: db.DuckDBPyRelation,
):
    return db.sql(
        "SELECT coverage_table.*,coverage_table.total_reads / (tot.total_reads_sum - coverage_table.total_reads) AS normalized_depth FROM coverage_table JOIN total_read_counts_per_sample tot ON coverage_table.sample_name = tot.sample_name"
    )


def pivoted_amplicon_ratio_table(
    ratio_table: db.DuckDBPyRelation,
    ref_table: db.DuckDBPyRelation,
):
    return db.sql(
        """SELECT * EXCLUDE(index) FROM 
        (PIVOT
            (SELECT ref_table.contig_id,ref_table.region_id,ratio_table.ratio,ratio_table.sample_name,index FROM ref_table JOIN ratio_table on ratio_table.region_id=ref_table.region_id
            ) ON sample_name USING first(ratio)
        ) ORDER BY index"""
    )


def pivoted_gene_ratio_table(
    ratio_table: db.DuckDBPyRelation,
    ref_table: db.DuckDBPyRelation,
):
    return db.sql(
        "SELECT * EXCLUDE(index) FROM (SELECT contig_id,GENE,avg(COLUMNS(* EXCLUDE(GENE,contig_id,region_id))) FROM (PIVOT (SELECT ref_table.GENE,ref_table.contig_id,ref_table.region_id,ratio_table.ratio,ratio_table.sample_name,index FROM ref_table JOIN ratio_table on ratio_table.region_id=ref_table.region_id) ON sample_name USING first(ratio)) GROUP BY (GENE,contig_id)) ORDER BY index"
    )


def export_duckdb_table_to_excel(
    table: db.DuckDBPyRelation,
    output_filename: Path,
    deletion_threshold=0.5,
    duplication_threshold=1.76,
):
    rows = table.fetchall()
    import xlsxwriter

    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet()

    cell_formatBOLD = workbook.add_format({"bold": True})

    cell_formatRED = workbook.add_format({"font_color": "red"})
    cell_formatBLUE = workbook.add_format({"font_color": "blue"})

    col_names = table.columns
    for i, col_name in enumerate(col_names):
        worksheet.write(0, i, col_name, cell_formatBOLD)

    for i, row in enumerate(rows, start=1):
        for j, cell in enumerate(row):
            col_name = col_names[j]
            # That's how I identify the sample name column (containing the ratio values)
            if col_name not in ("index", "region_id", "contig_id", "GENE"):
                cell = round(cell, 3)
                if cell < deletion_threshold:
                    worksheet.write(i, j, cell, cell_formatRED)
                elif cell >= duplication_threshold:
                    worksheet.write(i, j, cell, cell_formatBLUE)
                else:
                    worksheet.write(i, j, cell)
            else:
                worksheet.write(i, j, cell)

    workbook.close()


def cnv_call(
    input_bed_filenames: list[Path],
    design_bed_filename,
    workdir,
    reference_coverage_bed_filename,
    duplication_threshold,
    deletion_threshold,
    is_reference_run=False,
):
    input_bed_filenames = "[" + ", ".join([f"'{p}'" for p in input_bed_filenames]) + "]"

    ref_table = db.sql(
        f"SELECT row_number() OVER () AS index,contig_id,GENE,region_id FROM read_csv('{design_bed_filename}',sep='\t')"
    )
    cov_table = db.sql(
        f"SELECT region_id,total_reads,parse_filename(filename,True) AS sample_name FROM read_csv({input_bed_filenames},sep='\t',filename=True,union_by_name=True)"
    )
    aggregate_samples_coverage(
        ref_table, cov_table, workdir / "all_samples_coverage.tsv"
    )
    total_read_counts_per_sample = db.sql(
        "SELECT sample_name,sum(total_reads) AS total_reads_sum FROM cov_table GROUP BY sample_name"
    )
    normalized_depth_table = normalize_coverage_table(
        cov_table, total_read_counts_per_sample
    )
    if is_reference_run:
        reference_coverage_table = build_reference_coverage_table(
            normalized_depth_table,
            ref_table,
        )
        db.sql(
            f"COPY (SELECT * EXCLUDE(index) FROM reference_coverage_table ORDER BY index) TO '{reference_coverage_bed_filename}' (DELIMITER '\t')"
        )
        return 0
    else:
        reference_coverage_table = db.sql(
            f"SELECT row_number() OVER () AS index,region_id,avg_normalized_depth FROM read_csv('{reference_coverage_bed_filename}',sep='\t')"
        )
        ratio_table = db.sql(
            "SELECT reference_coverage_table.region_id,sample_name,normalized_depth/avg_normalized_depth AS ratio FROM reference_coverage_table JOIN normalized_depth_table ON reference_coverage_table.region_id=normalized_depth_table.region_id"
        )

        final_amplicon_ratio_table = pivoted_amplicon_ratio_table(
            ratio_table, ref_table
        )
        final_gene_ratio_table = pivoted_gene_ratio_table(ratio_table, ref_table)

        export_duckdb_table_to_excel(
            final_amplicon_ratio_table, workdir / "amplicon_ratio.xlsx"
        )
        export_duckdb_table_to_excel(
            final_gene_ratio_table, workdir / "gene_ratio.xlsx"
        )


def aggregate_samples_coverage(
    ref_table: db.DuckDBPyRelation,
    cov_table: db.DuckDBPyRelation,
    output_filename: Path,
):
    db.sql(
        f"COPY (SELECT * EXCLUDE(index) FROM (PIVOT (SELECT ref.index,ref.GENE,ref.contig_id,ref.region_id,cov.total_reads,cov.sample_name FROM ref_table ref JOIN cov_table cov ON cov.region_id=ref.region_id) ON sample_name USING first(total_reads) ) ORDER BY index) TO '{output_filename}' (DELIMITER '\t')"
    )


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
    worksheet.write("B1", "region_id", bold)

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
    import sys

    sys.exit(main())
