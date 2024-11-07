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
        args.deletion_threshold,
        args.duplication_threshold,
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
    deletion_threshold,
    duplication_threshold,
    is_reference_run=False,
):

    if not design_bed.exists():
        raise FileNotFoundError(design_bed)

    if not reference_coverage_bed_filename:
        raise ValueError(
            "Reference coverage bed file is required, whether you wish to create one or use an existing one to call CNVs."
        )

    if is_reference_run:
        if reference_coverage_bed_filename.exists():
            raise FileExistsError(reference_coverage_bed_filename)
    else:
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
            deletion_threshold,
            duplication_threshold,
            is_reference_run,
        )

    if bed_files:

        cnv_call(
            bed_files,
            design_bed,
            workdir,
            reference_coverage_bed_filename,
            deletion_threshold,
            duplication_threshold,
            is_reference_run,
        )


def build_reference_coverage_table(
    normalized_depth_table: db.DuckDBPyRelation,
    design_bed_table: db.DuckDBPyRelation,
):
    return db.sql(
        f"""SELECT design_bed_table.region_id,avg(normalized_depth_table.normalized_depth),index AS avg_normalized_depth FROM ({normalized_depth_table.sql_query()}) normalized_depth_table JOIN ({design_bed_table.sql_query()}) design_bed_table ON design_bed_table.region_id=normalized_depth_table.region_id GROUP BY region_id"""
    )


def normalize_coverage_table(
    coverage_table: db.DuckDBPyRelation,
    total_read_counts_per_sample: db.DuckDBPyRelation,
):
    return db.sql(
        f"""SELECT coverage_table.*,coverage_table.total_reads / (tot.total_reads_sum - coverage_table.total_reads) AS normalized_depth FROM ({coverage_table.sql_query()}) coverage_table JOIN ({total_read_counts_per_sample.sql_query()}) tot ON coverage_table.sample_name = tot.sample_name"""
    )


def pivoted_amplicon_ratio_table(
    ratio_table: db.DuckDBPyRelation,
    ref_table: db.DuckDBPyRelation,
):
    return db.sql(
        f"""SELECT * EXCLUDE(index) FROM 
        (PIVOT
            (SELECT ref_table.contig_id,ref_table.region_id,ratio_table.ratio,ratio_table.sample_name,index FROM ({ref_table.sql_query()}) ref_table JOIN ({ratio_table.sql_query()}) ratio_table ON ratio_table.region_id=ref_table.region_id
            ) ON sample_name USING first(ratio)
        ) ORDER BY index"""
    )


def pivoted_gene_ratio_table(
    ratio_table: db.DuckDBPyRelation,
    ref_table: db.DuckDBPyRelation,
):
    return db.sql(
        f"SELECT * EXCLUDE(index) FROM (SELECT contig_id,GENE,avg(COLUMNS(* EXCLUDE(GENE,contig_id,region_id))) FROM (PIVOT (SELECT ref_table.GENE,ref_table.contig_id,ref_table.region_id,ratio_table.ratio,ratio_table.sample_name,index FROM ({ref_table.sql_query()}) ref_table JOIN ({ratio_table.sql_query()}) ratio_table ON ratio_table.region_id=ref_table.region_id) ON sample_name USING first(ratio)) GROUP BY (GENE,contig_id)) ORDER BY index"
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
    deletion_threshold,
    duplication_threshold,
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
        f"""SELECT sample_name,sum(total_reads) AS total_reads_sum FROM ({cov_table.sql_query()}) cov_table GROUP BY sample_name"""
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
            f"COPY (SELECT * EXCLUDE(index) FROM ({reference_coverage_table.sql_query()}) reference_coverage_table ORDER BY index) TO '{reference_coverage_bed_filename}' (DELIMITER '\t')"
        )
        return 0
    else:
        reference_coverage_table = db.sql(
            f"SELECT row_number() OVER () AS index,region_id,avg_normalized_depth FROM read_csv('{reference_coverage_bed_filename}',sep='\t')"
        )
        ratio_table = db.sql(
            f"""SELECT reference_coverage_table.region_id,sample_name,normalized_depth/avg_normalized_depth AS ratio FROM ({reference_coverage_table.sql_query()}) reference_coverage_table JOIN ({normalized_depth_table.sql_query()}) normalized_depth_table ON reference_coverage_table.region_id=normalized_depth_table.region_id"""
        )

        final_amplicon_ratio_table = pivoted_amplicon_ratio_table(
            ratio_table, ref_table
        )
        final_gene_ratio_table = pivoted_gene_ratio_table(ratio_table, ref_table)

        export_duckdb_table_to_excel(
            final_amplicon_ratio_table,
            workdir / "amplicon_ratio.xlsx",
            deletion_threshold,
            duplication_threshold,
        )
        export_duckdb_table_to_excel(
            final_gene_ratio_table,
            workdir / "gene_ratio.xlsx",
            deletion_threshold,
            duplication_threshold,
        )

        write_cnv_report(
            final_amplicon_ratio_table,
            ["region_id"],
            workdir / "cnv_amplicons_report.html",
            deletion_threshold,
            duplication_threshold,
            x_axis_label="Amplicon",
            y_axis_label="Ratio",
        )

        write_cnv_report(
            final_gene_ratio_table,
            ["GENE"],
            workdir / "cnv_genes_report.html",
            deletion_threshold,
            duplication_threshold,
            x_axis_label="Gene",
            y_axis_label="Ratio",
            x_ticks_as_labels=True,
        )


def aggregate_samples_coverage(
    ref_table: db.DuckDBPyRelation,
    cov_table: db.DuckDBPyRelation,
    output_filename: Path,
):
    db.sql(
        f"COPY (SELECT * EXCLUDE(index) FROM (PIVOT (SELECT ref.index,ref.GENE,ref.contig_id,ref.region_id,cov.total_reads,cov.sample_name FROM ({ref_table.sql_query()}) ref JOIN ({cov_table.sql_query()}) cov ON cov.region_id=ref.region_id) ON sample_name USING first(total_reads) ) ORDER BY index) TO '{output_filename}' (DELIMITER '\t')"
    )


def write_cnv_report(
    table: db.DuckDBPyRelation,
    col_names: list[str],
    output_filename: Path,
    deletion_threshold: float,
    duplication_threshold: float,
    x_axis_label="Amplicon",
    y_axis_label="Ratio",
    x_ticks_as_labels=False,
):
    from jinja2 import BaseLoader, Environment

    from html_report_template import template

    sample_names = [
        f'"{c}"'
        for c in table.columns
        if c not in ("index", "region_id", "GENE", "contig_id")
    ]

    rtemplate = Environment(loader=BaseLoader).from_string(template)
    data = {
        "samples": {sample_name.replace('"', ""): {} for sample_name in sample_names}
    }
    for sample_name in sample_names:
        unquoted_sample_name = sample_name.replace('"', "")
        deletions = (
            table.select(sample_name, "contig_id", *col_names)
            .filter(f"""{sample_name} < {deletion_threshold}""")
            .pl()
            .to_dicts()
        )

        duplications = (
            table.select(sample_name, "contig_id", *col_names)
            .filter(f"""{sample_name} >= {duplication_threshold}""")
            .pl()
            .to_dicts()
        )
        data["samples"][unquoted_sample_name]["deletion"] = [
            f"{r['contig_id']} - {r[col_names[0]]} : Ratio = {r[unquoted_sample_name]:.3f}"
            for r in deletions
        ]
        data["samples"][unquoted_sample_name]["duplication"] = [
            f"{d['contig_id']} - {d[col_names[0]]} : Ratio = {d[unquoted_sample_name]:.3f}"
            for d in duplications
        ]
        data["samples"][unquoted_sample_name]["graph"] = plot_cnv_results(
            table,
            col_names[0],
            sample_name,
            deletion_threshold,
            duplication_threshold,
            x_axis_label=x_axis_label,
            y_axis_label=y_axis_label,
            title_label=f"{unquoted_sample_name} CNV results",
            x_ticks_as_labels=x_ticks_as_labels,
        )
    with open(output_filename, "w") as f:
        f.write(rtemplate.render(**data))


def plot_cnv_results(
    pivoted_table: db.DuckDBPyRelation,
    column_name: str,
    sample_name: str,
    deletion_threshold=0.5,
    duplication_threshold=1.76,
    x_axis_label="Amplicon",
    y_axis_label="Ratio",
    title_label="CNV results",
    x_ticks_as_labels=False,
):
    # Returns the base64 encoded image as a string
    import base64
    import io

    import matplotlib.pyplot as plt

    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)
    data = (
        db.sql(
            f""" SELECT {column_name},{sample_name},if({sample_name} < {deletion_threshold},'deletion',if({sample_name} >= {duplication_threshold},'duplication','normal')) AS group FROM ({pivoted_table.sql_query()}) pivoted_table"""
        )
        .pl()
        .to_numpy()
    )
    xtick = []
    xtick_labels = []
    for data_ in data:
        x = data_[0]
        y = data_[1]
        group = data_[2]
        color = {
            "deletion": "red",
            "duplication": "green",
            "normal": "blue",
        }[group]
        if x_ticks_as_labels and group != "normal":
            xtick.append(x)
            xtick_labels.append(x)
        ax.scatter([x], [y], alpha=0.8, c=color, edgecolors="none", s=30)

    ax.set_xticks(xtick, xtick_labels, rotation=90)

    ax.set_xlabel(x_axis_label)
    ax.set_ylabel(y_axis_label)
    ax.set_title(title_label)

    x_lim = ax.get_xlim()

    ax.plot(
        x_lim,
        [deletion_threshold, deletion_threshold],
        color="black",
        linestyle="solid",
    )
    ax.plot(
        x_lim,
        [duplication_threshold, duplication_threshold],
        color="black",
        linestyle="solid",
    )
    ax.plot(x_lim, [2.4, 2.4], color="black", linestyle="dashdot")

    plt.autoscale()

    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    result = base64.b64encode(buf.read()).decode("utf-8")
    plt.close()
    return result


if __name__ == "__main__":
    import sys

    sys.exit(main())
