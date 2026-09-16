import os
import json
import re
import zipfile

import pandas as pd

from .document_utils import (
    extract_word_text,
    extract_placeholders,
    generate_certificate
)


BASE_DIR = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)


INPUT_DIR = os.path.join(
    BASE_DIR,
    "input"
)


OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)


EXCEL_PATH = os.path.join(
    INPUT_DIR,
    "section13.xlsx"
)


WORD_PATH = os.path.join(
    INPUT_DIR,
    "certificate.docx"
)


CERTIFICATE_DIR = os.path.join(
    OUTPUT_DIR,
    "certificates"
)


def ensure_directories():

    os.makedirs(
        INPUT_DIR,
        exist_ok=True
    )

    os.makedirs(
        CERTIFICATE_DIR,
        exist_ok=True
    )


def safe_filename(name):

    name = str(name)

    name = re.sub(
        r"[^A-Za-z0-9 _-]",
        "",
        name
    )

    name = name.strip()

    name = name.replace(
        " ",
        "_"
    )

    if not name:

        name = "Participant"

    return name


def inspect_files() -> str:
    """
    Inspect the fixed input files section13.xlsx and certificate.docx.

    This tool reads the Excel column names and sample records,
    and reads the Word template text and placeholders.

    Use this tool FIRST before generating certificates.
    """

    ensure_directories()


    if not os.path.exists(
        EXCEL_PATH
    ):

        return json.dumps({

            "status": "ERROR",

            "message":
                f"Excel file not found: {EXCEL_PATH}"
        })


    if not os.path.exists(
        WORD_PATH
    ):

        return json.dumps({

            "status": "ERROR",

            "message":
                f"Word template not found: {WORD_PATH}"
        })


    # ==========================================
    # Excel
    # ==========================================

    df = pd.read_excel(
        EXCEL_PATH
    )


    df = df.dropna(
        how="all"
    )


    df.columns = [
        str(column).strip()
        for column in df.columns
    ]


    columns = list(
        df.columns
    )


    sample_records = (
        df.head(5)
        .fillna("")
        .to_dict(
            orient="records"
        )
    )


    missing = (
        df.isnull()
        .sum()
        .to_dict()
    )


    duplicates = int(
        df.duplicated().sum()
    )


    # ==========================================
    # Word
    # ==========================================

    word_text = extract_word_text(
        WORD_PATH
    )


    placeholders = extract_placeholders(
        WORD_PATH
    )


    result = {

        "status": "SUCCESS",

        "excel": {

            "file":
                "section13.xlsx",

            "records":
                len(df),

            "columns":
                columns,

            "sample_records":
                sample_records,

            "missing_values":
                missing,

            "duplicates":
                duplicates
        },

        "word": {

            "file":
                "certificate.docx",

            "placeholders":
                placeholders,

            "placeholder_count":
                len(placeholders),

            "text_preview":
                word_text[:3000]
        }
    }


    return json.dumps(
        result,
        indent=2,
        default=str
    )


def generate_certificates(
    mapping_json: str
) -> str:
    """
    Generate personalized certificates using section13.xlsx
    and the user's certificate.docx template.

    mapping_json must be a JSON object where each Word
    placeholder is mapped to an Excel column.

    Example:
    {
        "Recipient_Name": "Name",
        "Recipient_Designation": "Designation",
        "Signatory_Name": "Signatory",
        "Award_Date": "Date",
        "Course_Name": "Course"
    }

    The original certificate.docx is used as the template.
    Its design is not recreated.
    """

    ensure_directories()


    # ==========================================
    # Validate files
    # ==========================================

    if not os.path.exists(
        EXCEL_PATH
    ):

        return json.dumps({

            "status": "ERROR",

            "message":
                "input/section13.xlsx not found."
        })


    if not os.path.exists(
        WORD_PATH
    ):

        return json.dumps({

            "status": "ERROR",

            "message":
                "input/certificate.docx not found."
        })


    # ==========================================
    # Parse mapping
    # ==========================================

    try:

        mapping = json.loads(
            mapping_json
        )

    except Exception as e:

        return json.dumps({

            "status": "ERROR",

            "message":
                f"Invalid mapping JSON: {e}"
        })


    # ==========================================
    # Read Excel
    # ==========================================

    df = pd.read_excel(
        EXCEL_PATH
    )


    df = df.dropna(
        how="all"
    )


    df.columns = [
        str(column).strip()
        for column in df.columns
    ]


    # ==========================================
    # Validate mapping
    # ==========================================

    valid_columns = set(
        df.columns
    )


    invalid_columns = []


    for placeholder, column in mapping.items():

        if (
            column is not None
            and
            column not in valid_columns
        ):

            invalid_columns.append(
                {
                    "placeholder":
                        placeholder,

                    "column":
                        column
                }
            )


    if invalid_columns:

        return json.dumps({

            "status": "ERROR",

            "message":
                "Mapping contains invalid Excel columns.",

            "invalid":
                invalid_columns,

            "valid_columns":
                list(df.columns)
        })


    # ==========================================
    # Clear old certificates
    # ==========================================

    for filename in os.listdir(
        CERTIFICATE_DIR
    ):

        path = os.path.join(
            CERTIFICATE_DIR,
            filename
        )

        if os.path.isfile(path):

            os.remove(path)


    # ==========================================
    # Generate
    # ==========================================

    results = []


    for index, row in df.iterrows():

        record = row.to_dict()


        name_column = mapping.get(
            "Recipient_Name"
        )


        if name_column:

            recipient_name = str(
                record.get(
                    name_column,
                    ""
                )
            )

        else:

            recipient_name = (
                f"Participant_{index + 1}"
            )


        if not recipient_name.strip():

            recipient_name = (
                f"Participant_{index + 1}"
            )


        filename = (
            safe_filename(
                recipient_name
            )
            + ".docx"
        )


        output_path = os.path.join(
            CERTIFICATE_DIR,
            filename
        )


        # Handle duplicate names
        counter = 1


        while os.path.exists(
            output_path
        ):

            filename = (
                safe_filename(
                    recipient_name
                )
                +
                f"_{counter}.docx"
            )

            output_path = os.path.join(
                CERTIFICATE_DIR,
                filename
            )

            counter += 1


        try:

            generate_certificate(
                WORD_PATH,
                output_path,
                record,
                mapping
            )


            results.append({

                "record":
                    index + 1,

                "recipient":
                    recipient_name,

                "file":
                    output_path,

                "status":
                    "SUCCESS"
            })


        except Exception as e:

            results.append({

                "record":
                    index + 1,

                "recipient":
                    recipient_name,

                "file":
                    "",

                "status":
                    "FAILED",

                "error":
                    str(e)
            })


    # ==========================================
    # Report
    # ==========================================

    report_path = os.path.join(
        OUTPUT_DIR,
        "generation_report.csv"
    )


    report = pd.DataFrame(
        results
    )


    report.to_csv(
        report_path,
        index=False
    )


    # ==========================================
    # ZIP
    # ==========================================

    zip_path = os.path.join(
        OUTPUT_DIR,
        "certificates.zip"
    )


    if os.path.exists(
        zip_path
    ):

        os.remove(
            zip_path
        )


    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for filename in os.listdir(
            CERTIFICATE_DIR
        ):

            path = os.path.join(
                CERTIFICATE_DIR,
                filename
            )


            if os.path.isfile(path):

                zip_file.write(
                    path,
                    filename
                )


    success = sum(
        1
        for item in results
        if item["status"] == "SUCCESS"
    )


    failed = sum(
        1
        for item in results
        if item["status"] == "FAILED"
    )


    return json.dumps({

        "status":
            "SUCCESS",

        "total_records":
            len(results),

        "generated":
            success,

        "failed":
            failed,

        "certificate_folder":
            CERTIFICATE_DIR,

        "zip_file":
            zip_path,

        "report_file":
            report_path,

        "results":
            results
    }, indent=2)