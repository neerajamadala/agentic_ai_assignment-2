import os
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree


WORD_NS = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

NS = {
    "w": WORD_NS
}


def extract_word_text(docx_path):
    """
    Extract readable text from a DOCX file.
    """

    text_parts = []

    with ZipFile(docx_path, "r") as zip_file:

        for filename in zip_file.namelist():

            if (
                filename.startswith("word/")
                and filename.endswith(".xml")
            ):

                try:

                    data = zip_file.read(
                        filename
                    )

                    root = etree.fromstring(
                        data
                    )

                    nodes = root.xpath(
                        ".//w:t",
                        namespaces=NS
                    )

                    for node in nodes:

                        if node.text:

                            text_parts.append(
                                node.text
                            )

                except Exception:
                    continue

    return " ".join(text_parts)


def extract_placeholders(docx_path):
    """
    Find {{placeholder}} values inside a Word document.
    """

    text = extract_word_text(
        docx_path
    )

    placeholders = re.findall(
        r"\{\{(.*?)\}\}",
        text
    )

    placeholders = [
        item.strip()
        for item in placeholders
    ]

    return list(
        dict.fromkeys(
            placeholders
        )
    )


def replace_in_xml(
    xml_data,
    replacements
):

    try:

        root = etree.fromstring(
            xml_data
        )

    except Exception:

        return xml_data


    paragraphs = root.xpath(
        ".//w:p",
        namespaces=NS
    )


    for paragraph in paragraphs:

        text_nodes = paragraph.xpath(
            ".//w:t",
            namespaces=NS
        )

        if not text_nodes:

            continue


        # Combine all Word runs
        full_text = "".join(
            node.text or ""
            for node in text_nodes
        )


        changed = False


        for placeholder, value in replacements.items():

            if placeholder in full_text:

                full_text = full_text.replace(
                    placeholder,
                    str(value)
                )

                changed = True


        if changed:

            # Put final text into first run
            text_nodes[0].text = full_text

            # Clear other runs
            for node in text_nodes[1:]:

                node.text = ""


    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True
    )


def replace_text_in_docx(
    template_path,
    output_path,
    replacements
):

    temporary_file = (
        output_path + ".tmp"
    )


    with ZipFile(
        template_path,
        "r"
    ) as source_zip:

        with ZipFile(
            temporary_file,
            "w",
            ZIP_DEFLATED
        ) as destination_zip:

            for item in source_zip.infolist():

                data = source_zip.read(
                    item.filename
                )


                if (
                    item.filename.startswith("word/")
                    and
                    item.filename.endswith(".xml")
                ):

                    data = replace_in_xml(
                        data,
                        replacements
                    )


                destination_zip.writestr(
                    item,
                    data
                )


    if os.path.exists(
        output_path
    ):

        os.remove(
            output_path
        )


    os.rename(
        temporary_file,
        output_path
    )


def generate_certificate(
    template_path,
    output_path,
    record,
    mapping
):

    replacements = {}


    for placeholder, column in mapping.items():

        if column is None:

            value = ""

        else:

            value = record.get(
                column,
                ""
            )


        if value is None:

            value = ""


        replacements[
            "{{" + placeholder + "}}"
        ] = str(value)


    replace_text_in_docx(
        template_path,
        output_path,
        replacements
    )