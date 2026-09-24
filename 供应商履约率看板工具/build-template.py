from __future__ import annotations

import base64
import re
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
SOURCE = (
    ROOT
    / "outputs"
    / "europe_supplier_performance_q2_2026_final"
    / "欧洲供应商履约率报表_2026年4-6月_最终严谨版.xlsx"
)
TARGET = Path(__file__).resolve().parent / "vendor" / "performance-report-template.xlsx"
SCRIPT_TARGET = (
    Path(__file__).resolve().parent / "vendor" / "performance-report-template.js"
)
MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CORE_PROPERTIES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Supplier fulfillment report</dc:title><dc:creator>OTWS Supplier Portal</dc:creator></cp:coreProperties>"""
APP_PROPERTIES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>OTWS Supplier Fulfillment Dashboard</Application></Properties>"""


def normalize_main_namespace(data: bytes) -> bytes:
    text = data.decode("utf-8")
    if f'xmlns:x="{MAIN_NAMESPACE}"' not in text:
        return data
    text = text.replace(f'xmlns:x="{MAIN_NAMESPACE}"', f'xmlns="{MAIN_NAMESPACE}"')
    text = re.sub(r"<(/?)x:", r"<\1", text)
    text = re.sub(r"(\s)x:([A-Za-z_][\w.-]*)=", r"\1\2=", text)
    return text.encode("utf-8")


def main() -> None:
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SOURCE, "r") as source_zip:
        with zipfile.ZipFile(
            TARGET,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
        ) as target_zip:
            for item in source_zip.infolist():
                data = source_zip.read(item.filename)
                if item.filename.endswith(".xml"):
                    data = normalize_main_namespace(data)
                target_zip.writestr(item, data)
            target_zip.writestr("docProps/core.xml", CORE_PROPERTIES)
            target_zip.writestr("docProps/app.xml", APP_PROPERTIES)
    encoded = base64.b64encode(TARGET.read_bytes()).decode("ascii")
    SCRIPT_TARGET.write_text(
        "(function(root){\"use strict\";root.SupplierFulfillmentTemplateBase64=\""
        + encoded
        + "\";})(typeof globalThis!==\"undefined\"?globalThis:this);\n",
        encoding="ascii",
    )
    print(TARGET)
    print(SCRIPT_TARGET)


if __name__ == "__main__":
    main()
