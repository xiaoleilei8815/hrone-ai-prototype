#!/usr/bin/env python3
"""Create an XLSX test fixture with formula caches removed and full recalculation forced."""

import re
import sys
import zipfile


FORMULA_CACHE = re.compile(rb"(<f(?:\s[^>]*)?>.*?</f>)<v>.*?</v>", re.DOTALL)
CALC_PR = re.compile(rb"<calcPr\b[^>]*/>")
CALC_RELATION = re.compile(
    rb"<Relationship\b[^>]*Type=\"[^\"]*/calcChain\"[^>]*/>"
)
CALC_CONTENT = re.compile(rb"<Override\b[^>]*PartName=\"/xl/calcChain.xml\"[^>]*/>")
FORCED_CALC = (
    b'<calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
)


def main():
    if len(sys.argv) != 3:
        raise SystemExit("usage: force-formula-recalc-fixture.py <source.xlsx> <output.xlsx>")

    source, output = sys.argv[1:]
    removed = 0
    with zipfile.ZipFile(source, "r") as input_zip, zipfile.ZipFile(
        output, "w", zipfile.ZIP_DEFLATED
    ) as output_zip:
        for item in input_zip.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            data = input_zip.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                data, count = FORMULA_CACHE.subn(rb"\1", data)
                removed += count
            elif item.filename == "xl/workbook.xml":
                if CALC_PR.search(data):
                    data = CALC_PR.sub(FORCED_CALC, data, count=1)
                else:
                    data = data.replace(b"</workbook>", FORCED_CALC + b"</workbook>")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = CALC_RELATION.sub(b"", data)
            elif item.filename == "[Content_Types].xml":
                data = CALC_CONTENT.sub(b"", data)
            output_zip.writestr(item, data)

    if removed == 0:
        raise RuntimeError("未找到任何公式缓存")
    print(f"removed_formula_caches={removed}")


if __name__ == "__main__":
    main()
