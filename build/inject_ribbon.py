#!/usr/bin/env python3
"""Inject ribbon/customUI14.xml into a PowerPoint .ppam add-in package.

Standard library ONLY (zipfile, xml.etree.ElementTree, shutil, os, sys) — the
corporate proxy blocks pip/npm, so this script must never grow a dependency.

A .ppam is an OPC ZIP package. Injection means:
  1. add the part            customUI/customUI14.xml
  2. add a relationship      _rels/.rels  ->  type .../2007/relationships/ui/extensibility
  3. ensure [Content_Types].xml has a Default mapping for the "xml" extension

The script is idempotent: any existing customUI part/relationship is stripped
before re-adding, so it can be run repeatedly against the same file.

Usage:
    python inject_ribbon.py <ribbon_xml> <source_ppam> [<output_ppam>]

With no <output_ppam>, the source file is rewritten in place (via a temp file).
Zip archives cannot be edited in place, so the package is always fully rewritten.
"""

import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

CUSTOMUI_PART = "customUI/customUI14.xml"
RELS_PART = "_rels/.rels"
CONTENT_TYPES_PART = "[Content_Types].xml"

# 2009/07 customUI XML (customUI14) still uses the 2007 *relationship* type.
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CUSTOMUI_NS = "http://schemas.microsoft.com/office/2009/07/customui"


def die(msg):
    sys.stderr.write("inject_ribbon: ERROR: %s\n" % msg)
    sys.exit(1)


def validate_ribbon_xml(path):
    """Parse the ribbon XML and sanity-check the root element + namespace.

    ElementTree only proves well-formedness; schema errors still surface only in
    PowerPoint (enable 'Show add-in user interface errors' to see them)."""
    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        die("ribbon XML is not well-formed: %s" % exc)
    root = tree.getroot()
    expected = "{%s}customUI" % CUSTOMUI_NS
    if root.tag != expected:
        die(
            "ribbon XML root is %s, expected %s "
            "(check the customUI14 namespace)" % (root.tag, expected)
        )
    with open(path, "rb") as fh:
        return fh.read()


def patched_rels(data):
    """Return .rels XML with exactly one customUI relationship."""
    ET.register_namespace("", RELS_NS)
    root = ET.fromstring(data)
    for rel in list(root):
        if rel.get("Type") == REL_TYPE:
            root.remove(rel)
    used_ids = {rel.get("Id") for rel in root}
    rid, n = "rIdSamSamUI", 0
    while rid in used_ids:
        n += 1
        rid = "rIdSamSamUI%d" % n
    ET.SubElement(
        root,
        "{%s}Relationship" % RELS_NS,
        {"Id": rid, "Type": REL_TYPE, "Target": CUSTOMUI_PART},
    )
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def patched_content_types(data):
    """Ensure a Default content type exists for the xml extension."""
    ET.register_namespace("", CT_NS)
    root = ET.fromstring(data)
    for node in root:
        if node.tag == "{%s}Default" % CT_NS and node.get("Extension") == "xml":
            return data  # already fine, keep byte-identical
    ET.SubElement(
        root,
        "{%s}Default" % CT_NS,
        {"Extension": "xml", "ContentType": "application/xml"},
    )
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def inject(ribbon_path, src_ppam, out_ppam):
    ribbon_bytes = validate_ribbon_xml(ribbon_path)

    if not zipfile.is_zipfile(src_ppam):
        die("%s is not a ZIP package — is it really a .ppam?" % src_ppam)

    in_place = os.path.abspath(src_ppam) == os.path.abspath(out_ppam)
    tmp_fd, tmp_path = tempfile.mkstemp(
        suffix=".ppam", dir=os.path.dirname(os.path.abspath(out_ppam))
    )
    os.close(tmp_fd)

    try:
        with zipfile.ZipFile(src_ppam, "r") as zin, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            names = zin.namelist()
            if RELS_PART not in names:
                die("package has no %s — corrupt .ppam?" % RELS_PART)
            if CONTENT_TYPES_PART not in names:
                die("package has no %s — corrupt .ppam?" % CONTENT_TYPES_PART)

            for item in zin.infolist():
                name = item.filename
                # strip any previous customUI part (and its own rels, if any)
                if name == CUSTOMUI_PART or name.startswith("customUI/"):
                    continue
                data = zin.read(name)
                if name == RELS_PART:
                    data = patched_rels(data)
                elif name == CONTENT_TYPES_PART:
                    data = patched_content_types(data)
                zout.writestr(name, data)

            zout.writestr(CUSTOMUI_PART, ribbon_bytes)

        if in_place or os.path.exists(out_ppam):
            os.replace(tmp_path, out_ppam)
        else:
            shutil.move(tmp_path, out_ppam)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    print("inject_ribbon: OK — wrote %s" % out_ppam)
    print("inject_ribbon: REMINDER — close and reopen ALL PowerPoint windows;")
    print("               PowerPoint caches the ribbon until a full restart.")


def main(argv):
    if len(argv) not in (3, 4):
        sys.stderr.write(__doc__)
        sys.exit(2)
    ribbon_path, src_ppam = argv[1], argv[2]
    out_ppam = argv[3] if len(argv) == 4 else src_ppam
    for path in (ribbon_path, src_ppam):
        if not os.path.isfile(path):
            die("file not found: %s" % path)
    inject(ribbon_path, src_ppam, out_ppam)


if __name__ == "__main__":
    main(sys.argv)
