import zipfile
from lxml import etree
from .config import TEMPLATE_FIRMAS, STUDENTS


def get_student_signature(student_name: str) -> bytes:
    name_lower = student_name.lower()
    if "pablo" in name_lower:
        return extract_signature("rId8")
    elif "angel" in name_lower or "rodrigo" in name_lower:
        return extract_signature("rId28")
    return None


def extract_signature(rel_id: str) -> bytes:
    with zipfile.ZipFile(TEMPLATE_FIRMAS) as z:
        rels_xml = z.read("word/_rels/document.xml.rels")
        rels_root = etree.fromstring(rels_xml)
        for child in rels_root:
            if child.get("Id") == rel_id:
                target = child.get("Target")
                return z.read(f"word/{target}")
    return None
