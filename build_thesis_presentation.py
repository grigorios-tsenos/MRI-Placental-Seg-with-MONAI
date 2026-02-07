#!/usr/bin/env python3
"""Build a styled 20-slide LibreOffice presentation (.odp) for the thesis."""

from __future__ import annotations

import copy
import mimetypes
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "presentation": "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def find_frame(page: ET.Element, pclass: str) -> ET.Element | None:
    for frame in page.findall(qn("draw", "frame")):
        if frame.get(qn("presentation", "class")) == pclass:
            return frame
    return None


def clear_textbox(frame: ET.Element) -> ET.Element:
    textbox = frame.find(qn("draw", "text-box"))
    if textbox is None:
        textbox = ET.SubElement(frame, qn("draw", "text-box"))
    for child in list(textbox):
        textbox.remove(child)
    return textbox


def set_paragraphs(frame: ET.Element, lines: list[str]) -> None:
    textbox = clear_textbox(frame)
    for line in lines:
        p = ET.SubElement(textbox, qn("text", "p"))
        p.text = line


def set_title(page: ET.Element, title: str) -> None:
    frame = find_frame(page, "title")
    if frame is None:
        raise RuntimeError("Title frame not found")
    set_paragraphs(frame, [title])


def set_outline(page: ET.Element, lines: list[str]) -> None:
    frame = find_frame(page, "outline")
    if frame is None:
        return
    bullet_lines = [f"• {line}" for line in lines]
    set_paragraphs(frame, bullet_lines)


def set_subtitle(page: ET.Element, lines: list[str]) -> None:
    frame = find_frame(page, "subtitle")
    if frame is None:
        return
    set_paragraphs(frame, lines)


def update_notes_page_number(page: ET.Element, page_no: int) -> None:
    notes = page.find(qn("presentation", "notes"))
    if notes is None:
        return
    thumb = notes.find(qn("draw", "page-thumbnail"))
    if thumb is None:
        return
    thumb.set(qn("draw", "page-number"), str(page_no))


def remove_outline_frame(page: ET.Element) -> None:
    frame = find_frame(page, "outline")
    if frame is not None:
        page.remove(frame)


def add_text_frame(
    page: ET.Element,
    style_name: str,
    pclass: str,
    x: str,
    y: str,
    w: str,
    h: str,
    lines: list[str],
    bullet: bool = False,
) -> None:
    frame = ET.SubElement(
        page,
        qn("draw", "frame"),
        {
            qn("presentation", "style-name"): style_name,
            qn("draw", "layer"): "layout",
            qn("svg", "x"): x,
            qn("svg", "y"): y,
            qn("svg", "width"): w,
            qn("svg", "height"): h,
            qn("presentation", "class"): pclass,
            qn("presentation", "user-transformed"): "true",
        },
    )
    if bullet:
        lines = [f"• {line}" for line in lines]
    set_paragraphs(frame, lines)


def add_image_frame(
    page: ET.Element,
    href: str,
    x: str,
    y: str,
    w: str,
    h: str,
    name: str,
) -> None:
    frame = ET.SubElement(
        page,
        qn("draw", "frame"),
        {
            qn("draw", "style-name"): "gr1",
            qn("draw", "name"): name,
            qn("draw", "layer"): "layout",
            qn("svg", "x"): x,
            qn("svg", "y"): y,
            qn("svg", "width"): w,
            qn("svg", "height"): h,
        },
    )
    ET.SubElement(
        frame,
        qn("draw", "image"),
        {
            qn("xlink", "href"): href,
            qn("xlink", "type"): "simple",
            qn("xlink", "show"): "embed",
            qn("xlink", "actuate"): "onLoad",
        },
    )


def slugify(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    return name.strip("_") or "image"


def build_presentation() -> None:
    root_dir = Path(__file__).resolve().parent
    template = Path("/Applications/LibreOffice.app/Contents/Resources/template/common/presnt/Midnightblue.otp")
    output = root_dir / "Thesis Doc" / "thesis_presentation_styled.odp"

    slides = [
        {
            "layout": "title",
            "title": "Τμηματοποίηση Πλακούντα σε MRI",
            "subtitle": [
                "με τεχνικές Βαθιάς Μάθησης",
                "Διπλωματική Εργασία - Γρηγόριος Τσένος",
                "Επιβλέπων: Γεώργιος Ματσόπουλος | Φεβρουάριος 2026",
            ],
        },
        {
            "layout": "text",
            "title": "Δομή Παρουσίασης",
            "bullets": [
                "Κίνητρο και πρόβλημα τμηματοποίησης πλακούντα σε 3D MRI",
                "Δεδομένα, προεπεξεργασία και κοινό πειραματικό πρωτόκολλο",
                "Αρχιτεκτονικές: CNN, Transformer, SSM/Mamba",
                "Ποσοτικά και ποιοτικά αποτελέσματα",
                "Συμπεράσματα, περιορισμοί και μελλοντικές κατευθύνσεις",
            ],
        },
        {
            "layout": "mix",
            "title": "Κλινικό Υπόβαθρο και Κίνητρο",
            "bullets": [
                "Ο πλακούντας είναι κρίσιμο όργανο για την έκβαση της κύησης",
                "Η MRI επιτρέπει ποσοτική ανάλυση όγκου και μορφολογίας",
                "Η χειροκίνητη τμηματοποίηση είναι χρονοβόρα και μεταβλητή",
                "Στόχος: αυτοματοποιημένη, συνεπής και αναπαραγώγιμη τμηματοποίηση",
            ],
            "image": "Thesis Doc/figures/Stages-Placenta.png",
        },
        {
            "layout": "text",
            "title": "Πρόβλημα και Στόχοι Εργασίας",
            "bullets": [
                "Δυαδική τμηματοποίηση πλακούντα σε ογκομετρικά MRI",
                "Σύγκριση 7+ αρχιτεκτονικών σε ίδιες συνθήκες εκπαίδευσης",
                "Αξιολόγηση με Dice, IoU και validation loss",
                "Ερμηνεία επιδόσεων με καμπύλες σύγκλισης και οπτική επιθεώρηση",
            ],
        },
        {
            "layout": "text",
            "title": "Σύνολο Δεδομένων",
            "bullets": [
                "N = 137 περιπτώσεις, με MRI όγκο + δυαδική μάσκα (.nii.gz)",
                "Διαχωρισμός 80/20 με σταθερό seed=121",
                "109 περιπτώσεις train και 28 validation",
                "Βασική μονάδα αξιολόγησης: ολόκληρος 3D όγκος (case-level)",
            ],
        },
        {
            "layout": "image",
            "title": "Παράδειγμα 3D MRI και Μάσκας Πλακούντα",
            "image": "Thesis Doc/figures/First Slice along segmentation and 3D view.png",
            "caption": "MRI τομή, αντίστοιχη μάσκα και 3D απεικόνιση περιοχής ενδιαφέροντος",
        },
        {
            "layout": "text",
            "title": "Προεπεξεργασία και Εκπαίδευση",
            "bullets": [
                "Orientation/Spacing εναρμόνιση και intensity normalization",
                "Foreground cropping και padding σε roi_size=(96,96,64)",
                "Patch-based εκπαίδευση με ελεγχόμενο pos/neg sampling",
                "Στοχαστικές επαυξήσεις: flips, affine, gaussian noise/smoothing",
                "120 epochs με κοινή ροή για όλες τις αρχιτεκτονικές",
            ],
        },
        {
            "layout": "text",
            "title": "Αρχιτεκτονικές που Συγκρίθηκαν",
            "bullets": [
                "CNN-based: U-Net, Attention U-Net, DynUNet, SegResNet",
                "Transformer-based: UNETR, SwinUNETR",
                "SSM/Mamba-based: SegMamba",
                "Όλα τα μοντέλα σε ενιαίο pipeline για δίκαιη σύγκριση",
            ],
        },
        {
            "layout": "image",
            "title": "U-Net και 3D Encoder-Decoder Λογική",
            "image": "Thesis Doc/figures/u-net-architecture.png",
            "caption": "Skip connections για μεταφορά λεπτομέρειας και ακριβέστερα όρια",
        },
        {
            "layout": "image",
            "title": "UNETR vs SwinUNETR",
            "image": "Thesis Doc/figures/swin vs vit.png",
            "caption": "Ιεραρχικός Swin encoder με local windows έναντι ViT-style encoder",
        },
        {
            "layout": "text",
            "title": "SSM/Mamba και SegMamba",
            "bullets": [
                "Selective State Space Modeling για μακρινές εξαρτήσεις",
                "Γραμμική κλιμάκωση ως προς το μήκος ακολουθίας",
                "Στόχος: καλύτερη αποδοτικότητα από πλήρες self-attention σε 3D",
                "Στο πείραμα: κορυφαία ισορροπία ποιότητας και σταθερότητας",
            ],
        },
        {
            "layout": "text",
            "title": "Μετρικές Αξιολόγησης",
            "bullets": [
                "Dice (DSC): βασική μετρική επικάλυψης πρόβλεψης/ground truth",
                "IoU (Jaccard): συμπληρωματική μετρική επικάλυψης",
                "Validation Loss: δείκτης συνολικής βελτιστοποίησης",
                "Συνδυαστική ανάγνωση: αριθμητικές τιμές + ποιοτική επιθεώρηση",
            ],
        },
        {
            "layout": "text",
            "title": "Κύρια Ποσοτικά Αποτελέσματα (Validation)",
            "bullets": [
                "SegMamba 1: Dice=0.8606 | IoU=0.7566 | ValLoss=0.1685",
                "SegResNet 1: Dice=0.8601 | IoU=0.7558 | ValLoss=0.1678",
                "SwinUNETR: Dice=0.8490 | IoU=0.7401 | ValLoss=0.1838",
                "UNETR: Dice=0.7720 | IoU=0.6345 | ValLoss=0.2842",
                "Διαφορές κορυφής μικρές (~0.002-0.003 Dice)",
            ],
        },
        {
            "layout": "text",
            "title": "Συνοπτική Κατάταξη και Ευρήματα",
            "bullets": [
                "Ομάδα κορυφής γύρω από Dice≈0.86: SegMamba + SegResNet",
                "SwinUNETR: καλύτερο Transformer, αλλά ~0.01 πίσω από κορυφή",
                "UNETR: σαφές outlier στο συγκεκριμένο split",
                "Οι βαριές παραλλαγές κερδίζουν λίγο αλλά μετρήσιμα",
            ],
        },
        {
            "layout": "split",
            "title": "Καμπύλες Εκπαίδευσης (CNN Μοντέλα)",
            "image_left": "Thesis Doc/figures/history_attentionUnet.png",
            "image_right": "Thesis Doc/figures/history_segresHeavy.png",
            "caption_left": "Attention U-Net",
            "caption_right": "SegResNet 1",
        },
        {
            "layout": "split",
            "title": "Καμπύλες Εκπαίδευσης (UNETR vs SegMamba)",
            "image_left": "Thesis Doc/figures/history_UNETR.png",
            "image_right": "Thesis Doc/figures/history_segmambaheavy.png",
            "caption_left": "UNETR",
            "caption_right": "SegMamba 1",
        },
        {
            "layout": "image",
            "title": "Ποιοτική Αξιολόγηση: SegResNet",
            "image": "Thesis Doc/figures/segmentation_masks_SegResHeavy.png",
            "caption": "Συνεκτικές μάσκες με καλή μορφολογική συμφωνία και περιορισμένα false positives",
        },
        {
            "layout": "split",
            "title": "Ποιοτική Αξιολόγηση: Outlier vs Κορυφή",
            "image_left": "Thesis Doc/figures/segmentation_masks_UNETR.png",
            "image_right": "Thesis Doc/figures/segmentation_masks_segmambaHEAVY.png",
            "caption_left": "UNETR",
            "caption_right": "SegMamba 1",
        },
        {
            "layout": "text",
            "title": "Συζήτηση, Περιορισμοί, Επεκτάσεις",
            "bullets": [
                "Η κατάταξη είναι ισχυρή συγκριτικά, αλλά βασίζεται σε ένα validation split",
                "Οι οπτικές συγκρίσεις είναι 2D τομές ενώ η αξιολόγηση είναι 3D",
                "Μελλοντικά: cross-validation, ανεξάρτητο test set, boundary metrics",
                "Ενσωμάτωση πλήρους end-to-end pipeline για αυτόματο ROI localization",
            ],
        },
        {
            "layout": "closing",
            "title": "Συμπεράσματα",
            "bullets": [
                "SegMamba και SegResNet έδωσαν την καλύτερη συνολική επίδοση",
                "Το SwinUNETR ήταν ανταγωνιστικό και σαφώς ισχυρότερο από UNETR",
                "Η επιλογή μοντέλου στην πράξη εξαρτάται και από VRAM/χρόνο inference",
                "Ευχαριστώ για την προσοχή σας",
            ],
        },
    ]

    assert len(slides) == 20, "Expected exactly 20 slides"

    with zipfile.ZipFile(template, "r") as zin:
        package = {name: zin.read(name) for name in zin.namelist()}

    content_root = ET.fromstring(package["content.xml"])
    manifest_root = ET.fromstring(package["META-INF/manifest.xml"])

    presentation = content_root.find(qn("office", "body")).find(qn("office", "presentation"))
    all_pages = presentation.findall(qn("draw", "page"))
    if len(all_pages) < 3:
        raise RuntimeError("Template does not contain expected page prototypes")

    proto_title = all_pages[0]
    proto_content = all_pages[1]
    proto_closing = all_pages[2]

    settings = presentation.find(qn("presentation", "settings"))

    for page in list(presentation.findall(qn("draw", "page"))):
        presentation.remove(page)

    image_map: dict[str, str] = {}
    image_counter = 1

    def register_image(rel_path: str) -> str:
        nonlocal image_counter
        src = (root_dir / rel_path).resolve()
        if not src.exists():
            raise FileNotFoundError(f"Image not found: {src}")
        if str(src) in image_map:
            return image_map[str(src)]

        ext = src.suffix.lower()
        safe_name = slugify(src.stem)
        target = f"Pictures/{image_counter:02d}_{safe_name}{ext}"
        image_counter += 1

        package[target] = src.read_bytes()
        image_map[str(src)] = target

        media_type = mimetypes.types_map.get(ext, "application/octet-stream")
        entry = ET.SubElement(
            manifest_root,
            qn("manifest", "file-entry"),
            {
                qn("manifest", "full-path"): target,
                qn("manifest", "media-type"): media_type,
            },
        )
        # Keep linter quiet about unused variable while still creating the node.
        _ = entry
        return target

    def make_page(i: int, spec: dict[str, object]) -> ET.Element:
        layout = spec["layout"]
        if layout == "title":
            page = copy.deepcopy(proto_title)
            page.set(qn("draw", "style-name"), "dp1")
        elif layout == "closing":
            page = copy.deepcopy(proto_closing)
            page.set(qn("draw", "style-name"), "dp3")
            page.set(qn("draw", "master-page-name"), "Midnightblue2")
        else:
            page = copy.deepcopy(proto_content)
            page.set(qn("draw", "style-name"), "dp3")
            page.set(qn("draw", "master-page-name"), "Midnightblue")

        page.set(qn("draw", "name"), f"page{i}")
        update_notes_page_number(page, i)

        title = str(spec.get("title", ""))
        set_title(page, title)

        if layout == "title":
            set_subtitle(page, [str(x) for x in spec.get("subtitle", [])])
        elif layout == "text":
            set_outline(page, [str(x) for x in spec.get("bullets", [])])
        elif layout == "closing":
            add_text_frame(
                page,
                "Midnightblue2-outline1",
                "outline",
                "8.8cm",
                "9.2cm",
                "17.6cm",
                "4.4cm",
                [str(x) for x in spec.get("bullets", [])],
                bullet=True,
            )
        elif layout == "mix":
            outline = find_frame(page, "outline")
            if outline is None:
                raise RuntimeError("Outline frame not found in mix slide")
            outline.set(qn("svg", "x"), "1cm")
            outline.set(qn("svg", "y"), "3.3cm")
            outline.set(qn("svg", "width"), "11.8cm")
            outline.set(qn("svg", "height"), "10.8cm")
            set_paragraphs(outline, [f"• {x}" for x in spec.get("bullets", [])])

            img_rel = str(spec["image"])
            href = register_image(img_rel)
            add_image_frame(page, href, "13.2cm", "3.3cm", "13.2cm", "10.8cm", f"img{i}")
        elif layout == "image":
            remove_outline_frame(page)
            href = register_image(str(spec["image"]))
            add_image_frame(page, href, "1.0cm", "3.0cm", "26.0cm", "10.9cm", f"img{i}")
            add_text_frame(
                page,
                "Midnightblue-outline1",
                "outline",
                "1.2cm",
                "14.2cm",
                "25.6cm",
                "0.7cm",
                [str(spec.get("caption", ""))],
                bullet=False,
            )
        elif layout == "split":
            remove_outline_frame(page)
            href_l = register_image(str(spec["image_left"]))
            href_r = register_image(str(spec["image_right"]))
            add_image_frame(page, href_l, "1.0cm", "3.2cm", "12.6cm", "9.9cm", f"img{i}_l")
            add_image_frame(page, href_r, "14.4cm", "3.2cm", "12.6cm", "9.9cm", f"img{i}_r")
            add_text_frame(
                page,
                "Midnightblue-outline1",
                "outline",
                "1.1cm",
                "13.4cm",
                "12.4cm",
                "0.8cm",
                [str(spec.get("caption_left", ""))],
                bullet=False,
            )
            add_text_frame(
                page,
                "Midnightblue-outline1",
                "outline",
                "14.5cm",
                "13.4cm",
                "12.4cm",
                "0.8cm",
                [str(spec.get("caption_right", ""))],
                bullet=False,
            )
        else:
            raise RuntimeError(f"Unsupported layout: {layout}")

        return page

    for idx, spec in enumerate(slides, start=1):
        presentation.append(make_page(idx, spec))

    if settings is not None:
        presentation.append(settings)

    package["content.xml"] = ET.tostring(content_root, encoding="utf-8", xml_declaration=False)

    # Update media types for ODP (not template)
    package["mimetype"] = b"application/vnd.oasis.opendocument.presentation"

    for entry in manifest_root.findall(qn("manifest", "file-entry")):
        if entry.get(qn("manifest", "full-path")) == "/":
            entry.set(qn("manifest", "media-type"), "application/vnd.oasis.opendocument.presentation")
            entry.set(qn("manifest", "version"), "1.2")

    package["META-INF/manifest.xml"] = ET.tostring(manifest_root, encoding="utf-8", xml_declaration=False)

    with zipfile.ZipFile(output, "w") as zout:
        # ODF requires mimetype first and uncompressed.
        zout.writestr(
            zipfile.ZipInfo("mimetype"),
            package["mimetype"],
            compress_type=zipfile.ZIP_STORED,
        )

        for name in sorted(package.keys()):
            if name == "mimetype":
                continue
            zout.writestr(name, package[name], compress_type=zipfile.ZIP_DEFLATED)

    print(f"Wrote: {output}")


if __name__ == "__main__":
    build_presentation()
