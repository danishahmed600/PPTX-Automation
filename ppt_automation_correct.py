import time
from pathlib import Path
from collections import defaultdict

import pythoncom
import win32com.client


# ============================================================
# PPTX TEMPLATE / MASTER AUTOMATION
#
# Windows only
# Requires:
#   pip install pywin32
#
# RULES:
#   1. Only slides explicitly mapped by the user are redesigned.
#   2. Mapped slides use the REAL template CustomLayout/Master.
#   3. Source slide master/theme/background is NOT copied to mapped slides.
#   4. Only source slide content shapes are copied to mapped slides.
#   5. Unmapped slides are copied as complete slides and remain AS-IS.
#   6. Original input/template files are never overwritten.
# ============================================================


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
MSO_PLACEHOLDER = 14

# Temporary layout preview images.
# Created after the template is selected and deleted automatically
# when the program finishes or is cancelled.
TEMP_PREVIEW_DIR = BASE_DIR / "temp_layout_previews"


# ============================================================
# TEMPORARY TEMPLATE LAYOUT PREVIEWS
# ============================================================

def safe_filename(value):
    """Convert a PowerPoint master/layout name into a Windows-safe filename."""
    import re

    value = str(value or "").strip()
    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", " ", value).strip()

    return value[:100] or "Unnamed"


def delete_preview_folder():
    """Delete the temporary layout-preview folder and all PNG files."""
    import shutil

    try:
        if TEMP_PREVIEW_DIR.exists():
            shutil.rmtree(
                TEMP_PREVIEW_DIR,
                ignore_errors=True
            )
    except Exception:
        pass


def add_preview_sample_content(slide):
    """
    Add temporary sample text to a preview slide.

    These objects are ONLY used to make the layout visually
    understandable in the PNG preview. They never go into the
    final presentation.
    """

    try:
        slide_width = float(
            slide.Parent.PageSetup.SlideWidth
        )

        slide_height = float(
            slide.Parent.PageSetup.SlideHeight
        )

        # Temporary title.
        title = slide.Shapes.AddTextbox(
            1,
            slide_width * 0.06,
            slide_height * 0.07,
            slide_width * 0.88,
            slide_height * 0.14
        )

        title.TextFrame.TextRange.Text = "AI Patshala"
        title.TextFrame.TextRange.Font.Size = 28
        title.TextFrame.TextRange.Font.Bold = True

        # Temporary body.
        body = slide.Shapes.AddTextbox(
            1,
            slide_width * 0.06,
            slide_height * 0.28,
            slide_width * 0.75,
            slide_height * 0.30
        )

        body.TextFrame.TextRange.Text = (
            "Template Layout Preview\n"
            "Sample title, text and visual content"
        )

        body.TextFrame.TextRange.Font.Size = 18

    except Exception:
        # The layout can still be exported even if a specialized
        # layout does not accept the sample text boxes.
        pass


def create_layout_preview_images(
    powerpoint,
    template_path,
    layouts
):
    """
    Create PNG previews for every CustomLayout.

    IMPORTANT:
    We DO NOT add preview slides to the user's opened template.

    Instead:
        1. Create temp_layout_previews/
        2. Make a temporary copy of the selected template.
        3. Open that temporary copy as writable.
        4. Create one temporary slide per CustomLayout.
        5. Export each slide to PNG.
        6. Delete the temporary presentation.
        7. Keep the PNG folder alive until final PPT creation.
        8. Delete the PNG folder automatically after creation.

    This avoids PowerPoint's read-only-template problem.
    """

    import shutil
    import tempfile

    # --------------------------------------------------------
    # ALWAYS CREATE THE FOLDER FIRST
    # --------------------------------------------------------

    delete_preview_folder()

    TEMP_PREVIEW_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    print("\n" + "=" * 70)
    print("CREATING TEMPORARY TEMPLATE LAYOUT PREVIEWS")
    print("=" * 70)

    print(
        f"Preview folder:"
    )
    print(
        TEMP_PREVIEW_DIR.resolve()
    )

    temp_root = Path(
        tempfile.mkdtemp(
            prefix="ppt_layout_preview_"
        )
    )

    temp_template_path = (
        temp_root / "template_preview_source.pptx"
    )

    preview_presentation = None

    created = 0

    try:
        # ----------------------------------------------------
        # COPY ORIGINAL TEMPLATE
        # ----------------------------------------------------

        shutil.copy2(
            str(template_path),
            str(temp_template_path)
        )

        print(
            "Created temporary template copy."
        )

        # ----------------------------------------------------
        # OPEN TEMP COPY AS WRITABLE
        # ----------------------------------------------------

        preview_presentation = (
            powerpoint.Presentations.Open(
                str(temp_template_path),
                ReadOnly=False,
                Untitled=False,
                WithWindow=False
            )
        )

        print(
            "Temporary template opened successfully."
        )

        # ----------------------------------------------------
        # CREATE PREVIEW FOR EVERY LAYOUT
        # ----------------------------------------------------

        for layout in layouts:

            preview_slide = None

            try:
                custom_layout = get_actual_custom_layout(
                    preview_presentation,
                    layout
                )

                preview_slide = (
                    preview_presentation.Slides.AddSlide(
                        preview_presentation.Slides.Count + 1,
                        custom_layout
                    )
                )

                add_preview_sample_content(
                    preview_slide
                )

                filename = (
                    f'{layout["id"]:03d}__'
                    f'{safe_filename(layout["master_name"])}__'
                    f'{safe_filename(layout["layout_name"])}.png'
                )

                preview_path = (
                    TEMP_PREVIEW_DIR / filename
                )

                preview_slide.Export(
                    str(preview_path),
                    "PNG",
                    1600,
                    900
                )

                if preview_path.exists():
                    created += 1

                    print(
                        f'[{layout["id"]}] '
                        f'{layout["layout_name"]} -> '
                        f'{preview_path.name}'
                    )
                else:
                    print(
                        f'[{layout["id"]}] '
                        f'{layout["layout_name"]} -> '
                        f'WARNING: PNG not created'
                    )

            except Exception as exc:

                print(
                    f'[{layout["id"]}] '
                    f'{layout["layout_name"]} -> '
                    f'WARNING: preview failed: {exc}'
                )

            finally:

                if preview_slide is not None:
                    try:
                        preview_slide.Delete()
                    except Exception:
                        pass

        # ----------------------------------------------------
        # CLOSE TEMP PRESENTATION
        # ----------------------------------------------------

        try:
            preview_presentation.Close()
        except Exception:
            pass

        preview_presentation = None

        print(
            "\nTemporary preview presentation deleted."
        )

    except Exception as exc:

        print(
            f"\nPreview generation error: {exc}"
        )

        if preview_presentation is not None:
            try:
                preview_presentation.Close()
            except Exception:
                pass

            preview_presentation = None

    finally:

        # ----------------------------------------------------
        # DELETE TEMPORARY PPTX COPY
        # ----------------------------------------------------

        try:
            shutil.rmtree(
                temp_root,
                ignore_errors=True
            )
        except Exception:
            pass

    print("-" * 70)
    print(
        f"Preview images created: {created}/{len(layouts)}"
    )

    if created:
        print(
            "SUCCESS: Layout preview folder is available "
            "until PPT creation finishes."
        )
    else:
        print(
            "WARNING: No layout preview images were created."
        )

    print("=" * 70)

    return created


# ============================================================
# FILE HELPERS
# ============================================================

def ensure_directories():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def select_file(folder, title):
    files = sorted(
        [
            path
            for path in folder.iterdir()
            if path.is_file() and path.suffix.lower() == ".pptx"
        ],
        key=lambda path: path.name.lower()
    )

    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)

    if not files:
        print("No .pptx files found in:")
        print(folder)
        return None

    for number, path in enumerate(files, 1):
        print(f"[{number}] {path.name}")

    print("=" * 70)

    while True:
        value = input("Select file number or filename: ").strip()

        if not value:
            continue

        if value.isdigit():
            index = int(value) - 1
            if 0 <= index < len(files):
                return files[index]

        for path in files:
            if path.name.lower() == value.lower():
                return path

        print("Invalid selection. Try again.")


# ============================================================
# TEXT / SHAPE HELPERS
# ============================================================

def safe_shape_text(shape):
    try:
        if shape.HasTextFrame:
            return str(shape.TextFrame.TextRange.Text or "").strip()
    except Exception:
        pass

    return ""


def normalize_text(text):
    if not text:
        return ""

    return " ".join(
        str(text)
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\x0b", " ")
        .split()
    ).strip().lower()


def get_slide_title(slide):
    try:
        if slide.Shapes.HasTitle:
            text = safe_shape_text(slide.Shapes.Title)
            if text:
                return text.replace("\r", " ").replace("\n", " ").strip()
    except Exception:
        pass

    try:
        for index in range(1, slide.Shapes.Count + 1):
            text = safe_shape_text(slide.Shapes(index))
            if text:
                return text.replace("\r", " ").replace("\n", " ").strip()
    except Exception:
        pass

    return "(Untitled slide)"


def show_input_slides(presentation):
    print("\n" + "=" * 70)
    print("INPUT SLIDES")
    print("=" * 70)

    for index in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(index)
        print(f"[{index}] {get_slide_title(slide)}")

    print("=" * 70)


# ============================================================
# TEMPLATE MASTER / LAYOUT HELPERS
# ============================================================

def get_template_layouts(presentation):
    layouts = []

    for master_index in range(1, presentation.Designs.Count + 1):
        design = presentation.Designs(master_index)
        master = design.SlideMaster

        try:
            master_name = str(master.Name or "").strip()
        except Exception:
            master_name = ""

        for layout_index in range(1, master.CustomLayouts.Count + 1):
            layout = master.CustomLayouts(layout_index)

            try:
                layout_name = str(layout.Name or "").strip()
            except Exception:
                layout_name = f"Layout {layout_index}"

            layouts.append(
                {
                    "id": len(layouts),
                    "master_index": master_index,
                    "master_name": master_name,
                    "layout_index": layout_index,
                    "layout_name": layout_name,
                }
            )

    return layouts


def get_actual_custom_layout(presentation, layout_info):
    """
    IMPORTANT:
    Always fetch the CustomLayout from the CURRENTLY OPEN
    template presentation. Never store COM layout objects.
    """
    design = presentation.Designs(layout_info["master_index"])
    master = design.SlideMaster
    return master.CustomLayouts(layout_info["layout_index"])


def show_masters(presentation):
    print("\n" + "=" * 70)
    print("TEMPLATE SLIDE MASTERS")
    print("=" * 70)

    for master_index in range(1, presentation.Designs.Count + 1):
        master = presentation.Designs(master_index).SlideMaster

        try:
            master_name = str(master.Name or "").strip()
        except Exception:
            master_name = ""

        print(
            f"MASTER [{master_index - 1}] | "
            f"Name: {master_name} | "
            f"Layouts: {master.CustomLayouts.Count}"
        )

    print("=" * 70)


def show_layouts(layouts):
    print("\n" + "=" * 70)
    print("TEMPLATE LAYOUTS")
    print("=" * 70)

    for item in layouts:
        master_name = item["master_name"] or "(unnamed master)"

        print(
            f'[{item["id"]}] '
            f'{item["layout_name"]} '
            f'| Master: {master_name}'
        )

    print("=" * 70)


def select_layout(layouts):
    while True:
        value = input("Select template layout number: ").strip()

        if value.isdigit():
            index = int(value)

            if 0 <= index < len(layouts):
                return layouts[index]

        print("Invalid layout number. Try again.")


# ============================================================
# SLIDE RANGE / MAPPING
# ============================================================

def parse_slide_range(value, slide_count):
    value = value.strip().lower()

    if value == "all":
        return list(range(1, slide_count + 1))

    result = set()

    for part in value.split(","):
        part = part.strip()

        if not part:
            continue

        if "-" in part:
            start_text, end_text = part.split("-", 1)

            if (
                not start_text.strip().isdigit()
                or not end_text.strip().isdigit()
            ):
                raise ValueError("Invalid slide range.")

            start = int(start_text)
            end = int(end_text)

            if start > end:
                start, end = end, start

            for number in range(start, end + 1):
                if not 1 <= number <= slide_count:
                    raise ValueError(
                        f"Slide {number} is outside the input slide range."
                    )

                result.add(number)

        else:
            if not part.isdigit():
                raise ValueError(f"Invalid slide number: {part}")

            number = int(part)

            if not 1 <= number <= slide_count:
                raise ValueError(
                    f"Slide {number} is outside the input slide range."
                )

            result.add(number)

    if not result:
        raise ValueError("No slides selected.")

    return sorted(result)


def configure_mapping(input_presentation, layouts):
    mapping = {}
    slide_count = input_presentation.Slides.Count

    print("\n" + "=" * 70)
    print("SLIDE -> TEMPLATE LAYOUT MAPPING")
    print("=" * 70)

    print(
        """
IMPORTANT:

Only slides that you explicitly map will receive a new
template Master/Layout.

Any slide you DO NOT map will remain AS-IS.

Examples:
    1
    1,3,5
    2-6
    all

You can create multiple mappings.

Example:
    Slides: 1
    Layout: 3

    Slides: 2-3
    Layout: 24

Press Enter when finished.
"""
    )

    while True:
        value = input(
            "Which input slides should receive a template layout? "
        ).strip()

        if not value:
            break

        try:
            slide_numbers = parse_slide_range(
                value,
                slide_count
            )
        except ValueError as exc:
            print(f"Invalid selection: {exc}")
            continue

        print("\nAvailable layouts:")
        print("-" * 70)

        for item in layouts:
            print(
                f'[{item["id"]}] '
                f'{item["layout_name"]} '
                f'| Master: {item["master_name"]}'
            )

        print("-" * 70)

        layout = select_layout(layouts)

        for slide_number in slide_numbers:
            mapping[slide_number] = layout["id"]

        print(
            f'Mapped slides {slide_numbers} -> '
            f'[{layout["id"]}] {layout["layout_name"]}'
        )

    return mapping


# ============================================================
# SOURCE DESIGN DETECTION
# ============================================================

def build_repeated_text_map(input_presentation):
    occurrences = defaultdict(set)

    for slide_number in range(
        1,
        input_presentation.Slides.Count + 1
    ):
        slide = input_presentation.Slides(slide_number)

        for shape_index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(shape_index)
            text = normalize_text(safe_shape_text(shape))

            if text:
                occurrences[text].add(slide_number)

    return occurrences


def is_full_slide_background_shape(
    shape,
    slide_width,
    slide_height
):
    try:
        left = float(shape.Left)
        top = float(shape.Top)
        width = float(shape.Width)
        height = float(shape.Height)

        return (
            left <= slide_width * 0.01
            and top <= slide_height * 0.01
            and width >= slide_width * 0.97
            and height >= slide_height * 0.97
        )
    except Exception:
        return False


def looks_like_old_header_footer(
    shape,
    repeated_texts,
    slide_width,
    slide_height,
    slide_count
):
    text = normalize_text(safe_shape_text(shape))

    if not text:
        return False

    # Explicit old-deck footer/header strings.
    known_text = {
        "aipatshala.in",
        "ai patshala",
        "‹#›",
    }

    if text in known_text:
        return True

    if text.startswith("source: ai patshala"):
        return True

    occurrences = repeated_texts.get(text, set())

    if (
        slide_count >= 3
        and len(occurrences) >= max(
            3,
            int(slide_count * 0.60)
        )
    ):
        try:
            top = float(shape.Top)
            bottom = top + float(shape.Height)

            if (
                top <= slide_height * 0.12
                or bottom >= slide_height * 0.88
            ):
                return True
        except Exception:
            pass

    return False


# ============================================================
# MAPPED SLIDE CONTENT COPY
# ============================================================

def remove_slide_level_placeholders(destination_slide):
    """
    Delete only slide-level placeholders created by the selected
    CustomLayout.

    The actual Slide Master is NOT touched.
    Master background, graphics, theme and formatting remain.
    """
    removed = 0

    for index in range(
        destination_slide.Shapes.Count,
        0,
        -1
    ):
        try:
            shape = destination_slide.Shapes(index)

            if shape.Type == MSO_PLACEHOLDER:
                shape.Delete()
                removed += 1

        except Exception:
            pass

    return removed


def copy_one_shape(
    source_shape,
    destination_slide
):
    """
    Copy one source shape and restore its original geometry.

    Copy/Paste is used rather than reconstructing text/images/tables,
    because this preserves complex PowerPoint objects much better.
    """

    left = None
    top = None
    width = None
    height = None
    rotation = None

    try:
        left = float(source_shape.Left)
        top = float(source_shape.Top)
        width = float(source_shape.Width)
        height = float(source_shape.Height)
        rotation = float(source_shape.Rotation)
    except Exception:
        pass

    source_shape.Copy()

    pasted = None

    for _ in range(40):
        try:
            pasted_range = destination_slide.Shapes.Paste()
            pasted = pasted_range.Item(1)
            break
        except Exception:
            time.sleep(0.10)

    if pasted is None:
        raise RuntimeError("PowerPoint could not paste the source shape.")

    if left is not None:
        pasted.Left = left
        pasted.Top = top
        pasted.Width = width
        pasted.Height = height
        pasted.Rotation = rotation

    return pasted


def copy_mapped_slide_content(
    source_slide,
    destination_slide,
    input_slide_width,
    input_slide_height,
    repeated_texts,
    input_slide_count
):
    """
    Copies only slide-level source content.

    The source Master/Layout/Theme is NEVER copied here.
    """

    stats = {
        "copied": 0,
        "skipped_background": 0,
        "skipped_header_footer": 0,
        "removed_template_placeholders": 0,
        "failed": 0,
    }

    stats["removed_template_placeholders"] = (
        remove_slide_level_placeholders(destination_slide)
    )

    for shape_index in range(
        1,
        source_slide.Shapes.Count + 1
    ):
        source_shape = source_slide.Shapes(shape_index)

        if is_full_slide_background_shape(
            source_shape,
            input_slide_width,
            input_slide_height
        ):
            stats["skipped_background"] += 1
            continue

        if looks_like_old_header_footer(
            source_shape,
            repeated_texts,
            input_slide_width,
            input_slide_height,
            input_slide_count
        ):
            stats["skipped_header_footer"] += 1
            continue

        try:
            copy_one_shape(
                source_shape,
                destination_slide
            )
            stats["copied"] += 1

        except Exception as exc:
            stats["failed"] += 1
            print(
                f"    WARNING: shape {shape_index} "
                f"could not be copied: {exc}"
            )

    return stats


# ============================================================
# SLIDE SIZE
# ============================================================

def match_slide_size(
    input_presentation,
    template_presentation
):
    input_width = float(
        input_presentation.PageSetup.SlideWidth
    )
    input_height = float(
        input_presentation.PageSetup.SlideHeight
    )

    template_width = float(
        template_presentation.PageSetup.SlideWidth
    )
    template_height = float(
        template_presentation.PageSetup.SlideHeight
    )

    print(
        f"\nInput slide size    : "
        f"{input_width:.1f} x {input_height:.1f}"
    )

    print(
        f"Template slide size : "
        f"{template_width:.1f} x {template_height:.1f}"
    )

    if (
        abs(input_width - template_width) > 0.1
        or abs(input_height - template_height) > 0.1
    ):
        print("Adjusting template slide size to input dimensions...")

        template_presentation.PageSetup.SlideWidth = input_width
        template_presentation.PageSetup.SlideHeight = input_height

        print(
            f"New template size   : "
            f"{float(template_presentation.PageSetup.SlideWidth):.1f} x "
            f"{float(template_presentation.PageSetup.SlideHeight):.1f}"
        )
    else:
        print("Slide sizes already match.")


# ============================================================
# CREATE OUTPUT
# ============================================================

def create_output(
    powerpoint,
    input_path,
    template_path,
    output_path,
    mapping
):
    input_presentation = None
    template_presentation = None

    try:
        print("\n" + "=" * 70)
        print("CREATING OUTPUT")
        print("=" * 70)

        # --------------------------------------------------------
        # OPEN INPUT
        # --------------------------------------------------------

        print("\nLoading input presentation...")

        input_presentation = powerpoint.Presentations.Open(
            str(input_path),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False
        )

        # --------------------------------------------------------
        # OPEN TEMPLATE
        # --------------------------------------------------------

        print("Loading template presentation...")

        template_presentation = powerpoint.Presentations.Open(
            str(template_path),
            ReadOnly=False,
            Untitled=False,
            WithWindow=False
        )

        input_slide_count = input_presentation.Slides.Count

        input_slide_width = float(
            input_presentation.PageSetup.SlideWidth
        )
        input_slide_height = float(
            input_presentation.PageSetup.SlideHeight
        )

        # The output is physically built inside a COPY of the template.
        # The originals are never saved.
        match_slide_size(
            input_presentation,
            template_presentation
        )

        layouts = get_template_layouts(
            template_presentation
        )

        if not layouts:
            raise RuntimeError(
                "No template CustomLayouts were found."
            )

        repeated_texts = build_repeated_text_map(
            input_presentation
        )

        print(
            f"\nInput slides    : "
            f"{input_slide_count}"
        )

        print(
            f"Template masters: "
            f"{template_presentation.Designs.Count}"
        )

        print(
            f"Template layouts: "
            f"{len(layouts)}"
        )

        # IMPORTANT:
        # These are the template's original sample slides.
        # They are deleted only AFTER all output slides are created.
        original_template_slide_count = (
            template_presentation.Slides.Count
        )

        # --------------------------------------------------------
        # BUILD OUTPUT IN INPUT SLIDE ORDER
        # --------------------------------------------------------

        for slide_number in range(
            1,
            input_slide_count + 1
        ):
            source_slide = input_presentation.Slides(
                slide_number
            )

            # ====================================================
            # UNMAPPED: COMPLETE SOURCE SLIDE, AS-IS
            # ====================================================

            if slide_number not in mapping:
                print(
                    f"\nSlide {slide_number:02d} -> "
                    f"KEEP AS-IS"
                )

                source_slide.Copy()

                pasted_slides = None

                for _ in range(40):
                    try:
                        pasted_slides = (
                            template_presentation.Slides.Paste(
                                template_presentation.Slides.Count + 1
                            )
                        )
                        break
                    except Exception:
                        time.sleep(0.10)

                if pasted_slides is None:
                    raise RuntimeError(
                        f"Could not paste unmapped slide "
                        f"{slide_number}."
                    )

                print(
                    "    Original master/theme/background preserved."
                )

                continue

            # ====================================================
            # MAPPED: REAL TEMPLATE MASTER + LAYOUT
            # ====================================================

            layout_info = layouts[
                mapping[slide_number]
            ]

            print(
                f"\nSlide {slide_number:02d} -> "
                f'[{layout_info["id"]}] '
                f'{layout_info["layout_name"]} '
                f'| Master: {layout_info["master_name"]}'
            )

            # Fresh CustomLayout from currently-open template.
            custom_layout = get_actual_custom_layout(
                template_presentation,
                layout_info
            )

            # This slide is genuinely attached to the template
            # Master / Theme / CustomLayout.
            destination_slide = (
                template_presentation.Slides.AddSlide(
                    template_presentation.Slides.Count + 1,
                    custom_layout
                )
            )

            stats = copy_mapped_slide_content(
                source_slide=source_slide,
                destination_slide=destination_slide,
                input_slide_width=input_slide_width,
                input_slide_height=input_slide_height,
                repeated_texts=repeated_texts,
                input_slide_count=input_slide_count
            )

            print(
                f"    Content copied       : "
                f"{stats['copied']}"
            )

            print(
                f"    Old background skip  : "
                f"{stats['skipped_background']}"
            )

            print(
                f"    Old header/footer    : "
                f"{stats['skipped_header_footer']}"
            )

            print(
                f"    Template placeholders: "
                f"{stats['removed_template_placeholders']}"
            )

            print(
                f"    Copy failures        : "
                f"{stats['failed']}"
            )

        # --------------------------------------------------------
        # DELETE ONLY ORIGINAL TEMPLATE SAMPLE SLIDES
        # --------------------------------------------------------

        print(
            f"\nRemoving "
            f"{original_template_slide_count} "
            f"original template sample slide(s)..."
        )

        for index in range(
            original_template_slide_count,
            0,
            -1
        ):
            template_presentation.Slides(index).Delete()

        # --------------------------------------------------------
        # SAVE NEW OUTPUT
        # --------------------------------------------------------

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        if output_path.exists():
            try:
                output_path.unlink()
            except Exception as exc:
                raise RuntimeError(
                    f"Cannot overwrite existing output file. "
                    f"Close it in PowerPoint first.\n{exc}"
                )

        print("\nSaving output:")
        print(output_path)

        template_presentation.SaveAs(
            str(output_path),
            PP_SAVE_AS_OPEN_XML_PRESENTATION
        )

        # PPT creation is complete -> remove temporary layout previews.
        delete_preview_folder()

        print("\n" + "=" * 70)
        print("SUCCESS")
        print("=" * 70)
        print(f"Output: {output_path}")
        print("=" * 70)

    finally:
        if input_presentation is not None:
            try:
                input_presentation.Close()
            except Exception:
                pass

        if template_presentation is not None:
            try:
                template_presentation.Close()
            except Exception:
                pass


# ============================================================
# MAIN
# ============================================================

def main():
    ensure_directories()

    pythoncom.CoInitialize()

    powerpoint = None
    input_presentation = None
    template_presentation = None

    try:
        print("\n" + "=" * 70)
        print("                         PPTX AUTOMATION")
        print("=" * 70)

        input_path = select_file(
            INPUT_DIR,
            "AVAILABLE INPUT PPTX FILES"
        )

        if input_path is None:
            return

        print(f"\nSelected input: {input_path.name}")

        template_path = select_file(
            TEMPLATES_DIR,
            "AVAILABLE TEMPLATE PPTX FILES"
        )

        if template_path is None:
            return

        print(f"\nSelected template: {template_path.name}")

        default_output_name = (
            f"{input_path.stem}_templated.pptx"
        )

        output_name = input(
            f"\nOutput filename "
            f"[{default_output_name}]: "
        ).strip()

        if not output_name:
            output_name = default_output_name

        if not output_name.lower().endswith(".pptx"):
            output_name += ".pptx"

        output_path = OUTPUT_DIR / output_name

        print("\nStarting Microsoft PowerPoint...")

        powerpoint = win32com.client.DispatchEx(
            "PowerPoint.Application"
        )

        powerpoint.Visible = True

        # --------------------------------------------------------
        # INSPECTION ONLY
        # --------------------------------------------------------

        print("\nLoading input presentation...")

        input_presentation = powerpoint.Presentations.Open(
            str(input_path),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False
        )

        print("Loading template presentation...")

        template_presentation = powerpoint.Presentations.Open(
            str(template_path),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False
        )

        input_slide_count = input_presentation.Slides.Count

        print(
            f"\nInput slides    : "
            f"{input_slide_count}"
        )

        print(
            f"Template masters: "
            f"{template_presentation.Designs.Count}"
        )

        template_layout_count = sum(
            template_presentation.Designs(i)
            .SlideMaster
            .CustomLayouts.Count
            for i in range(
                1,
                template_presentation.Designs.Count + 1
            )
        )

        print(
            f"Template layouts: "
            f"{template_layout_count}"
        )

        show_input_slides(
            input_presentation
        )

        show_masters(
            template_presentation
        )

        layouts = get_template_layouts(
            template_presentation
        )

        # Create visual PNG previews for ALL template layouts.
        # These remain available during layout selection and are
        # automatically deleted after the program exits.
        create_layout_preview_images(
            powerpoint=powerpoint,
            template_path=template_path,
            layouts=layouts
        )

        show_layouts(
            layouts
        )

        print(
            "\nTemporary layout previews are available here:"
        )
        print(
            TEMP_PREVIEW_DIR
        )

        mapping = configure_mapping(
            input_presentation,
            layouts
        )

        # --------------------------------------------------------
        # SAVE VALUES NEEDED AFTER INSPECTION FILES ARE CLOSED
        # --------------------------------------------------------

        # DO NOT access input_presentation after this point.
        input_slide_count = input_presentation.Slides.Count

        input_presentation.Close()
        input_presentation = None

        template_presentation.Close()
        template_presentation = None

        print("\n" + "=" * 70)
        print("FINAL MAPPING")
        print("=" * 70)

        if mapping:
            for slide_number in sorted(mapping):
                layout_info = layouts[
                    mapping[slide_number]
                ]

                print(
                    f"Slide {slide_number:02d} -> "
                    f'[{layout_info["id"]}] '
                    f'{layout_info["layout_name"]}'
                )

            unmapped = [
                number
                for number in range(
                    1,
                    input_slide_count + 1
                )
                if number not in mapping
            ]

            if unmapped:
                print(
                    f"Unmapped slides -> KEEP AS-IS: "
                    f"{unmapped}"
                )

        else:
            print(
                "No slides mapped."
            )
            print(
                "All slides will remain AS-IS."
            )

        print("=" * 70)

        confirm = input(
            "\nCreate output with this mapping? [Y/n]: "
        ).strip().lower()

        if confirm not in ("", "y", "yes"):
            print("\nCancelled.")
            delete_preview_folder()
            return

        create_output(
            powerpoint=powerpoint,
            input_path=input_path,
            template_path=template_path,
            output_path=output_path,
            mapping=mapping
        )

    except Exception as exc:
        print("\n" + "=" * 70)
        print("ERROR")
        print("=" * 70)
        print(repr(exc))
        print("=" * 70)

    finally:
        if input_presentation is not None:
            try:
                input_presentation.Close()
            except Exception:
                pass

        if template_presentation is not None:
            try:
                template_presentation.Close()
            except Exception:
                pass

        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass

        # Always remove temporary layout preview images.
        delete_preview_folder()

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
