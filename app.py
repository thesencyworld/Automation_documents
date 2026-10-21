from flask import Flask, render_template, request, send_file, flash
from docx import Document
from pathlib import Path
from copy import deepcopy
from datetime import datetime
from werkzeug.utils import secure_filename
import os
import re
import shutil
import zipfile
import tempfile
import urllib.request
import urllib.parse
import urllib.error
import socket
import time
import uuid


# ============================================================
# PATHS / RUNTIME
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "public" / "static"

# Vercel only gives writable temporary storage during a request.
RUNTIME_ROOT = Path("/tmp") if os.environ.get("VERCEL") else BASE_DIR / "runtime"
RUNTIME_ROOT.mkdir(parents=True, exist_ok=True)


# ============================================================
# FLASK
# ============================================================

app = Flask(
    __name__,
    template_folder=str(TEMPLATE_DIR),
    static_folder=str(STATIC_DIR),
)
app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY",
    "solar-document-automation"
)


# ============================================================
# PDF CONVERTER
# ============================================================
# Set this in Vercel:
# GOTENBERG_URL=https://YOUR-GOTENBERG-SERVICE.onrender.com
#
# The companion Docker service in pdf-service/ runs Gotenberg and
# performs DOCX -> PDF with LibreOffice outside Vercel.

GOTENBERG_URL = os.environ.get("GOTENBERG_URL", "").rstrip("/")


# ============================================================
# DOCUMENTS
# ============================================================

DOCUMENTS = {
    "declaration_dcr": (
        "declaration_dcr.docx",
        "Declaration_DCR.docx",
    ),
    "model_agreement": (
        "model_agreement.docx",
        "Model_Agreement.docx",
    ),
    "interconnection_agreement": (
        "interconnection_agreement.docx",
        "Interconnection_Agreement.docx",
    ),
    "self_certificate": (
        "self_certificate.docx",
        "Self_Certificate.docx",
    ),
    "project_completion_report": (
        "project_completion_report.docx",
        "Project_Completion_Report.docx",
    ),
}


# ============================================================
# IMAGE TYPES
# ============================================================

ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp"}


# ============================================================
# DATE
# ============================================================

def normalize_date(value: str):
    if not value:
        return "", "", ""

    try:
        dt = datetime.strptime(value, "%Y-%m-%d")
        return (
            dt.strftime("%d"),
            dt.strftime("%B").upper(),
            dt.strftime("%Y"),
        )
    except ValueError:
        return value, value, value


# ============================================================
# SERIAL NUMBERS
# ============================================================

def clean_serials(raw: str):
    return [
        x.strip()
        for x in re.split(r"[\n,;]+", raw or "")
        if x.strip()
    ]


# ============================================================
# FORM DATA
# ============================================================

def data_from_form(form):
    serial_inputs = form.getlist("panel_serial_no")

    if serial_inputs:
        serials = [
            x.strip()
            for x in serial_inputs
            if x and x.strip()
        ]
    else:
        serials = clean_serials(
            form.get("panels_serial_no", "")
        )

    day, month, year = normalize_date(
        form.get("date_agreement", "")
    )

    return {
        "consumer_name": form.get("consumer_name", "").strip(),
        "consumer_number": form.get("consumer_number", "").strip(),
        "address": form.get("address", "").strip(),
        "place": form.get("place", "").strip(),
        "application_number": form.get("application_number", "").strip(),
        "pv_module_capacity": form.get("pv_module_capacity", "").strip(),
        "number_of_modules": form.get("number_of_modules", "").strip(),
        "pv_module_make": form.get("pv_module_make", "").strip(),
        "cell_manufacturer": form.get("cell_manufacturer", "").strip(),
        "cell_gst_invoice": form.get("cell_gst_invoice", "").strip(),
        "date_agreement": form.get("date_agreement", "").strip(),
        "date_day": day,
        "date_month": month,
        "date_year": year,
        "kilo_watt_capacity": form.get("kilo_watt_capacity", "").strip(),
        "consumer_sign": form.get("consumer_sign", "").strip(),
        "witness_sign": form.get("witness_sign", "").strip(),
        "panels_serial_no": serials,
        "inverter_make": form.get("inverter_make", "").strip(),
        "inverter_capacity": form.get("inverter_capacity", "").strip(),
        "inverter_model": form.get("inverter_model", "").strip(),
        "inverter_sr_no": form.get("inverter_sr_no", "").strip(),
    }


# ============================================================
# REPLACEMENTS
# ============================================================

def replacements(d):
    serial_mapping = {}

    old_serials = [
        "AS2605152B2342",
        "AS2605152B2396",
        "AS2605152B2277",
        "AS2605152B2381",
        "AS2605152B2523",
    ]

    for i, old_serial in enumerate(old_serials):
        serial_mapping[old_serial] = (
            d["panels_serial_no"][i]
            if i < len(d["panels_serial_no"])
            else ""
        )

    return {
        "<CONSUMER-NAME>": d["consumer_name"],
        "<CONSUMER-NUMBER>": d["consumer_number"],
        "<CONSUMER-ADDRESS>": d["address"],
        "<PLACE>": d["place"],
        "<APPLICATION-NUMBER>": d["application_number"],
        "<PV-KW>": (
            d["kilo_watt_capacity"]
            .replace(" KW", "")
            .replace("KW", "")
            .strip()
        ),
        "<PV-MODULE-CAPACITY>": (
            d["pv_module_capacity"]
            .replace(" Wp", "")
            .replace("Wp", "")
            .strip()
        ),
        "<NUMBER-OF-PV-MODULE>": d["number_of_modules"].zfill(2),
        "<PV-MODULE-MAKE>": d["pv_module_make"],
        "<CELL-MENUFACTURE>": d["cell_manufacturer"],
        "<GST-NUMBER>": d["cell_gst_invoice"],
        "<DATE>": d["date_day"],
        "<MONTH>": d["date_month"],
        "<YEAR>": d["date_year"],
        "<INV-MAKE>": d["inverter_make"],
        "<INVERTER-CAPACITY>": d["inverter_capacity"],
        "<INVERTER-SR-NO>": d["inverter_sr_no"],
        "<INV-MODEL>": d["inverter_model"],
        "<PV-KWP>": (
            d["kilo_watt_capacity"]
            .replace(" KW", " KWP")
            .replace("KW", "KWP")
            .strip()
        ),
        **serial_mapping,
    }


# ============================================================
# TEXT REPLACEMENT
# ============================================================

def replace_paragraph(paragraph, mapping):
    if not paragraph.runs:
        return

    for placeholder, value in mapping.items():
        value = "" if value is None else str(value)

        while True:
            full_text = "".join(
                run.text or "" for run in paragraph.runs
            )

            if placeholder not in full_text:
                break

            start = full_text.find(placeholder)
            end = start + len(placeholder)

            char_pos = 0
            start_run_index = None
            end_run_index = None
            start_offset = 0
            end_offset = 0

            for i, run in enumerate(paragraph.runs):
                run_text = run.text or ""
                run_start = char_pos
                run_end = char_pos + len(run_text)

                if start_run_index is None and (
                    run_start <= start < run_end
                ):
                    start_run_index = i
                    start_offset = start - run_start

                if run_start < end <= run_end:
                    end_run_index = i
                    end_offset = end - run_start
                    break

                char_pos = run_end

            if start_run_index is None or end_run_index is None:
                break

            if start_run_index == end_run_index:
                run = paragraph.runs[start_run_index]
                text = run.text or ""
                run.text = (
                    text[:start_offset]
                    + value
                    + text[end_offset:]
                )
            else:
                first_run = paragraph.runs[start_run_index]
                last_run = paragraph.runs[end_run_index]

                first_text = first_run.text or ""
                last_text = last_run.text or ""

                first_run.text = first_text[:start_offset] + value

                for i in range(
                    start_run_index + 1,
                    end_run_index,
                ):
                    paragraph.runs[i].text = ""

                last_run.text = last_text[end_offset:]


def replace_in_table(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_paragraph(paragraph, mapping)

            for nested in cell.tables:
                replace_in_table(nested, mapping)


def replace_in_document(doc, mapping):
    for paragraph in doc.paragraphs:
        replace_paragraph(paragraph, mapping)

    for table in doc.tables:
        replace_in_table(table, mapping)

    for section in doc.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                replace_paragraph(paragraph, mapping)

            for table in part.tables:
                replace_in_table(table, mapping)


# ============================================================
# SERIAL TABLE
# ============================================================

def _set_cell_text_preserve_style(cell, text, bold=True):
    paragraph = (
        cell.paragraphs[0]
        if cell.paragraphs
        else cell.add_paragraph()
    )

    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)

    run = paragraph.add_run(str(text))
    run.bold = bold
    return run


def _copy_row_style(source_row, target_row):
    for source_cell, target_cell in zip(
        source_row.cells,
        target_row.cells,
    ):
        if source_cell._tc.tcPr is not None:
            target_tcPr = target_cell._tc.get_or_add_tcPr()
            target_tcPr.clear_content()

            for child in source_cell._tc.tcPr:
                target_tcPr.append(deepcopy(child))

        if source_cell.paragraphs and target_cell.paragraphs:
            source_pPr = source_cell.paragraphs[0]._p.pPr
            target_pPr = (
                target_cell.paragraphs[0]._p.get_or_add_pPr()
            )

            if source_pPr is not None:
                target_pPr.clear_content()

                for child in source_pPr:
                    target_pPr.append(deepcopy(child))


def add_serial_table_if_needed(doc, serials):
    if not serials:
        return False

    for table in doc.tables:
        heading_index = None

        for i, row in enumerate(table.rows):
            row_text = " ".join(
                (cell.text or "").strip()
                for cell in row.cells
            ).upper()

            if "PANELS SERIAL NO." in row_text:
                heading_index = i
                break

        if heading_index is None:
            continue

        first_panel_index = heading_index + 1

        if first_panel_index >= len(table.rows):
            return False

        blank_template_index = min(
            first_panel_index + 4,
            len(table.rows) - 1,
        )
        blank_template_row = table.rows[blank_template_index]

        available_rows = len(table.rows) - first_panel_index

        while available_rows < len(serials):
            new_row_xml = deepcopy(blank_template_row._element)
            table._tbl.append(new_row_xml)
            available_rows += 1

        panel_rows = list(
            table.rows[first_panel_index:]
        )

        for i, serial in enumerate(serials):
            row = panel_rows[i]

            if len(row.cells) >= 1:
                _set_cell_text_preserve_style(
                    row.cells[0],
                    i + 1,
                    True,
                )

            if len(row.cells) >= 2:
                _set_cell_text_preserve_style(
                    row.cells[1],
                    serial,
                    True,
                )

        for row in panel_rows[len(serials):]:
            if len(row.cells) >= 1:
                _set_cell_text_preserve_style(
                    row.cells[0],
                    "",
                    False,
                )

            if len(row.cells) >= 2:
                _set_cell_text_preserve_style(
                    row.cells[1],
                    "",
                    False,
                )

        return True

    return False


# ============================================================
# IMAGE REPLACEMENT
# ============================================================

def replace_images_in_docx(docx_path, new_image_path):
    docx_path = Path(docx_path)
    new_image_path = Path(new_image_path)

    if not new_image_path.exists():
        return False

    temp_path = docx_path.with_name(
        docx_path.stem + "_image_tmp.docx"
    )

    new_image_data = new_image_path.read_bytes()
    replacement_extensions = (".bmp")

    replaced = False

    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            with zipfile.ZipFile(
                temp_path,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as zout:

                for item in zin.infolist():
                    filename = item.filename.lower()

                    if (
                        filename.startswith("word/media/")
                        and filename.endswith(replacement_extensions)
                        and not replaced
                    ):
                        zout.writestr(
                            item,
                            new_image_data,
                        )
                        replaced = True
                    else:
                        zout.writestr(
                            item,
                            zin.read(item.filename),
                        )

        temp_path.replace(docx_path)
        return replaced

    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        return False


def save_uploaded_image(image_file, work_dir):
    if not image_file or not image_file.filename:
        raise ValueError("Please select an image.")

    original_name = secure_filename(image_file.filename)
    extension = Path(original_name).suffix.lower()

    if extension not in ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError(
            "Only PNG, JPG, JPEG and BMP images are allowed."
        )

    image_path = work_dir / (
        "replacement_image" + extension
    )
    image_file.save(image_path)

    return image_path


# ============================================================
# HTTP MULTIPART HELPERS
# ============================================================

def _multipart_body(fields, file_field, filename, file_bytes):
    boundary = (
        "----PythonMultipart"
        + uuid.uuid4().hex
    ).encode("ascii")

    body = bytearray()

    for name, value in fields.items():
        body.extend(b"--" + boundary + b"\r\n")
        body.extend(
            (
                f'Content-Disposition: form-data; '
                f'name="{name}"\r\n\r\n'
            ).encode("utf-8")
        )
        body.extend(str(value).encode("utf-8"))
        body.extend(b"\r\n")

    body.extend(b"--" + boundary + b"\r\n")
    body.extend(
        (
            f'Content-Disposition: form-data; '
            f'name="{file_field}"; '
            f'filename="{filename}"\r\n'
        ).encode("utf-8")
    )
    body.extend(
        b"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    )
    body.extend(file_bytes)
    body.extend(b"\r\n")
    body.extend(b"--" + boundary + b"--\r\n")

    return bytes(body), (
        "multipart/form-data; boundary="
        + boundary.decode("ascii")
    )


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convert DOCX using the external Gotenberg service.

    A transient 502/503/504 from the converter is retried. The Gunicorn
    worker timeout must be higher than the maximum time allowed here.
    """
    if not GOTENBERG_URL:
        raise RuntimeError(
            "PDF converter is not configured. Set the Render environment "
            "variable GOTENBERG_URL to the public URL of your Gotenberg service."
        )

    endpoint = GOTENBERG_URL + "/forms/libreoffice/convert"
    body, content_type = _multipart_body(
        {},
        "files",
        docx_path.name,
        docx_path.read_bytes(),
    )

    max_attempts = 3
    retryable_statuses = {429, 500, 502, 503, 504}
    last_error = None

    for attempt in range(1, max_attempts + 1):
        req = urllib.request.Request(
            endpoint,
            data=body,
            method="POST",
            headers={
                "Content-Type": content_type,
                "Content-Length": str(len(body)),
                "User-Agent": "Solar-Document-Automation/1.1",
            },
        )

        try:
            with urllib.request.urlopen(req, timeout=45) as response:
                pdf_bytes = response.read()

            if not pdf_bytes.startswith(b"%PDF"):
                raise RuntimeError(
                    "PDF converter responded, but its response was not a valid PDF."
                )

            pdf_path.write_bytes(pdf_bytes)
            return

        except urllib.error.HTTPError as exc:
            last_error = exc
            # Read a small part of the upstream response for useful logs.
            try:
                detail = exc.read(500).decode("utf-8", errors="replace").strip()
            except Exception:
                detail = ""
            app.logger.warning(
                "Gotenberg conversion failed (attempt %s/%s, HTTP %s). %s",
                attempt, max_attempts, exc.code, detail[:300]
            )
            if exc.code not in retryable_statuses or attempt == max_attempts:
                raise RuntimeError(
                    f"PDF conversion service failed with HTTP {exc.code}. "
                    "Check that GOTENBERG_URL points to a running Gotenberg service."
                ) from exc

        except (urllib.error.URLError, TimeoutError, socket.timeout, ConnectionError) as exc:
            last_error = exc
            app.logger.warning(
                "Gotenberg connection failed (attempt %s/%s): %s",
                attempt, max_attempts, exc
            )
            if attempt == max_attempts:
                raise RuntimeError(
                    "Could not reach the PDF conversion service after 3 attempts. "
                    "Check the Gotenberg Render service and GOTENBERG_URL."
                ) from exc

        if attempt < max_attempts:
            time.sleep(2 ** attempt)

    raise RuntimeError(f"PDF conversion failed: {last_error}")


# ============================================================
# GENERATE DOCUMENTS
# ============================================================

def generate_documents(data, image_path, work_dir):
    mapping = replacements(data)
    output_paths = []

    for key, (template_name, output_name) in DOCUMENTS.items():
        source = TEMPLATE_DIR / template_name

        if not source.exists():
            raise FileNotFoundError(
                f"Template not found: {source}"
            )

        docx_target = work_dir / output_name

        doc = Document(source)

        replace_in_document(
            doc,
            mapping,
        )

        if key == "project_completion_report":
            add_serial_table_if_needed(
                doc,
                data["panels_serial_no"],
            )

        doc.save(docx_target)

        replace_images_in_docx(
            docx_target,
            image_path,
        )

        pdf_target = work_dir / (
            Path(output_name).stem + ".pdf"
        )

        convert_docx_to_pdf(
            docx_target,
            pdf_target,
        )

        output_paths.append(
            (docx_target, pdf_target)
        )

    return output_paths


# ============================================================
# ZIP
# ============================================================

def make_zip(pairs, work_dir):
    zip_path = work_dir / "Generated_Documents.zip"

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED,
    ) as zf:
        for docx_path, pdf_path in pairs:
            zf.write(
                docx_path,
                arcname=docx_path.name,
            )
            zf.write(
                pdf_path,
                arcname=pdf_path.name,
            )

    return zip_path


# ============================================================
# HOME
# ============================================================

@app.route("/", methods=["GET", "POST"])
def index():
    defaults = {
        "consumer_name": "",
        "consumer_number": "",
        "address": "",
        "place": "",
        "application_number": "",
        "pv_module_capacity": "",
        "number_of_modules": "",
        "pv_module_make": "",
        "cell_manufacturer": "",
        "cell_gst_invoice": "",
        "date_agreement": "",
        "kilo_watt_capacity": "",
        "consumer_sign": "",
        "witness_sign": "",
        "panels_serial_no": "",
        "inverter_make": "",
        "inverter_capacity": "",
        "inverter_model": "",
        "inverter_sr_no": "",
    }

    if request.method == "POST":
        work_dir = Path(
            tempfile.mkdtemp(
                prefix="docgen_",
                dir=str(RUNTIME_ROOT),
            )
        )

        try:
            data = data_from_form(request.form)

            required_fields = [
                "consumer_name",
                "consumer_number",
                "address",
                "place",
                "application_number",
                "pv_module_capacity",
                "number_of_modules",
                "pv_module_make",
                "cell_manufacturer",
                "date_agreement",
                "kilo_watt_capacity",
                "inverter_make",
                "inverter_capacity",
                "inverter_model",
                "inverter_sr_no",
            ]

            missing = [
                k
                for k in required_fields
                if not data.get(k)
            ]

            if missing:
                flash(
                    "Please complete all required fields.",
                    "error",
                )
                defaults.update(
                    request.form.to_dict()
                )
                defaults["panels_serial_no"] = "\n".join(
                    data["panels_serial_no"]
                )
                return render_template(
                    "index.html",
                    form=defaults,
                )

            try:
                module_count = int(
                    str(data["number_of_modules"]).strip()
                )
            except ValueError:
                module_count = 0

            if module_count < 1 or module_count > 500:
                flash(
                    "Number of PV Modules must be between 1 and 500.",
                    "error",
                )
                defaults.update(
                    request.form.to_dict()
                )
                defaults["panels_serial_no"] = "\n".join(
                    data["panels_serial_no"]
                )
                return render_template(
                    "index.html",
                    form=defaults,
                )

            if len(data["panels_serial_no"]) != module_count:
                flash(
                    f"Please enter exactly {module_count} panel serial numbers.",
                    "error",
                )
                defaults.update(
                    request.form.to_dict()
                )
                defaults["panels_serial_no"] = "\n".join(
                    data["panels_serial_no"]
                )
                return render_template(
                    "index.html",
                    form=defaults,
                )

            image_file = request.files.get(
                "document_image"
            )

            if not image_file or not image_file.filename:
                flash(
                    "Please select an image.",
                    "error",
                )
                defaults.update(
                    request.form.to_dict()
                )
                defaults["panels_serial_no"] = "\n".join(
                    data["panels_serial_no"]
                )
                return render_template(
                    "index.html",
                    form=defaults,
                )

            image_path = save_uploaded_image(
                image_file,
                work_dir,
            )

            pairs = generate_documents(
                data,
                image_path,
                work_dir,
            )

            zip_path = make_zip(
                pairs,
                work_dir,
            )

            # Files are stored under /tmp and remain available only for
            # the lifetime of the current Vercel instance. The response
            # immediately downloads the ZIP, avoiding persistent storage.
            return send_file(
                zip_path,
                as_attachment=True,
                download_name="Generated_Documents.zip",
                mimetype="application/zip",
            )

        except Exception as exc:
            app.logger.exception(
                "Document generation error"
            )

            flash(
                f"Error: {exc}",
                "error",
            )

            defaults.update(
                request.form.to_dict()
            )
            defaults["panels_serial_no"] = "\n".join(
                data_from_form(request.form)["panels_serial_no"]
            )

            return render_template(
                "index.html",
                form=defaults,
            )

        finally:
            # Do not delete work_dir here: send_file may still be reading
            # the ZIP. Vercel will clean /tmp automatically.
            pass

    return render_template(
        "index.html",
        form=defaults,
    )


# ============================================================
# HEALTH
# ============================================================

@app.route("/health")
def health():
    return {
        "status": "ok",
        "pdf_converter_configured": bool(GOTENBERG_URL),
    }


# ============================================================
# LOCAL DOWNLOAD ROUTE
# ============================================================

@app.route("/download/<path:filename>")
def download(filename):
    # Kept for compatibility with existing templates.
    return "Files are downloaded automatically after generation.", 404


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=True,
    )
