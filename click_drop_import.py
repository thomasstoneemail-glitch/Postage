"""Convert plaintext address batches into Royal Mail Click & Drop XLSX."""
from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import logging
from logging.handlers import RotatingFileHandler
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional

if importlib.util.find_spec("rich"):
    from rich.console import Console
    from rich.progress import Progress, SpinnerColumn, TextColumn
else:  # pragma: no cover - optional dependency
    Console = None
    Progress = None
    SpinnerColumn = None
    TextColumn = None

POSTCODE_RE = re.compile(
    r"\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b",
    re.IGNORECASE,
)
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE)
PHONE_RE = re.compile(r"\+?\d[\d\s().-]{6,}\d")

COMPANY_KEYWORDS = ("LTD", "LIMITED", "LLP", "INC", "INC.", "PLC", "COMPANY")
COUNTRY_ALIASES = {"UK", "GB", "UNITED KINGDOM", "GREAT BRITAIN"}

DEFAULT_HEADERS = [
    "Order Reference",
    "Full Name",
    "Company Name",
    "Address line 1",
    "Address line 2",
    "Address line 3",
    "City",
    "County",
    "Postcode",
    "Country",
    "Weight",
    "Format",
    "Service Code",
    "Email",
    "Phone",
]


@dataclass
class AddressRecord:
    """Normalized address record for Click & Drop."""

    order_reference: str
    full_name: str
    company_name: str
    address1: str
    address2: str
    address3: str
    city: str
    county: str
    postcode: str
    country: str
    weight: float
    parcel_format: str
    service_code: str
    email: str
    phone: str


@dataclass
class ParseResult:
    """Result of parsing a block."""

    record: Optional[AddressRecord]
    errors: list[str]
    raw_block: str
    block_index: int


def setup_logging(log_dir: Path, debug: bool) -> logging.Logger:
    """Configure rotating logging with redaction."""
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "app.log"
    logger = logging.getLogger("click_drop_import")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    handler = RotatingFileHandler(
        log_path,
        maxBytes=1_000_000,
        backupCount=7,
        encoding="utf-8",
    )
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.handlers = [handler]
    logger.propagate = False
    return logger


def normalize_whitespace(text: str) -> str:
    """Normalize whitespace and stray commas."""
    cleaned = re.sub(r"\s+", " ", text.replace("\t", " ")).strip(" ,")
    return cleaned


def normalize_postcode(raw: str) -> str:
    """Normalize UK postcode to uppercase with single space."""
    raw = raw.strip().upper().replace(" ", "")
    if len(raw) <= 3:
        return raw
    return f"{raw[:-3]} {raw[-3:]}"


def is_country_line(line: str) -> bool:
    return normalize_whitespace(line).upper() in COUNTRY_ALIASES


def detect_company(name_line: str) -> tuple[str, str]:
    """Return full_name, company_name based on keyword heuristics."""
    upper = name_line.upper()
    if any(keyword in upper for keyword in COMPANY_KEYWORDS):
        return "", name_line
    return name_line, ""


def extract_contact_lines(lines: list[str]) -> tuple[list[str], str, str]:
    """Extract email and phone lines, returning remaining lines."""
    email = ""
    phone = ""
    remaining: list[str] = []
    for line in lines:
        line_clean = normalize_whitespace(line)
        if not line_clean:
            continue
        email_match = EMAIL_RE.search(line_clean)
        phone_match = PHONE_RE.search(line_clean)
        if email_match and not email:
            email = email_match.group(0)
        if phone_match and not phone:
            phone = phone_match.group(0)
        if email_match or phone_match:
            continue
        remaining.append(line_clean)
    return remaining, email, phone


def split_blocks(text: str) -> list[str]:
    """Split text into address blocks by blank lines."""
    blocks = [block.strip() for block in re.split(r"\n\s*\n+", text.strip()) if block.strip()]
    return blocks


def parse_block(
    block: str,
    index: int,
    defaults: dict[str, str | float | list[str]],
) -> ParseResult:
    """Parse one address block into a normalized record."""
    raw_lines = [normalize_whitespace(line) for line in block.splitlines()]
    raw_lines = [line for line in raw_lines if line]

    remaining, email, phone = extract_contact_lines(raw_lines)
    errors: list[str] = []

    if remaining and is_country_line(remaining[-1]):
        remaining = remaining[:-1]

    if not remaining:
        return ParseResult(None, ["missing name line"], block, index)

    name_line = remaining[0]
    full_name, company_name = detect_company(name_line)
    if not full_name and not company_name:
        errors.append("missing name/company")

    content_lines = remaining[1:]

    postcode = ""
    city = ""
    postcode_line_index: Optional[int] = None

    for idx in range(len(content_lines) - 1, -1, -1):
        line = content_lines[idx]
        match = POSTCODE_RE.search(line)
        if match:
            postcode = normalize_postcode(match.group(1))
            postcode_line_index = idx
            remainder = normalize_whitespace(POSTCODE_RE.sub("", line))
            if remainder:
                city = remainder
            break

    if not postcode:
        errors.append("missing postcode")
    if postcode_line_index is not None:
        content_lines.pop(postcode_line_index)

    if not city:
        if content_lines:
            city = content_lines.pop(-1)
        else:
            errors.append("missing city")

    address_lines = content_lines
    county = ""
    if len(address_lines) >= 3:
        county = address_lines.pop(-1)

    if not address_lines:
        errors.append("missing address line 1")

    address1 = address_lines[0] if address_lines else ""
    address2 = address_lines[1] if len(address_lines) > 1 else ""
    address3 = ""
    if len(address_lines) > 2:
        address3 = ", ".join(address_lines[2:])

    if errors:
        return ParseResult(None, errors, block, index)

    record = AddressRecord(
        order_reference=defaults["order_reference"],
        full_name=full_name,
        company_name=company_name,
        address1=address1,
        address2=address2,
        address3=address3,
        city=city,
        county=county,
        postcode=postcode,
        country=str(defaults["country"]),
        weight=float(defaults["weight"]),
        parcel_format=str(defaults["format"]),
        service_code=str(defaults["service_code"]),
        email=email,
        phone=phone,
    )
    return ParseResult(record, [], block, index)


def generate_order_references(count: int) -> Iterator[str]:
    """Generate unique order references for a run."""
    date_prefix = dt.datetime.now().strftime("%Y%m%d")
    for idx in range(1, count + 1):
        yield f"{date_prefix}{idx:03d}"


def build_xlsx(records: list[AddressRecord], output_path: Path) -> None:
    """Write records to an XLSX file."""
    if not importlib.util.find_spec("openpyxl"):
        raise ImportError("openpyxl")
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Orders"

    sheet.append(DEFAULT_HEADERS)
    for record in records:
        sheet.append(
            [
                record.order_reference,
                record.full_name,
                record.company_name,
                record.address1,
                record.address2,
                record.address3,
                record.city,
                record.county,
                record.postcode,
                record.country,
                record.weight,
                record.parcel_format,
                record.service_code,
                record.email,
                record.phone,
            ]
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def self_test() -> int:
    """Run a simple parser self-test without writing files."""
    sample = (
        "Jane Doe\n"
        "10 Downing Street\n"
        "London SW1A2AA\n\n"
        "Example Ltd\n"
        "1 Test Road\n"
        "Testshire\n"
        "Testville\n"
        "AB12 3CD\n"
    )
    defaults = {
        "order_reference": "20260101001",
        "country": "GB",
        "weight": 0.5,
        "format": "Parcel",
        "service_code": "",
    }
    blocks = split_blocks(sample)
    results = [parse_block(block, idx + 1, defaults) for idx, block in enumerate(blocks)]
    if len(results) != 2:
        print("Self-test failed: incorrect block count.")
        return 1
    first = results[0].record
    if not first or first.postcode != "SW1A 2AA" or first.city != "London":
        print("Self-test failed: postcode normalization or city detection.")
        return 1
    second = results[1].record
    if not second or second.company_name != "Example Ltd":
        print("Self-test failed: company detection.")
        return 1
    print("Self-test passed.")
    return 0


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(
        description="Convert plaintext address batches into Click & Drop XLSX",
    )
    parser.add_argument("--input", required=False, help="Path to input text file")
    parser.add_argument("--outdir", default=".\\output", help="Output directory")
    parser.add_argument(
        "--filename",
        default=None,
        help="Output XLSX filename (default orders_YYYYMMDD_HHMMSS.xlsx)",
    )
    parser.add_argument("--default-weight", type=float, default=0.5)
    parser.add_argument("--default-format", default="Parcel")
    parser.add_argument("--service-code", default="")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--save-rejects", action="store_true")
    parser.add_argument("--quiet", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args()


def print_status(console: Optional[Console], quiet: bool, message: str) -> None:
    if quiet:
        return
    if console:
        console.print(message)
    else:
        print(message)


def process_blocks(
    blocks: list[str],
    defaults: dict[str, str | float | list[str]],
    logger: logging.Logger,
    console: Optional[Console],
    quiet: bool,
    save_rejects: bool,
) -> tuple[list[AddressRecord], list[ParseResult]]:
    """Process blocks into records with progress feedback."""
    records: list[AddressRecord] = []
    rejects: list[ParseResult] = []

    if console and Progress and not quiet:
        with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
            task = progress.add_task("Parsing address blocks...", total=len(blocks))
            for idx, block in enumerate(blocks, start=1):
                defaults["order_reference"] = defaults["order_refs"][idx - 1]
                result = parse_block(block, idx, defaults)
                if result.record:
                    records.append(result.record)
                else:
                    if not save_rejects:
                        result.raw_block = ""
                    rejects.append(result)
                    logger.warning("Block %s rejected: %s", idx, ", ".join(result.errors))
                progress.advance(task)
    else:
        for idx, block in enumerate(blocks, start=1):
            defaults["order_reference"] = defaults["order_refs"][idx - 1]
            result = parse_block(block, idx, defaults)
            if result.record:
                records.append(result.record)
            else:
                if not save_rejects:
                    result.raw_block = ""
                rejects.append(result)
                logger.warning("Block %s rejected: %s", idx, ", ".join(result.errors))
            if not quiet:
                print_status(console, quiet, f"Processed {idx}/{len(blocks)} blocks")

    return records, rejects


def main() -> int:
    args = parse_args()

    if args.self_test:
        return self_test()

    if not args.input:
        print("Error: --input is required unless --self-test is used.")
        return 4

    input_path = Path(args.input)
    outdir = Path(args.outdir)

    log_dir = Path(".\\logs")
    logger = setup_logging(log_dir, args.debug)

    console = Console() if Console else None

    try:
        text = input_path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError as exc:
        print_status(console, args.quiet, f"IO error reading input: {exc}")
        return 4

    blocks = split_blocks(text)
    del text

    if not blocks:
        print_status(console, args.quiet, "No address blocks found.")
        return 0

    order_refs = list(generate_order_references(len(blocks)))
    defaults: dict[str, str | float | list[str]] = {
        "order_reference": "",
        "order_refs": order_refs,
        "country": "GB",
        "weight": args.default_weight,
        "format": args.default_format,
        "service_code": args.service_code,
    }

    records, rejects = process_blocks(
        blocks,
        defaults,
        logger,
        console,
        args.quiet,
        args.save_rejects,
    )

    if rejects:
        print_status(console, args.quiet, "Validation errors encountered:")
        for reject in rejects:
            summary = ", ".join(reject.errors)
            print_status(console, args.quiet, f" - Block {reject.block_index}: {summary}")

    if rejects and args.strict:
        print_status(console, args.quiet, "Strict mode enabled: aborting due to validation errors.")
        return 2

    if args.save_rejects and rejects:
        reject_path = outdir / "rejects.txt"
        try:
            reject_path.parent.mkdir(parents=True, exist_ok=True)
            with reject_path.open("w", encoding="utf-8") as handle:
                for reject in rejects:
                    handle.write(reject.raw_block.strip())
                    handle.write("\n\n")
            print_status(console, args.quiet, f"Rejects saved to {reject_path.resolve()}")
        except OSError as exc:
            print_status(console, args.quiet, f"Failed to write rejects file: {exc}")
            return 4

    if not records:
        print_status(console, args.quiet, "No valid records to export.")
        return 0

    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = args.filename or f"orders_{timestamp}.xlsx"
    output_path = (outdir / filename).resolve()

    try:
        build_xlsx(records, output_path)
    except ImportError:
        print_status(
            console,
            args.quiet,
            "Missing dependency: openpyxl. Install with: py -m pip install openpyxl rich",
        )
        return 3
    except OSError as exc:
        print_status(console, args.quiet, f"IO error writing XLSX: {exc}")
        return 4

    print_status(console, args.quiet, "Summary:")
    print_status(console, args.quiet, f" - Blocks found: {len(blocks)}")
    print_status(console, args.quiet, f" - Rows exported: {len(records)}")
    print_status(console, args.quiet, f" - Rows rejected: {len(rejects)}")
    print_status(console, args.quiet, f" - Output XLSX: {output_path}")
    print_status(
        console,
        args.quiet,
        "Next step: Copy the XLSX into your Click & Drop Desktop Watch folder manually.",
    )
    print_status(
        console,
        args.quiet,
        "Reminder: Consider securely deleting the input/output files and clearing Recycle Bin if privacy-sensitive.",
    )
    print_status(
        console,
        args.quiet,
        "Tip: Avoid placing sensitive files in cloud-synced folders (e.g., OneDrive) unless intended.",
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
