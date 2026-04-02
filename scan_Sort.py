#!/usr/bin/env python3
"""Scan- & Sort-Tool mit GUI (NAPS2 + Tesseract).

Features:
- GUI zum Starten und Überwachen des Scan-/Sortierprozesses
- Scanquelle wählbar: ADF oder Glasplatte
- Automatische Ausrichtung via Tesseract OSD + Pillow-Rotation
- OCR-basierte Sortierung in automatisch angelegte, sinnvolle Ordnerstruktur
"""

from __future__ import annotations

import argparse
import csv
import json
import shutil
import subprocess
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
from typing import Callable, Dict, Iterable, List, Optional

try:
    import tkinter as tk
    from tkinter import ttk
except Exception:  # noqa: BLE001
    tk = None
    ttk = None

try:
    from PIL import Image
except Exception:  # noqa: BLE001
    Image = None


DEFAULT_RULES: Dict[str, List[str]] = {
    "Rechnungen": ["rechnung", "invoice", "betrag", "zahlbar", "kundennummer"],
    "Versicherung": ["versicherung", "police", "schadennummer", "haftpflicht"],
    "Steuer": ["finanzamt", "steuer", "steuerbescheid", "elster"],
    "Gesundheit": ["arzt", "krankenkasse", "diagnose", "rezept", "labor"],
    "Verträge": ["vertrag", "laufzeit", "kündigung", "vertragspartner"],
}


@dataclass
class PageResult:
    source: Path
    destination: Path
    category: str
    matched_keywords: List[str]


def run_command(cmd: List[str], verbose: bool = False) -> subprocess.CompletedProcess:
    if verbose:
        print("[CMD]", " ".join(cmd))
    return subprocess.run(cmd, capture_output=True, text=True, check=False)


def run_scan(
    naps2_path: str,
    profile: str,
    batch_dir: Path,
    image_format: str,
    source: str,
    verbose: bool,
) -> List[Path]:
    output_pattern = batch_dir / f"scan_%04d.{image_format}"

    cmd = [
        naps2_path,
        "scan",
        "--profile",
        profile,
        "--output",
        str(output_pattern),
    ]

    # Je nach NAPS2-Version kann --source variieren. Diese Werte funktionieren bei
    # vielen Setups; falls nicht, bitte im Profil selbst hinterlegen.
    source_map = {
        "adf": "feeder",
        "glasplatte": "glass",
    }
    cmd.extend(["--source", source_map[source]])

    proc = run_command(cmd, verbose=verbose)
    if proc.returncode != 0:
        raise RuntimeError(
            "NAPS2-Scan fehlgeschlagen.\n"
            f"Returncode: {proc.returncode}\n"
            f"STDOUT: {proc.stdout.strip()}\n"
            f"STDERR: {proc.stderr.strip()}"
        )

    pages = sorted(batch_dir.glob(f"*.{image_format}"))
    if not pages:
        raise RuntimeError("Keine Bilddateien erzeugt. Prüfe Scanner, Profil und Ausgabeformat.")

    return pages


def detect_rotation_degrees(tesseract_path: str, image_path: Path, verbose: bool) -> int:
    """Liest OSD-Daten und liefert nötige Rotationsgrade (0/90/180/270)."""
    cmd = [tesseract_path, str(image_path), "stdout", "--psm", "0"]
    proc = run_command(cmd, verbose=verbose)
    if proc.returncode != 0:
        return 0

    for line in proc.stdout.splitlines():
        if "Rotate:" in line:
            value = line.split(":", 1)[1].strip()
            try:
                deg = int(value)
                if deg in {0, 90, 180, 270}:
                    return deg
            except ValueError:
                return 0
    return 0


def auto_orient_pages(
    pages: Iterable[Path],
    tesseract_path: str,
    verbose: bool,
    status_cb: Optional[Callable[[str], None]] = None,
) -> None:
    """Automatische Seitenausrichtung via OSD + Pillow."""
    if Image is None:
        if status_cb:
            status_cb("Ausrichtung übersprungen: Pillow nicht installiert.")
        return

    for idx, page in enumerate(pages, start=1):
        deg = detect_rotation_degrees(tesseract_path=tesseract_path, image_path=page, verbose=verbose)
        if deg == 0:
            continue
        if status_cb:
            status_cb(f"Ausrichtung Seite {idx}: rotiere um {deg}°")
        with Image.open(page) as img:
            rotated = img.rotate(-deg, expand=True)
            rotated.save(page)


def ocr_page(tesseract_path: str, image_path: Path, lang: str, verbose: bool) -> str:
    cmd = [tesseract_path, str(image_path), "stdout", "-l", lang]
    proc = run_command(cmd, verbose=verbose)
    if proc.returncode != 0:
        raise RuntimeError(f"Tesseract-Fehler bei {image_path.name}: {proc.stderr.strip()}")
    return proc.stdout


def load_rules(rules_file: Path | None) -> Dict[str, List[str]]:
    if not rules_file:
        return DEFAULT_RULES

    with rules_file.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict) or not data:
        raise ValueError("Regeldatei muss sein: {\"Kategorie\": [\"keyword\", ...]}")

    normalized: Dict[str, List[str]] = {}
    for category, keywords in data.items():
        if not isinstance(category, str) or not category.strip():
            raise ValueError("Kategorie muss ein nicht-leerer String sein.")
        if not isinstance(keywords, list) or not keywords:
            raise ValueError(f"Kategorie '{category}' braucht eine Keyword-Liste.")
        normalized[category.strip()] = [str(k).strip().lower() for k in keywords if str(k).strip()]

    return normalized


def find_category(text: str, rules: Dict[str, List[str]], default_category: str) -> tuple[str, List[str]]:
    lower = text.lower()
    best_category = default_category
    best_matches: List[str] = []

    for category, keywords in rules.items():
        matches = [kw for kw in keywords if kw in lower]
        if len(matches) > len(best_matches):
            best_matches = matches
            best_category = category

    return best_category, best_matches


def safe_filename(path: Path) -> Path:
    if not path.exists():
        return path

    stem, suffix = path.stem, path.suffix
    i = 1
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def build_destination_dir(target_root: Path, category: str) -> Path:
    """Sinnvolle Struktur: sorted/<Jahr>/<Monat>/<Kategorie>/"""
    year = datetime.now().strftime("%Y")
    month = datetime.now().strftime("%m")
    dest = target_root / year / month / category
    dest.mkdir(parents=True, exist_ok=True)
    return dest


def sort_pages(
    pages: Iterable[Path],
    target_root: Path,
    tesseract_path: str,
    lang: str,
    rules: Dict[str, List[str]],
    default_category: str,
    verbose: bool,
    status_cb: Optional[Callable[[str], None]] = None,
) -> List[PageResult]:
    results: List[PageResult] = []
    page_list = list(pages)

    for i, page in enumerate(page_list, start=1):
        if status_cb:
            status_cb(f"OCR Seite {i}/{len(page_list)}: {page.name}")
        text = ocr_page(tesseract_path=tesseract_path, image_path=page, lang=lang, verbose=verbose)
        category, matched_keywords = find_category(text=text, rules=rules, default_category=default_category)

        destination_dir = build_destination_dir(target_root=target_root, category=category)
        destination = safe_filename(destination_dir / page.name)
        shutil.move(str(page), destination)

        results.append(PageResult(page, destination, category, matched_keywords))

        if status_cb:
            match_text = ", ".join(matched_keywords) if matched_keywords else "kein Treffer"
            status_cb(f"Sortiert -> {category} ({match_text})")

    return results


def write_report(batch_dir: Path, results: List[PageResult]) -> Path:
    report_path = batch_dir / "sort_report.csv"
    with report_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["quelle", "ziel", "kategorie", "treffer"])
        for row in results:
            writer.writerow([str(row.source), str(row.destination), row.category, "; ".join(row.matched_keywords)])
    return report_path


def process_scan_and_sort(
    profile: str,
    scan_root: Path,
    target_root: Path,
    naps2_path: str,
    tesseract_path: str,
    lang: str,
    image_format: str,
    rules_file: Optional[Path],
    default_category: str,
    source: str,
    auto_orient: bool,
    verbose: bool,
    status_cb: Optional[Callable[[str], None]] = None,
) -> tuple[List[PageResult], Path]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    batch_dir = scan_root / f"batch_{timestamp}"
    batch_dir.mkdir(parents=True, exist_ok=True)

    if status_cb:
        status_cb("Lade Regeln...")
    rules = load_rules(rules_file)

    if status_cb:
        status_cb("Starte Scan...")
    pages = run_scan(
        naps2_path=naps2_path,
        profile=profile,
        batch_dir=batch_dir,
        image_format=image_format,
        source=source,
        verbose=verbose,
    )

    if status_cb:
        status_cb(f"Scan abgeschlossen: {len(pages)} Seite(n)")

    if auto_orient:
        if status_cb:
            status_cb("Starte automatische Ausrichtung...")
        auto_orient_pages(pages=pages, tesseract_path=tesseract_path, verbose=verbose, status_cb=status_cb)

    results = sort_pages(
        pages=pages,
        target_root=target_root,
        tesseract_path=tesseract_path,
        lang=lang,
        rules=rules,
        default_category=default_category,
        verbose=verbose,
        status_cb=status_cb,
    )

    report = write_report(batch_dir=batch_dir, results=results)
    return results, report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scan + OCR + automatische Ordner-Sortierung (mit GUI)")

    parser.add_argument("--gui", action="store_true", help="GUI starten (Standard, wenn kein --profile gesetzt ist)")
    parser.add_argument("--profile", help="NAPS2-Profilname für den Scanner")
    parser.add_argument("--scan-root", default="scans", help="Basisordner für Rohscans")
    parser.add_argument("--target-root", default="sorted", help="Basisordner für sortierte Dateien")
    parser.add_argument("--naps2-path", default="naps2.console", help="Pfad zu NAPS2 Console")
    parser.add_argument("--tesseract-path", default="tesseract", help="Pfad zu Tesseract")
    parser.add_argument("--lang", default="deu+eng", help="OCR-Sprache(n), z. B. deu+eng")
    parser.add_argument("--image-format", default="png", choices=["png", "jpg", "tif", "tiff"])
    parser.add_argument("--rules-file", type=Path, help="JSON-Datei mit eigenen Regeln")
    parser.add_argument("--default-category", default="Unsortiert")
    parser.add_argument("--source", default="adf", choices=["adf", "glasplatte"], help="Einzug oder Glasplatte")
    parser.add_argument("--auto-orient", action="store_true", help="Automatische Ausrichtung aktivieren")
    parser.add_argument("--verbose", action="store_true")

    return parser.parse_args()


class ScanSortGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Scan Sort - NAPS2 + Tesseract")
        self.events: Queue[tuple[str, str]] = Queue()
        self.worker_running = False

        self.profile_var = tk.StringVar(value="Brother")
        self.scan_root_var = tk.StringVar(value="scans")
        self.target_root_var = tk.StringVar(value="sorted")
        self.naps2_var = tk.StringVar(value="naps2.console")
        self.tesseract_var = tk.StringVar(value="tesseract")
        self.lang_var = tk.StringVar(value="deu+eng")
        self.source_var = tk.StringVar(value="adf")
        self.rules_var = tk.StringVar(value="")
        self.auto_orient_var = tk.BooleanVar(value=True)

        self._build_ui()
        self._poll_events()

    def _build_ui(self) -> None:
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)

        fields = [
            ("NAPS2 Profil", self.profile_var),
            ("Scan Root", self.scan_root_var),
            ("Target Root", self.target_root_var),
            ("NAPS2 Pfad", self.naps2_var),
            ("Tesseract Pfad", self.tesseract_var),
            ("Sprache", self.lang_var),
            ("Rules JSON", self.rules_var),
        ]

        for r, (label, var) in enumerate(fields):
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=2)
            ttk.Entry(frm, textvariable=var).grid(row=r, column=1, sticky="ew", pady=2)

        row = len(fields)
        ttk.Label(frm, text="Quelle").grid(row=row, column=0, sticky="w", pady=2)
        src_combo = ttk.Combobox(frm, textvariable=self.source_var, values=["adf", "glasplatte"], state="readonly")
        src_combo.grid(row=row, column=1, sticky="ew", pady=2)

        row += 1
        ttk.Checkbutton(frm, text="Automatische Ausrichtung", variable=self.auto_orient_var).grid(row=row, column=1, sticky="w")

        row += 1
        self.start_btn = ttk.Button(frm, text="Scan starten", command=self.start_scan)
        self.start_btn.grid(row=row, column=1, sticky="w", pady=6)

        row += 1
        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=row, column=0, columnspan=2, sticky="ew", pady=4)

        row += 1
        ttk.Label(frm, text="Status").grid(row=row, column=0, sticky="w")
        row += 1
        self.status_text = tk.Text(frm, height=14, wrap="word")
        self.status_text.grid(row=row, column=0, columnspan=2, sticky="nsew")
        frm.rowconfigure(row, weight=1)

    def log(self, message: str) -> None:
        self.status_text.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.status_text.see("end")

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "log":
                    self.log(payload)
                elif event == "done":
                    self.worker_running = False
                    self.progress.stop()
                    self.start_btn.config(state="normal")
                    self.log(payload)
                elif event == "error":
                    self.worker_running = False
                    self.progress.stop()
                    self.start_btn.config(state="normal")
                    self.log(f"Fehler: {payload}")
        except Empty:
            pass

        self.root.after(150, self._poll_events)

    def _emit_log(self, message: str) -> None:
        self.events.put(("log", message))

    def start_scan(self) -> None:
        if self.worker_running:
            return

        self.worker_running = True
        self.start_btn.config(state="disabled")
        self.progress.start(8)
        self.log("Starte Workflow...")

        def worker() -> None:
            try:
                rules_file = Path(self.rules_var.get()) if self.rules_var.get().strip() else None
                results, report = process_scan_and_sort(
                    profile=self.profile_var.get().strip(),
                    scan_root=Path(self.scan_root_var.get().strip()),
                    target_root=Path(self.target_root_var.get().strip()),
                    naps2_path=self.naps2_var.get().strip(),
                    tesseract_path=self.tesseract_var.get().strip(),
                    lang=self.lang_var.get().strip(),
                    image_format="png",
                    rules_file=rules_file,
                    default_category="Unsortiert",
                    source=self.source_var.get().strip(),
                    auto_orient=self.auto_orient_var.get(),
                    verbose=False,
                    status_cb=self._emit_log,
                )
                self.events.put(("done", f"Fertig: {len(results)} Seite(n) verarbeitet. Report: {report}"))
            except Exception as exc:  # noqa: BLE001
                self.events.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()


def run_gui() -> int:
    if tk is None or ttk is None:
        print("Tkinter ist nicht verfügbar. Bitte per CLI mit --profile nutzen.", file=sys.stderr)
        return 1
    root = tk.Tk()
    app = ScanSortGUI(root)
    _ = app
    root.mainloop()
    return 0


def run_cli(args: argparse.Namespace) -> int:
    if not args.profile:
        print("Für CLI muss --profile gesetzt sein oder --gui genutzt werden.", file=sys.stderr)
        return 1

    try:
        results, report = process_scan_and_sort(
            profile=args.profile,
            scan_root=Path(args.scan_root),
            target_root=Path(args.target_root),
            naps2_path=args.naps2_path,
            tesseract_path=args.tesseract_path,
            lang=args.lang,
            image_format=args.image_format,
            rules_file=args.rules_file,
            default_category=args.default_category,
            source=args.source,
            auto_orient=args.auto_orient,
            verbose=args.verbose,
            status_cb=lambda msg: print(msg),
        )
        print(f"Fertig: {len(results)} Seite(n) verarbeitet.")
        print(f"Report: {report}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Fehler: {exc}", file=sys.stderr)
        return 1


def main() -> int:
    args = parse_args()
    if args.gui or not args.profile:
        return run_gui()
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
