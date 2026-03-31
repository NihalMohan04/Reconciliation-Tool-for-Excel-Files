import argparse
import json
import os
import sys
from pathlib import Path
from tkinter import filedialog, messagebox
from collections.abc import Callable

import customtkinter as ctk

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from excel_tool.processor import reconcile_workbooks


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run reconciliation via GUI (default) or CLI mode."
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Run in CLI mode using Source/Target/Recon under --root",
    )
    parser.add_argument(
        "--root",
        default=str(PROJECT_ROOT),
        help="Project root for CLI mode containing Source, Target, and Recon folders",
    )
    return parser.parse_args()


def _excel_files_by_name(folder_path: Path) -> dict[str, Path]:
    supported_suffixes = {".xlsx", ".xlsm"}
    files = [
        file_path
        for file_path in folder_path.iterdir()
        if file_path.is_file() and file_path.suffix.lower() in supported_suffixes
    ]
    return {file_path.name.lower(): file_path for file_path in files}


def _common_excel_names(source_dir: Path, target_dir: Path) -> list[str]:
    if not source_dir.exists():
        raise FileNotFoundError(f"Source folder not found: {source_dir}")
    if not target_dir.exists():
        raise FileNotFoundError(f"Target folder not found: {target_dir}")

    source_files = _excel_files_by_name(source_dir)
    target_files = _excel_files_by_name(target_dir)
    return sorted(set(source_files.keys()) & set(target_files.keys()))


def _run_reconciliation(
    source_dir: Path,
    target_dir: Path,
    recon_dir: Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict:
    source_files = _excel_files_by_name(source_dir)
    target_files = _excel_files_by_name(target_dir)
    common_names = sorted(set(source_files.keys()) & set(target_files.keys()))

    if not common_names:
        raise FileNotFoundError(
            "No same-named Excel files found between Source and Target folders."
        )

    summaries = []
    total_files = len(common_names)
    for file_index, common_name in enumerate(common_names, start=1):
        source_file = source_files[common_name]
        target_file = target_files[common_name]
        output_file = recon_dir / f"Recon_{source_file.stem}.xlsx"

        if progress_callback is not None:
            progress_callback(file_index, total_files, source_file.name)

        summary = reconcile_workbooks(
            source_file=str(source_file),
            target_file=str(target_file),
            output_file=str(output_file),
        )
        summaries.append(summary)

    total_matched = sum(item["matched_rows"] for item in summaries)
    total_not_matched = sum(item["not_matched_rows"] for item in summaries)
    total_processed = sum(item["processed_rows"] for item in summaries)

    return {
        "processed_files": len(summaries),
        "total_matched_rows": total_matched,
        "total_not_matched_rows": total_not_matched,
        "total_processed_rows": total_processed,
        "files": summaries,
    }


def _default_recon_dir(source_dir: Path, target_dir: Path) -> Path:
    if source_dir.parent == target_dir.parent:
        return source_dir.parent / "Recon"
    return PROJECT_ROOT / "Recon"


def run_cli(project_root: Path) -> int:
    source_dir = project_root / "Source"
    target_dir = project_root / "Target"
    recon_dir = project_root / "Recon"

    result = _run_reconciliation(source_dir=source_dir, target_dir=target_dir, recon_dir=recon_dir)
    print(json.dumps(result, indent=2))
    return 0


class ReconApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Recon Tool for Viatris")
        self.geometry("640x360")
        self.resizable(False, False)

        self.source_var = ctk.StringVar(value="")
        self.target_var = ctk.StringVar(value="")
        self.recon_var = ctk.StringVar(value="")

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ctk.CTkFrame(self)
        outer.pack(fill="both", expand=True, padx=16, pady=16)

        title_label = ctk.CTkLabel(
            outer,
            text="Recon Tool for Viatris",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_label.pack(anchor="w", pady=(8, 14), padx=12)

        source_row = ctk.CTkFrame(outer, fg_color="transparent")
        source_row.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(source_row, text="Source Folder", width=120, anchor="w").pack(side="left")
        ctk.CTkEntry(source_row, textvariable=self.source_var).pack(side="left", fill="x", expand=True, padx=8)
        ctk.CTkButton(source_row, text="Browse", width=90, command=self._pick_source).pack(side="left")

        target_row = ctk.CTkFrame(outer, fg_color="transparent")
        target_row.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(target_row, text="Target Folder", width=120, anchor="w").pack(side="left")
        ctk.CTkEntry(target_row, textvariable=self.target_var).pack(side="left", fill="x", expand=True, padx=8)
        ctk.CTkButton(target_row, text="Browse", width=90, command=self._pick_target).pack(side="left")

        recon_row = ctk.CTkFrame(outer, fg_color="transparent")
        recon_row.pack(fill="x", padx=12, pady=(10, 2))
        ctk.CTkLabel(recon_row, text="Output Folder", width=120, anchor="w").pack(side="left")
        recon_entry = ctk.CTkEntry(recon_row, textvariable=self.recon_var, state="readonly")
        recon_entry.pack(side="left", fill="x", expand=True, padx=8)

        self.status_label = ctk.CTkLabel(outer, text="Select Source and Target folders.", anchor="w")
        self.status_label.pack(fill="x", padx=12, pady=(14, 8))

        action_row = ctk.CTkFrame(outer, fg_color="transparent")
        action_row.pack(fill="x", padx=12, pady=(8, 12))
        self.run_button = ctk.CTkButton(action_row, text="Run Reconciliation", state="disabled", command=self._run)
        self.run_button.pack(side="left")
        ctk.CTkButton(action_row, text="Exit", width=90, command=self.destroy).pack(side="right")

    def _pick_source(self) -> None:
        selected = filedialog.askdirectory(title="Select Source folder")
        if selected:
            self.source_var.set(selected)
            self._refresh_validation()

    def _pick_target(self) -> None:
        selected = filedialog.askdirectory(title="Select Target folder")
        if selected:
            self.target_var.set(selected)
            self._refresh_validation()

    def _refresh_validation(self) -> None:
        source_text = self.source_var.get().strip()
        target_text = self.target_var.get().strip()

        if not source_text or not target_text:
            self.recon_var.set("")
            self.run_button.configure(state="disabled")
            self.status_label.configure(text="Select Source and Target folders.")
            return

        source_dir = Path(source_text)
        target_dir = Path(target_text)
        recon_dir = _default_recon_dir(source_dir, target_dir)
        self.recon_var.set(str(recon_dir))

        try:
            common_names = _common_excel_names(source_dir, target_dir)
        except FileNotFoundError as error:
            self.run_button.configure(state="disabled")
            self.status_label.configure(text=str(error))
            return

        if not common_names:
            self.run_button.configure(state="disabled")
            self.status_label.configure(
                text="No same-named .xlsx/.xlsm files found in Source and Target."
            )
            return

        self.run_button.configure(state="normal")
        self.status_label.configure(text=f"Ready: {len(common_names)} matching file pair(s) found.")

    def _run(self) -> None:
        source_dir = Path(self.source_var.get().strip())
        target_dir = Path(self.target_var.get().strip())
        recon_dir = Path(self.recon_var.get().strip())

        self.run_button.configure(state="disabled")
        self.status_label.configure(text="Running reconciliation...")
        self.update_idletasks()

        def _on_progress(current: int, total: int, filename: str) -> None:
            self.status_label.configure(text=f"Processing {current}/{total}: {filename}")
            self.update_idletasks()

        try:
            result = _run_reconciliation(
                source_dir=source_dir,
                target_dir=target_dir,
                recon_dir=recon_dir,
                progress_callback=_on_progress,
            )
        except Exception as error:
            self.status_label.configure(text="Reconciliation failed.")
            messagebox.showerror("Recon Tool by nimo", str(error), parent=self)
            self._refresh_validation()
            return

        summary = (
            f"Processed files: {result['processed_files']}\n"
            f"Total matched rows: {result['total_matched_rows']}\n"
            f"Total not matched rows: {result['total_not_matched_rows']}\n"
            f"Total processed rows: {result['total_processed_rows']}\n\n"
            f"Output folder:\n{recon_dir}"
        )
        self.status_label.configure(text="Completed successfully.")
        messagebox.showinfo("Recon Tool by nimo", summary, parent=self)
        try:
            os.startfile(recon_dir)
        except OSError as error:
            messagebox.showwarning(
                "Recon Tool by nimo",
                f"Reconciliation completed, but could not open output folder.\n\n{error}",
                parent=self,
            )
        self._refresh_validation()


def run_gui() -> int:
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = ReconApp()
    app.mainloop()
    return 0


def main() -> int:
    args = parse_args()
    if args.cli:
        return run_cli(Path(args.root))
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
