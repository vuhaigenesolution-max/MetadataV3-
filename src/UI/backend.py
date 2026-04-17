import os
import sys

_SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from Logic.CombieFile import process_combine_files
from Logic.Primer_T7 import process_primer_t7


def run_backend(source_mode, source_path, template_path, output_path, sheet_template="Sample", progress_callback=None):
    if source_mode == "file":
        folder = os.path.dirname(source_path)
        files, stats = process_combine_files(
            folder=folder,
            template_path=template_path,
            output_folder=output_path,
            sheet_template=sheet_template,
            progress_callback=progress_callback,
            filter_file=os.path.basename(source_path),
        )
    else:
        files, stats = process_combine_files(
            folder=source_path,
            template_path=template_path,
            output_folder=output_path,
            sheet_template=sheet_template,
            progress_callback=progress_callback,
        )

    # T7 primer check tự động trên các file vừa xuất
    t7_result = None
    if files:
        try:
            t7_result = process_primer_t7(file_paths=files, output_folder=output_path)
        except Exception as e:
            t7_result = {"message": "error", "error": str(e),
                         "total": 0, "duplicates": 0, "valid": 0, "output_path": None}

    return {
        "files":              files,
        "collisions":         stats["total_collisions"],
        "collision_per_file": stats["collision_per_file"],
        "error_files":        stats["error_files"],
        "t7":                 t7_result,
    }
