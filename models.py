from dataclasses import dataclass
from typing import List, Optional


@dataclass
class FileTask:
    task_type: str
    file_paths: List[str]
    output_dir: str
    split_column: Optional[str] = None