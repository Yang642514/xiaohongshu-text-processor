import os
import zipfile
from typing import Optional


def make_zip(input_file: str, zip_output_dir: str,
             images_dir: Optional[str] = None,
             zip_name_suffix: Optional[str] = None) -> str:
    os.makedirs(zip_output_dir, exist_ok=True)
    base = os.path.basename(input_file)
    name, _ = os.path.splitext(base)
    # 支持配图版命名：<主题>-配图版.zip
    if zip_name_suffix:
        name = f"{name}-{zip_name_suffix}"
    zip_path = os.path.join(zip_output_dir, f"{name}.zip")

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Excel 位于压缩包根目录
        zf.write(input_file, arcname=os.path.basename(input_file))
        # 图片与 Excel 同级（根目录），不再放入 images/ 子目录
        if images_dir and os.path.isdir(images_dir):
            for root, _, files in os.walk(images_dir):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    zf.write(fpath, arcname=os.path.basename(fpath))
    return zip_path