import os
import zipfile


def make_zip(input_file: str, zip_output_dir: str) -> str:
    os.makedirs(zip_output_dir, exist_ok=True)
    base = os.path.basename(input_file)
    name, _ = os.path.splitext(base)
    zip_path = os.path.join(zip_output_dir, f"{name}.zip")

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(input_file, arcname=os.path.basename(input_file))
        # 预留：将图片或其他资源加入压缩包
    return zip_path