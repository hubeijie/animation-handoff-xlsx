#!/usr/bin/env python3
"""灰底软抠图（与四角背景色接近的像素渐隐为透明），并导出多枚网页背景用变体。"""

from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageOps

ASSETS = Path(__file__).resolve().parent.parent / "static" / "assets"


def matte_uniform_gray(
    rgba: Image.Image,
    bg: tuple[float, float, float] = (210, 210, 210),
    t0: float = 6.0,
    t1: float = 40.0,
) -> Image.Image:
    a = np.asarray(rgba.convert("RGBA"), dtype=np.float32)
    rgb = a[:, :, :3]
    d = np.linalg.norm(rgb - np.array(bg, dtype=np.float32), axis=2)
    alpha = np.clip((d - t0) / (t1 - t0), 0.0, 1.0) * 255.0
    out = a.copy()
    out[:, :, 3] = np.minimum(out[:, :, 3], alpha)
    return Image.fromarray(out.astype(np.uint8))


def trim_transparent(im: Image.Image, pad: int = 4) -> Image.Image:
    bbox = im.getbbox()
    if not bbox:
        return im
    x0, y0, x1, y1 = bbox
    x0 = max(0, x0 - pad)
    y0 = max(0, y0 - pad)
    x1 = min(im.width, x1 + pad)
    y1 = min(im.height, y1 + pad)
    return im.crop((x0, y0, x1, y1))


def rotate_expand(im: Image.Image, deg: float) -> Image.Image:
    return im.rotate(
        deg,
        resample=Image.Resampling.BICUBIC,
        expand=True,
        fillcolor=(0, 0, 0, 0),
    )


def main() -> None:
    raw_path = ASSETS / "xiaobai_raw.png"
    rembg_path = ASSETS / "xiaobai_cutout.png"
    if not raw_path.is_file():
        raise SystemExit(f"缺少源图: {raw_path}")

    if rembg_path.is_file():
        base = trim_transparent(Image.open(rembg_path).convert("RGBA"))
        print("使用 rembg 抠图:", rembg_path)
    else:
        base = trim_transparent(matte_uniform_gray(Image.open(raw_path)))
        print("使用灰底软抠图（可安装 rembg 后生成 xiaobai_cutout.png 再运行本脚本）")

    base.save(ASSETS / "xiaobai_pose_neutral.png", optimize=True)

    ImageOps.mirror(base).save(ASSETS / "xiaobai_pose_mirror.png", optimize=True)
    rotate_expand(base, 10).save(ASSETS / "xiaobai_pose_tilt_r.png", optimize=True)
    rotate_expand(base, -12).save(ASSETS / "xiaobai_pose_tilt_l.png", optimize=True)
    ImageOps.mirror(rotate_expand(base, 8)).save(
        ASSETS / "xiaobai_pose_tilt_mirror.png", optimize=True
    )

    base.save(ASSETS / "xiaobai.png", optimize=True)
    print("已写入:", ASSETS)


if __name__ == "__main__":
    main()
