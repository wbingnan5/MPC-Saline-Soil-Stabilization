# -*- coding: utf-8 -*-
"""Generate publication-ready Fig. 14 from the original electrochemical data.

The script is intentionally non-destructive:
  * it only reads the 15 vendor .bin files and the photographs embedded in
    the adjacent corrosion PowerPoint file;
  * it never edits or deletes an input file;
  * existing generated outputs are protected unless
    --overwrite-generated is supplied explicitly.

Default output (in the directory containing this script):
  Fig14a.png--Fig14f.png as six separate 600-dpi RGB panels.
  Fig14_legend.png as the matching shared legend strip.

Only PNG panel files are generated. The script does not create a composite
figure, vector/TIFF files, source-data tables, or a QA report.

Usage:
    python plot_Fig14.py
    python plot_Fig14.py --output-dir "D:/another/folder"
    python plot_Fig14.py --overwrite-generated
    python plot_Fig14.py --legend-only

The exposed steel area is 1.00 cm^2 (1.0 cm x 1.0 cm).
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import math
import platform
import struct
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable
from zipfile import ZipFile

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
from matplotlib.lines import Line2D
from PIL import Image, ImageOps, __version__ as PILLOW_VERSION


# ---------------------------------------------------------------------------
# Figure contract
# ---------------------------------------------------------------------------
# Core conclusion:
# MPC stabilization suppresses steel corrosion in saline silty clay by
# increasing interfacial polarization resistance and low-frequency impedance
# and by restricting electrolyte transport; the effect persists at 28 d.
#
# Evidence map:
# a, qualitative surface condition; b, OCP evolution; c, polarization kinetics;
# d, complex impedance; e, barrier magnitude; f, capacitive/interfacial response.
#
# Archetype: image plate + quantitative electrochemical grid.
# Target: Arabian Journal for Science and Engineering; six separate 84-mm-wide
# RGB panels and one shared legend strip exported only as 600-dpi PNG files.


WORKING_AREA_CM2 = 1.0 * 1.0
OCP_DT_S = 0.1
LOW_FREQUENCY_REFERENCE_HZ = 0.1
NYQUIST_MIN_FREQUENCY_HZ = 0.1
FINAL_WIDTH_MM = 174.0
FINAL_HEIGHT_MM = 190.0
MM_PER_INCH = 25.4
PNG_DPI = 600
PANEL_WIDTH_MM = 84.0
PANEL_HEIGHT_MM = 66.0
LEGEND_WIDTH_MM = 174.0
LEGEND_HEIGHT_MM = 10.0


@dataclass(frozen=True)
class GroupStyle:
    key: str
    label: str
    color: str
    linestyle: object
    marker: str
    filled: bool
    order: int
    subdir: str
    prefix: str


GROUPS = (
    GroupStyle(
        "BC", "Blank control", "#5A5A5A", (0, (1.2, 1.8)), "o", True, 0,
        "空白", "空白",
    ),
    GroupStyle(
        "SC7", "SC–7 d", "#B45A3A", (0, (5.0, 2.0)), "^", False, 1,
        "7d", "盐碱土-7d",
    ),
    GroupStyle(
        "SC28", "SC–28 d", "#A33F2B", "-", "^", True, 2,
        "28d", "盐碱土-28d",
    ),
    GroupStyle(
        "MPC7", "MPC–SC–7 d", "#3478AD", (0, (2.0, 1.5)), "s", False, 3,
        "7d", "MPC改-7d",
    ),
    GroupStyle(
        "MPC28", "MPC–SC–28 d", "#145A96", "-", "s", True, 4,
        "28d", "MPC改-28d",
    ),
)


# The 28-d MPC impedance filename is the only filename that omits "28d".
FILE_MAP = {
    "BC": {
        "OCPT": Path("空白") / "空白-OCPT.bin",
        "TAFEL": Path("空白") / "空白-TAFEL.bin",
        "IMP": Path("空白") / "空白-IMP.bin",
    },
    "SC7": {
        "OCPT": Path("7d") / "盐碱土-7d-OCPT.bin",
        "TAFEL": Path("7d") / "盐碱土-7d-TAFEL.bin",
        "IMP": Path("7d") / "盐碱土-7d-IMP.bin",
    },
    "MPC7": {
        "OCPT": Path("7d") / "MPC改-7d-OCPT.bin",
        "TAFEL": Path("7d") / "MPC改-7d-TAFEL.bin",
        "IMP": Path("7d") / "MPC改-7d-IMP.bin",
    },
    "SC28": {
        "OCPT": Path("28d") / "盐碱土-28d-OCPT.bin",
        "TAFEL": Path("28d") / "盐碱土-28d-TAFEL.bin",
        "IMP": Path("28d") / "盐碱土-28d-IMP.bin",
    },
    "MPC28": {
        "OCPT": Path("28d") / "MPC改-28d-OCPT.bin",
        "TAFEL": Path("28d") / "MPC改-28d-TAFEL.bin",
        "IMP": Path("28d") / "MPC改-IMP.bin",
    },
}


# The crop rectangles reproduce the crop settings stored in the source PPTX.
# Values follow PowerPoint's srcRect order: left, top, right, bottom, each in
# one-hundred-thousandths. No selective brightness/contrast processing is used.
PHOTO_SPECS = {
    "SC7": {
        "member": "ppt/media/image3.png",
        "crop": (21851, 19259, 31852, 45278),
        "rotate": 0,
        "label": "SC–7 d",
    },
    "SC28": {
        "member": "ppt/media/image4.png",
        "crop": (27605, 9573, 25735, 55469),
        "rotate": 0,
        "label": "SC–28 d",
    },
    "MPC7": {
        "member": "ppt/media/image1.png",
        "crop": (36451, 15503, 41221, 67088),
        "rotate": 0,
        "label": "MPC–SC–7 d",
    },
    "MPC28": {
        "member": "ppt/media/image2.png",
        "crop": (28886, 8889, 37654, 65933),
        "rotate": 180,
        "label": "MPC–SC–28 d",
    },
}


def configure_matplotlib() -> None:
    """Apply restrained, journal-scale typography and line styling."""
    # Fail explicitly instead of silently substituting another font.
    for style, weight in (
        ("normal", "normal"),
        ("italic", "normal"),
        ("normal", "bold"),
        ("italic", "bold"),
    ):
        font_manager.findfont(
            font_manager.FontProperties(
                family="Times New Roman", style=style, weight=weight
            ),
            fallback_to_default=False,
        )

    mpl.rcParams.update(
        {
            "font.family": "serif",
            "font.serif": ["Times New Roman"],
            "text.usetex": False,
            "mathtext.fontset": "custom",
            "mathtext.rm": "Times New Roman",
            "mathtext.it": "Times New Roman:italic",
            "mathtext.bf": "Times New Roman:bold",
            "mathtext.bfit": "Times New Roman:bold:italic",
            "mathtext.sf": "Times New Roman",
            "mathtext.tt": "Times New Roman",
            "mathtext.cal": "Times New Roman:italic",
            "mathtext.fallback": "stix",
            "axes.formatter.use_mathtext": True,
            "font.size": 8.0,
            "axes.labelsize": 8.5,
            "axes.titlesize": 8.0,
            "xtick.labelsize": 8.0,
            "ytick.labelsize": 8.0,
            "legend.fontsize": 8.0,
            "axes.linewidth": 0.8,
            "lines.linewidth": 1.25,
            "xtick.major.width": 0.75,
            "ytick.major.width": 0.75,
            "xtick.minor.width": 0.55,
            "ytick.minor.width": 0.55,
            "xtick.major.size": 3.0,
            "ytick.major.size": 3.0,
            "xtick.minor.size": 1.8,
            "ytick.minor.size": 1.8,
            "xtick.direction": "out",
            "ytick.direction": "out",
            "legend.frameon": False,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "svg.fonttype": "none",
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
            "savefig.facecolor": "white",
            "figure.facecolor": "white",
            "axes.facecolor": "white",
        }
    )


def read_f32(raw: bytes, offset: int) -> float:
    return float(struct.unpack_from("<f", raw, offset)[0])


def identify_method(raw: bytes, path: Path) -> str:
    if raw[:4] != b"\x80\xf2\x1d\x00":
        raise ValueError(f"Unexpected vendor magic in {path}")
    nchar = int(struct.unpack_from("<I", raw, 4)[0])
    method = raw[8 : 8 + nchar].decode("ascii")
    if method not in {"OCPT", "TAFEL", "IMP"}:
        raise ValueError(f"Unsupported method {method!r} in {path}")
    return method


def parse_ocp(path: Path) -> dict[str, np.ndarray | float]:
    raw = path.read_bytes()
    if identify_method(raw, path) != "OCPT":
        raise ValueError(f"Expected OCPT data: {path}")
    if (len(raw) - 1694) % 4:
        raise ValueError(f"Unexpected OCPT payload length: {path}")
    potential = np.frombuffer(raw, dtype="<f4", offset=1694).astype(float)
    if potential.size == 0 or not np.isclose(potential[0], 0.0):
        raise ValueError(f"OCPT sentinel missing: {path}")
    potential = potential[1:]
    if potential.size < 100 or not np.all(np.isfinite(potential)):
        raise ValueError(f"Invalid OCPT payload: {path}")
    if np.any((potential < -2.0) | (potential > 2.0)):
        raise ValueError(f"OCPT values outside the expected voltage range: {path}")
    time_s = np.arange(potential.size, dtype=float) * OCP_DT_S
    tail_n = min(int(round(10.0 / OCP_DT_S)), potential.size)
    tail = potential[-tail_n:]
    return {
        "time_s": time_s,
        "potential_V": potential,
        "terminal_mean_V": float(np.mean(tail)),
        "terminal_temporal_sd_V": float(np.std(tail, ddof=0)),
        "duration_s": float(time_s[-1]),
    }


def zero_crossing(x: np.ndarray, y: np.ndarray) -> tuple[float, int]:
    crossings = np.flatnonzero(np.signbit(y[:-1]) != np.signbit(y[1:]))
    if crossings.size != 1:
        raise ValueError(f"Expected one current sign crossing, found {crossings.size}")
    idx = int(crossings[0])
    xzero = x[idx] + (0.0 - y[idx]) * (x[idx + 1] - x[idx]) / (y[idx + 1] - y[idx])
    return float(xzero), idx


def linear_fit_with_r2(x: np.ndarray, y: np.ndarray) -> tuple[float, float, float]:
    slope, intercept = np.polyfit(x, y, 1)
    predicted = slope * x + intercept
    residual = float(np.sum((y - predicted) ** 2))
    total = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1.0 - residual / total if total > 0 else float("nan")
    return float(slope), float(intercept), float(r2)


def parse_tafel(path: Path) -> dict[str, np.ndarray | float]:
    raw = path.read_bytes()
    if identify_method(raw, path) != "TAFEL":
        raise ValueError(f"Expected TAFEL data: {path}")
    if (len(raw) - 1680) % 4:
        raise ValueError(f"Unexpected TAFEL payload length: {path}")

    current_A = np.frombuffer(raw, dtype="<f4", offset=1680).astype(float)
    start_V = read_f32(raw, 1088)
    end_V = read_f32(raw, 1092)
    step_V = abs(read_f32(raw, 1112))
    if step_V <= 0 or current_A.size < 100:
        raise ValueError(f"Invalid TAFEL acquisition settings: {path}")

    expected_n = int(round(abs(end_V - start_V) / step_V))
    if expected_n != current_A.size:
        raise ValueError(
            f"TAFEL point count does not match the recorded sweep-potential "
            f"range in {path}: expected {expected_n}, found {current_A.size}"
        )

    direction = math.copysign(1.0, end_V - start_V)
    potential_V = start_V + np.arange(current_A.size, dtype=float) * direction * step_V
    ezero_V, crossing_index = zero_crossing(potential_V, current_A)
    eta_V = potential_V - ezero_V
    j_A_cm2 = current_A / WORKING_AREA_CM2

    # Apparent local polarization resistance: inverse slope of j versus E
    # within +/-10 mV around the measured zero-current crossing.
    rp_mask = np.abs(eta_V) <= 0.010
    if np.count_nonzero(rp_mask) < 8:
        raise ValueError(f"Insufficient +/-10 mV data for local Rp: {path}")
    dj_dE, _, rp_r2 = linear_fit_with_r2(potential_V[rp_mask], j_A_cm2[rp_mask])
    rp_app_ohm_cm2 = abs(1.0 / dj_dE)

    # Identical 30--100 mV Tafel windows are used for all curves. The result is
    # deliberately reported as an apparent fitted current density.
    safe_log_j = np.log10(np.clip(np.abs(j_A_cm2), 1e-20, None))
    anodic = (eta_V >= 0.030) & (eta_V <= 0.100) & (j_A_cm2 > 0)
    cathodic = (eta_V <= -0.030) & (eta_V >= -0.100) & (j_A_cm2 < 0)
    if np.count_nonzero(anodic) < 10 or np.count_nonzero(cathodic) < 10:
        raise ValueError(f"Insufficient common Tafel window: {path}")
    a_slope, a_intercept, a_r2 = linear_fit_with_r2(
        potential_V[anodic], safe_log_j[anodic]
    )
    c_slope, c_intercept, c_r2 = linear_fit_with_r2(
        potential_V[cathodic], safe_log_j[cathodic]
    )
    fit_ecorr_V = (c_intercept - a_intercept) / (a_slope - c_slope)
    log_jcorr = a_slope * fit_ecorr_V + a_intercept
    jcorr_app_A_cm2 = 10.0 ** log_jcorr

    return {
        "potential_V": potential_V,
        "eta_V": eta_V,
        "current_A": current_A,
        "j_A_cm2": j_A_cm2,
        "zero_current_potential_V": ezero_V,
        "zero_crossing_index": float(crossing_index),
        "rp_app_ohm_cm2": float(rp_app_ohm_cm2),
        "rp_fit_r2": float(rp_r2),
        "jcorr_app_A_cm2": float(jcorr_app_A_cm2),
        "tafel_fit_ecorr_V": float(fit_ecorr_V),
        "tafel_anodic_slope_dec_V": float(a_slope),
        "tafel_cathodic_slope_dec_V": float(c_slope),
        "tafel_anodic_r2": float(a_r2),
        "tafel_cathodic_r2": float(c_r2),
        "start_V": float(start_V),
        "end_V": float(end_V),
        "step_V": float(step_V),
    }


def parse_eis(path: Path) -> dict[str, np.ndarray | float]:
    raw = path.read_bytes()
    if identify_method(raw, path) != "IMP":
        raise ValueError(f"Expected IMP data: {path}")
    if (len(raw) - 1682) % 16:
        raise ValueError(f"Unexpected IMP payload length: {path}")
    payload = np.frombuffer(raw, dtype="<f4", offset=1682).reshape(-1, 4).astype(float)
    set_frequency_Hz = payload[:, 0]
    frequency_Hz = payload[:, 1]
    modulus_ohm = payload[:, 2]
    phase_deg = payload[:, 3]

    if payload.shape != (85, 4):
        raise ValueError(f"Expected 85 EIS points in {path}, found {payload.shape[0]}")
    if not np.allclose(set_frequency_Hz, frequency_Hz, rtol=1e-6, atol=1e-9):
        raise ValueError(f"Set and measured frequencies differ unexpectedly: {path}")
    if not np.all(np.diff(frequency_Hz) < 0):
        raise ValueError(f"EIS frequencies are not strictly descending: {path}")
    if np.any(modulus_ohm <= 0):
        raise ValueError(f"Non-positive EIS modulus: {path}")

    z_area = modulus_ohm * np.exp(1j * np.deg2rad(phase_deg)) * WORKING_AREA_CM2
    reference_index = int(np.argmin(np.abs(frequency_Hz - LOW_FREQUENCY_REFERENCE_HZ)))
    return {
        "frequency_Hz": frequency_Hz,
        "modulus_raw_ohm": modulus_ohm,
        "modulus_area_ohm_cm2": modulus_ohm * WORKING_AREA_CM2,
        "phase_deg": phase_deg,
        "zreal_area_ohm_cm2": np.real(z_area),
        "minus_zimag_area_ohm_cm2": -np.imag(z_area),
        "reference_index": float(reference_index),
        "modulus_at_0p1_ohm_cm2": float(modulus_ohm[reference_index] * WORKING_AREA_CM2),
        "phase_at_0p1_deg": float(phase_deg[reference_index]),
        "low_frequency_nonstationary_points": float(
            np.count_nonzero((frequency_Hz < 0.1) & (phase_deg < -90.0))
        ),
    }


def load_all_data(base_dir: Path) -> dict[str, dict[str, dict[str, np.ndarray | float]]]:
    data: dict[str, dict[str, dict[str, np.ndarray | float]]] = {}
    for style in GROUPS:
        data[style.key] = {}
        for method, rel_path in FILE_MAP[style.key].items():
            source = base_dir / rel_path
            if not source.is_file():
                raise FileNotFoundError(source)
            if method == "OCPT":
                data[style.key][method] = parse_ocp(source)
            elif method == "TAFEL":
                data[style.key][method] = parse_tafel(source)
            elif method == "IMP":
                data[style.key][method] = parse_eis(source)
    return data


def powerpoint_crop(image: Image.Image, crop: tuple[int, int, int, int]) -> Image.Image:
    left, top, right, bottom = crop
    width, height = image.size
    box = (
        round(width * left / 100000),
        round(height * top / 100000),
        round(width * (1.0 - right / 100000)),
        round(height * (1.0 - bottom / 100000)),
    )
    if not (0 <= box[0] < box[2] <= width and 0 <= box[1] < box[3] <= height):
        raise ValueError(f"Invalid PowerPoint crop box {box} for image size {image.size}")
    return image.crop(box)


def load_photos(pptx_path: Path, max_side_px: int = 1400) -> dict[str, Image.Image]:
    if not pptx_path.is_file():
        raise FileNotFoundError(pptx_path)
    photos: dict[str, Image.Image] = {}
    with ZipFile(pptx_path) as archive:
        for key, spec in PHOTO_SPECS.items():
            raw = archive.read(str(spec["member"]))
            image = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
            image = powerpoint_crop(image, spec["crop"])
            rotation = int(spec["rotate"])
            if rotation:
                image = image.rotate(rotation, expand=True)
            # Downsampling only controls vector-file size; no color or local image
            # enhancement is applied. The retained pixels exceed panel needs at 600 dpi.
            if max(image.size) > max_side_px:
                scale = max_side_px / max(image.size)
                target = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
                image = image.resize(target, Image.Resampling.LANCZOS)
            photos[key] = image
    return photos


def style_axes(ax: mpl.axes.Axes) -> None:
    ax.tick_params(which="both", direction="out", pad=2.0)
    ax.spines["left"].set_linewidth(0.8)
    ax.spines["bottom"].set_linewidth(0.8)
    ax.grid(False)


def panel_label(ax: mpl.axes.Axes, letter: str, x: float = -0.16, y: float = 1.06) -> None:
    ax.text(
        x,
        y,
        letter,
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=10,
        fontweight="bold",
        color="black",
        clip_on=False,
    )


def line_kwargs(style: GroupStyle, markevery: int | tuple[int, int] | None = None) -> dict:
    return {
        "color": style.color,
        "linestyle": style.linestyle,
        "linewidth": 1.25,
        "marker": style.marker,
        "markersize": 3.2,
        "markerfacecolor": style.color if style.filled else "white",
        "markeredgecolor": style.color,
        "markeredgewidth": 0.75,
        "markevery": markevery,
        "solid_capstyle": "round",
        "zorder": 3,
    }


def legend_handles() -> list[Line2D]:
    handles: list[Line2D] = []
    for style in GROUPS:
        handles.append(Line2D([0], [0], label=style.label, **line_kwargs(style, None)))
    return handles


def draw_photo_panel(ax: mpl.axes.Axes, photos: dict[str, Image.Image]) -> None:
    ax.set_axis_off()
    layout = (
        ("SC7", [0.00, 0.53, 0.485, 0.43]),
        ("SC28", [0.515, 0.53, 0.485, 0.43]),
        ("MPC7", [0.00, 0.02, 0.485, 0.43]),
        ("MPC28", [0.515, 0.02, 0.485, 0.43]),
    )
    for key, bounds in layout:
        image_ax = ax.inset_axes(bounds)
        image_ax.imshow(photos[key], interpolation="nearest")
        image_ax.set_xticks([])
        image_ax.set_yticks([])
        for spine in image_ax.spines.values():
            spine.set_visible(True)
            spine.set_color("#6B6B6B")
            spine.set_linewidth(0.55)
        image_ax.set_title(str(PHOTO_SPECS[key]["label"]), fontsize=8.0, pad=1.5)
    # Keep the letter inside the outer photo-panel bounds so it remains visible
    # both in the full composite and in the separately exported Fig14a file.
    panel_label(ax, "a", x=0.00, y=1.04)


def draw_ocp_panel(
    ax: mpl.axes.Axes,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> None:
    for style in GROUPS:
        d = data[style.key]["OCPT"]
        t = np.asarray(d["time_s"])
        e = np.asarray(d["potential_V"])
        every = max(1, len(t) // 12)
        ax.plot(t, e, label=style.label, **line_kwargs(style, (every // 2, every)))
        ax.plot(
            t[-1], e[-1], marker=style.marker, markersize=4.2,
            markerfacecolor=style.color if style.filled else "white",
            markeredgecolor=style.color, markeredgewidth=0.8, linestyle="none", zorder=5,
        )
    ax.set_xlabel(r"Measurement time, $t$ (s)", labelpad=2)
    ax.set_ylabel(r"$E_{\mathrm{OCP}}$ (V vs. reference electrode)", labelpad=2)
    ax.set_xlim(left=0)
    ax.set_ylim(-0.65, -0.30)
    style_axes(ax)
    panel_label(ax, "b")


def draw_polarization_panel(
    ax: mpl.axes.Axes,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> None:
    for style in GROUPS:
        d = data[style.key]["TAFEL"]
        eta = np.asarray(d["eta_V"])
        current_density = np.abs(np.asarray(d["j_A_cm2"]))
        valid = current_density > 1e-12
        every = max(1, np.count_nonzero(valid) // 14)
        ax.plot(
            current_density[valid], eta[valid], label=style.label,
            **line_kwargs(style, (every // 2, every)),
        )
    ax.axhline(0.0, color="#8A8A8A", linestyle=(0, (3, 2)), linewidth=0.75, zorder=1)
    ax.set_xscale("log")
    ax.set_xlim(1e-8, 1e-1)
    ax.set_ylim(-0.25, 0.15)
    ax.set_xlabel(r"Absolute current density, $|j|$ (A cm$^{-2}$)", labelpad=2)
    ax.set_ylabel(r"Relative sweep potential, $E-E_{i=0}$ (V)", labelpad=2)
    style_axes(ax)
    panel_label(ax, "c", x=0.015, y=0.99)


def draw_nyquist_panel(
    ax: mpl.axes.Axes,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> None:
    max_extent = 0.0
    for style in GROUPS:
        d = data[style.key]["IMP"]
        f = np.asarray(d["frequency_Hz"])
        zr = np.asarray(d["zreal_area_ohm_cm2"])
        minus_zi = np.asarray(d["minus_zimag_area_ohm_cm2"])
        valid = f >= NYQUIST_MIN_FREQUENCY_HZ - 1e-12
        every = max(1, np.count_nonzero(valid) // 10)
        ax.plot(zr[valid], minus_zi[valid], label=style.label, **line_kwargs(style, every))
        idx = int(d["reference_index"])
        ax.plot(
            zr[idx], minus_zi[idx], marker=style.marker, markersize=5.0,
            markerfacecolor=style.color if style.filled else "white",
            markeredgecolor="black", markeredgewidth=0.65, linestyle="none", zorder=6,
        )
        max_extent = max(max_extent, float(np.max(zr[valid])), float(np.max(minus_zi[valid])))
    upper = math.ceil((max_extent * 1.08) / 10.0) * 10.0
    ax.set_xlim(0, upper)
    ax.set_ylim(-2, upper - 2)
    ax.set_aspect("equal", adjustable="box")
    ax.set_xlabel(r"$Z^\prime A$ ($\Omega$ cm$^2$)", labelpad=2)
    ax.set_ylabel(r"$-Z^{\prime\prime}A$ ($\Omega$ cm$^2$)", labelpad=2)
    style_axes(ax)
    panel_label(ax, "d")


def add_low_frequency_region(ax: mpl.axes.Axes) -> None:
    # A solid very-light grey is used instead of transparency so that the EPS,
    # PDF, SVG, TIFF, and PNG exports remain visually consistent.
    ax.axvspan(1e-2, 0.1, facecolor="#EEEEEE", alpha=1.0, edgecolor="none", zorder=0)
    ax.axvline(0.1, color="#7A7A7A", linestyle=(0, (2.5, 2.0)), linewidth=0.7, zorder=1)
    ax.text(
        0.018, 0.94, "$f<0.1$ Hz", transform=ax.get_xaxis_transform(),
        fontsize=8.0, color="#5F5F5F", ha="left", va="top",
    )


def draw_bode_modulus_panel(
    ax: mpl.axes.Axes,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> None:
    add_low_frequency_region(ax)
    for style in GROUPS:
        d = data[style.key]["IMP"]
        f = np.asarray(d["frequency_Hz"])
        zmag = np.asarray(d["modulus_area_ohm_cm2"])
        ax.plot(f, zmag, label=style.label, **line_kwargs(style, 6))
        idx = int(d["reference_index"])
        ax.plot(
            f[idx], zmag[idx], marker=style.marker, markersize=4.8,
            markerfacecolor=style.color if style.filled else "white",
            markeredgecolor="black", markeredgewidth=0.6, linestyle="none", zorder=6,
        )
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_xlim(1e-2, 1e5)
    ax.set_ylim(7.0, 300)
    ax.set_xlabel(r"Frequency, $f$ (Hz)", labelpad=2)
    ax.set_ylabel(r"$|Z|A$ ($\Omega$ cm$^2$)", labelpad=2)
    style_axes(ax)
    panel_label(ax, "e")


def draw_bode_phase_panel(
    ax: mpl.axes.Axes,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> None:
    add_low_frequency_region(ax)
    for style in GROUPS:
        d = data[style.key]["IMP"]
        f = np.asarray(d["frequency_Hz"])
        negative_phase = -np.asarray(d["phase_deg"])
        ax.plot(f, negative_phase, label=style.label, **line_kwargs(style, 6))
        idx = int(d["reference_index"])
        ax.plot(
            f[idx], negative_phase[idx], marker=style.marker, markersize=4.8,
            markerfacecolor=style.color if style.filled else "white",
            markeredgecolor="black", markeredgewidth=0.6, linestyle="none", zorder=6,
        )
    ax.set_xscale("log")
    ax.set_xlim(1e-2, 1e5)
    ax.set_ylim(-5, 170)
    ax.set_xlabel(r"Frequency, $f$ (Hz)", labelpad=2)
    ax.set_ylabel(r"Negative phase angle, $-\varphi$ ($^\circ$)", labelpad=2)
    style_axes(ax)
    panel_label(ax, "f")


def make_composite(
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
    photos: dict[str, Image.Image],
) -> plt.Figure:
    fig = plt.figure(
        figsize=(FINAL_WIDTH_MM / MM_PER_INCH, FINAL_HEIGHT_MM / MM_PER_INCH),
        facecolor="white",
    )
    grid = fig.add_gridspec(
        3,
        2,
        left=0.095,
        right=0.985,
        bottom=0.065,
        top=0.925,
        wspace=0.33,
        hspace=0.43,
        height_ratios=(1.08, 1.0, 1.0),
    )
    ax_a = fig.add_subplot(grid[0, 0])
    ax_b = fig.add_subplot(grid[0, 1])
    ax_c = fig.add_subplot(grid[1, 0])
    ax_d = fig.add_subplot(grid[1, 1])
    ax_e = fig.add_subplot(grid[2, 0])
    ax_f = fig.add_subplot(grid[2, 1])

    draw_photo_panel(ax_a, photos)
    draw_ocp_panel(ax_b, data)
    draw_polarization_panel(ax_c, data)
    draw_nyquist_panel(ax_d, data)
    draw_bode_modulus_panel(ax_e, data)
    draw_bode_phase_panel(ax_f, data)

    fig.legend(
        handles=legend_handles(),
        loc="upper center",
        bbox_to_anchor=(0.54, 0.988),
        ncol=5,
        handlelength=2.2,
        handletextpad=0.45,
        columnspacing=0.95,
        borderaxespad=0,
    )
    return fig


def make_standalone_photo(photos: dict[str, Image.Image]) -> plt.Figure:
    fig = plt.figure(
        figsize=(PANEL_WIDTH_MM / MM_PER_INCH, PANEL_HEIGHT_MM / MM_PER_INCH),
        facecolor="white",
    )
    ax = fig.add_axes([0.05, 0.04, 0.93, 0.93])
    draw_photo_panel(ax, photos)
    return fig


def make_standalone_plot(
    drawer: Callable[[mpl.axes.Axes, dict], None],
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> plt.Figure:
    fig = plt.figure(
        figsize=(PANEL_WIDTH_MM / MM_PER_INCH, PANEL_HEIGHT_MM / MM_PER_INCH),
        facecolor="white",
    )
    ax = fig.add_axes([0.20, 0.19, 0.76, 0.74])
    drawer(ax, data)
    return fig


def make_legend_strip() -> plt.Figure:
    """Create the shared five-group legend with the exact panel encodings."""
    fig = plt.figure(
        figsize=(LEGEND_WIDTH_MM / MM_PER_INCH, LEGEND_HEIGHT_MM / MM_PER_INCH),
        facecolor="white",
    )
    fig.legend(
        handles=legend_handles(),
        loc="center",
        bbox_to_anchor=(0.5, 0.5),
        ncol=5,
        handlelength=2.2,
        handletextpad=0.45,
        columnspacing=0.95,
        borderaxespad=0,
    )
    return fig


def output_paths(output_dir: Path, legend_only: bool = False) -> list[Path]:
    legend_path = output_dir / "Fig14_legend.png"
    if legend_only:
        return [legend_path]
    return [
        *(output_dir / f"Fig14{letter}.png" for letter in "abcdef"),
        legend_path,
    ]


def protect_existing_outputs(paths: Iterable[Path], overwrite: bool) -> None:
    existing = [p for p in paths if p.exists()]
    if existing and not overwrite:
        rendered = "\n".join(f"  - {p}" for p in existing)
        raise FileExistsError(
            "Generated output files already exist. No file was changed. "
            "Use --overwrite-generated only if you intentionally want to replace "
            f"previously generated Fig. 14 outputs:\n{rendered}"
        )


def save_figure(fig: plt.Figure, path: Path) -> None:
    """Save one standalone panel as a 600-dpi RGB PNG."""
    png_buffer = io.BytesIO()
    fig.savefig(png_buffer, format="png", dpi=PNG_DPI, facecolor="white")
    png_buffer.seek(0)
    with Image.open(png_buffer) as rendered:
        rendered.convert("RGB").save(
            path,
            format="PNG",
            dpi=(PNG_DPI, PNG_DPI),
        )
    plt.close(fig)


def write_csv(path: Path, header: list[str], rows: Iterable[Iterable[object]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def interpolate_abs_current(tafel: dict[str, np.ndarray | float], eta_target: float) -> float:
    eta = np.asarray(tafel["eta_V"])
    current_density = np.abs(np.asarray(tafel["j_A_cm2"]))
    order = np.argsort(eta)
    if eta_target < eta[order][0] or eta_target > eta[order][-1]:
        raise ValueError(f"Requested eta={eta_target} V lies outside measured data")
    return float(np.interp(eta_target, eta[order], current_density[order]))


def write_source_data(
    output_dir: Path,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
) -> None:
    ocp_rows = []
    tafel_rows = []
    eis_rows = []
    summary_rows = []

    for style in GROUPS:
        ocp = data[style.key]["OCPT"]
        for t, e in zip(np.asarray(ocp["time_s"]), np.asarray(ocp["potential_V"])):
            ocp_rows.append((style.key, style.label, t, e))

        tafel = data[style.key]["TAFEL"]
        for e, eta, current, current_density in zip(
            np.asarray(tafel["potential_V"]),
            np.asarray(tafel["eta_V"]),
            np.asarray(tafel["current_A"]),
            np.asarray(tafel["j_A_cm2"]),
        ):
            tafel_rows.append((style.key, style.label, e, eta, current, current_density))

        eis = data[style.key]["IMP"]
        for f, zraw, zarea, phase, zr, minus_zi in zip(
            np.asarray(eis["frequency_Hz"]),
            np.asarray(eis["modulus_raw_ohm"]),
            np.asarray(eis["modulus_area_ohm_cm2"]),
            np.asarray(eis["phase_deg"]),
            np.asarray(eis["zreal_area_ohm_cm2"]),
            np.asarray(eis["minus_zimag_area_ohm_cm2"]),
        ):
            eis_rows.append((style.key, style.label, f, zraw, zarea, phase, zr, minus_zi))

        summary_rows.append(
            (
                style.key,
                style.label,
                ocp["duration_s"],
                ocp["terminal_mean_V"],
                ocp["terminal_temporal_sd_V"],
                tafel["zero_current_potential_V"],
                tafel["jcorr_app_A_cm2"],
                tafel["rp_app_ohm_cm2"],
                tafel["rp_fit_r2"],
                tafel["tafel_anodic_r2"],
                tafel["tafel_cathodic_r2"],
                eis["modulus_at_0p1_ohm_cm2"],
                eis["phase_at_0p1_deg"],
                int(eis["low_frequency_nonstationary_points"]),
            )
        )

    write_csv(
        output_dir / "Fig14_source_OCP.csv",
        ["group_key", "group_label", "time_s", "open_circuit_potential_V"],
        ocp_rows,
    )
    write_csv(
        output_dir / "Fig14_source_polarization.csv",
        [
            "group_key", "group_label", "sweep_potential_V_from_bin_header",
            "potential_relative_to_sweep_zero_current_V", "current_A",
            "current_density_A_cm2",
        ],
        tafel_rows,
    )
    write_csv(
        output_dir / "Fig14_source_EIS.csv",
        [
            "group_key", "group_label", "frequency_Hz", "modulus_raw_ohm",
            "modulus_area_ohm_cm2", "phase_deg", "Zreal_area_ohm_cm2",
            "minus_Zimag_area_ohm_cm2",
        ],
        eis_rows,
    )
    write_csv(
        output_dir / "Fig14_summary_metrics.csv",
        [
            "group_key", "group_label", "OCP_duration_s",
            "OCP_terminal_10s_mean_V", "OCP_terminal_10s_temporal_SD_V",
            "interpolated_sweep_zero_current_potential_V", "apparent_jcorr_A_cm2",
            "apparent_local_Rp_ohm_cm2", "local_Rp_fit_R2",
            "anodic_Tafel_fit_R2", "cathodic_Tafel_fit_R2",
            "modulus_at_0p1Hz_ohm_cm2", "phase_at_0p1Hz_deg",
            "points_below_0p1Hz_with_phase_below_minus90deg",
        ],
        summary_rows,
    )

    pair_rows = []
    for age, sc_key, mpc_key in (("7 d", "SC7", "MPC7"), ("28 d", "SC28", "MPC28")):
        sc_tafel = data[sc_key]["TAFEL"]
        mpc_tafel = data[mpc_key]["TAFEL"]
        sc_eis = data[sc_key]["IMP"]
        mpc_eis = data[mpc_key]["IMP"]
        pair_rows.append(
            (
                age,
                "0.1-Hz impedance ratio, MPC-SC/SC",
                float(mpc_eis["modulus_at_0p1_ohm_cm2"])
                / float(sc_eis["modulus_at_0p1_ohm_cm2"]),
            )
        )
        pair_rows.append(
            (
                age,
                "local polarization-resistance ratio, MPC-SC/SC",
                float(mpc_tafel["rp_app_ohm_cm2"]) / float(sc_tafel["rp_app_ohm_cm2"]),
            )
        )
        pair_rows.append(
            (
                age,
                "apparent jcorr reduction by MPC-SC (%)",
                100.0
                * (
                    1.0
                    - float(mpc_tafel["jcorr_app_A_cm2"])
                    / float(sc_tafel["jcorr_app_A_cm2"])
                ),
            )
        )
        for eta in (-0.100, -0.050, 0.050, 0.100):
            sc_j = interpolate_abs_current(sc_tafel, eta)
            mpc_j = interpolate_abs_current(mpc_tafel, eta)
            pair_rows.append(
                (
                    age,
                    f"absolute-current-density reduction at eta={eta:+.3f} V (%)",
                    100.0 * (1.0 - mpc_j / sc_j),
                )
            )
    write_csv(
        output_dir / "Fig14_pairwise_effects.csv",
        ["exposure_age", "metric", "value"],
        pair_rows,
    )


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_qa_report(
    path: Path,
    base_dir: Path,
    pptx_path: Path,
    data: dict[str, dict[str, dict[str, np.ndarray | float]]],
    photos: dict[str, Image.Image],
) -> None:
    lines = [
        "FIGURE 14 DATA AND EXPORT QA REPORT",
        "====================================",
        "",
        "Figure claim:",
        "MPC stabilization increases interfacial and transport resistance and",
        "suppresses steel-corrosion kinetics in saline silty clay at 7 and 28 d.",
        "",
        f"Python: {platform.python_version()}",
        f"NumPy: {np.__version__}",
        f"Matplotlib: {mpl.__version__}",
        f"Pillow: {PILLOW_VERSION}",
        f"Working electrode area: {WORKING_AREA_CM2:.4f} cm^2",
        f"Composite size: {FINAL_WIDTH_MM:.1f} x {FINAL_HEIGHT_MM:.1f} mm",
        "TIFF: RGB, 600 dpi, LZW compression",
        "Primary editable outputs: SVG and PDF",
        "",
        "Data-processing rules:",
        "- OCP: raw samples, 0.1-s interval, no smoothing.",
        f"- Polarization: current divided by {WORKING_AREA_CM2:g} cm^2; the sweep-potential axis",
        "  was reconstructed from the acquisition fields at byte offsets 1088/1092",
        "  and the 1-mV step field at offset 1112.",
        "- Each curve is referenced to its interpolated zero-current position in",
        "  the recorded sweep; this position is neither OCP nor an independently",
        "  validated corrosion potential. Raw points are connected without smoothing.",
        "- Apparent local Rp: inverse slope of j versus E within +/-10 mV.",
        "- Apparent jcorr: identical 30--100 mV anodic/cathodic log|j| windows.",
        f"- EIS: impedance multiplied by {WORKING_AREA_CM2:g} cm^2.",
        "- Nyquist: f >= 0.1 Hz only; Bode: all raw frequencies retained and",
        "  f < 0.1 Hz visibly shaded as the low-frequency non-stationary region.",
        "- Photographs: PowerPoint crop/rotation reproduced; no brightness,",
        "  contrast, saturation, sharpening, local editing, or segmentation.",
        "",
        "Quantitative checks:",
    ]
    for style in GROUPS:
        ocp = data[style.key]["OCPT"]
        tafel = data[style.key]["TAFEL"]
        eis = data[style.key]["IMP"]
        lines.append(
            f"- {style.label}: OCP(last 10 s)="
            f"{float(ocp['terminal_mean_V']):.6f} +/- "
            f"{float(ocp['terminal_temporal_sd_V']):.6f} V; "
            f"apparent jcorr={float(tafel['jcorr_app_A_cm2']):.6e} A cm^-2; "
            f"Rp,app={float(tafel['rp_app_ohm_cm2']):.2f} ohm cm^2; "
            f"|Z|A(0.1 Hz)={float(eis['modulus_at_0p1_ohm_cm2']):.2f} ohm cm^2; "
            f"phase(0.1 Hz)={float(eis['phase_at_0p1_deg']):.2f} deg."
        )
    lines.extend(["", "Photograph panels:"])
    for key in ("SC7", "SC28", "MPC7", "MPC28"):
        lines.append(f"- {PHOTO_SPECS[key]['label']}: displayed pixels {photos[key].size}")

    lines.extend(["", "Input SHA-256 hashes:"])
    for style in GROUPS:
        for method, rel_path in FILE_MAP[style.key].items():
            source = base_dir / rel_path
            lines.append(f"- {style.key}/{method}: {sha256(source)}")
    lines.append(f"- corrosion photograph source: {sha256(pptx_path)}")
    lines.extend(
        [
            "",
            "Interpretation boundary:",
            "Only one electrochemical record is available for each condition. The",
            "curves and derived values describe the supplied measurements; no",
            "between-specimen error bars or inferential statistics were invented.",
            "The surface photographs are qualitative because illumination and field",
            "of view differ among the original photographs.",
            "",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8")


def validate_png_outputs(paths: Iterable[Path]) -> list[str]:
    messages: list[str] = []
    panel_size = (
        round(PANEL_WIDTH_MM / MM_PER_INCH * PNG_DPI),
        round(PANEL_HEIGHT_MM / MM_PER_INCH * PNG_DPI),
    )
    legend_size = (
        round(LEGEND_WIDTH_MM / MM_PER_INCH * PNG_DPI),
        round(LEGEND_HEIGHT_MM / MM_PER_INCH * PNG_DPI),
    )
    for png_path in paths:
        expected_size = legend_size if png_path.stem == "Fig14_legend" else panel_size
        with Image.open(png_path) as png:
            dpi = png.info.get("dpi", (None, None))
            if png.mode != "RGB":
                raise ValueError(f"{png_path.name} is not RGB")
            if png.size != expected_size:
                raise ValueError(
                    f"Unexpected {png_path.name} dimensions {png.size}; "
                    f"expected {expected_size}"
                )
            if any(value is None or abs(value - PNG_DPI) > 1 for value in dpi):
                raise ValueError(f"Unexpected {png_path.name} resolution: {dpi}")
            messages.append(
                f"{png_path.name}: {png.size[0]}x{png.size[1]} px, "
                f"mode={png.mode}, dpi={dpi}"
            )
    return messages


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-dir",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Directory containing the 7d, 28d, and blank-control subdirectories.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory; defaults to --base-dir.",
    )
    parser.add_argument(
        "--photo-pptx",
        type=Path,
        default=None,
        help="PowerPoint containing the original corrosion photographs.",
    )
    parser.add_argument(
        "--overwrite-generated",
        action="store_true",
        help="Allow replacement of previously generated Figure 14 PNG files only.",
    )
    parser.add_argument(
        "--legend-only",
        action="store_true",
        help="Generate only Fig14_legend.png, leaving Fig14a--Fig14f untouched.",
    )
    args = parser.parse_args()

    base_dir = args.base_dir.resolve()
    output_dir = (args.output_dir or base_dir).resolve()
    pptx_path = (args.photo_pptx or (base_dir.parent / "腐蚀.pptx")).resolve()
    if not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=False)

    planned = output_paths(output_dir, legend_only=args.legend_only)
    protect_existing_outputs(planned, args.overwrite_generated)
    configure_matplotlib()

    if args.legend_only:
        save_figure(make_legend_strip(), planned[0])
    else:
        data = load_all_data(base_dir)
        photos = load_photos(pptx_path)

        standalone_figures = {
            "Fig14a": make_standalone_photo(photos),
            "Fig14b": make_standalone_plot(draw_ocp_panel, data),
            "Fig14c": make_standalone_plot(draw_polarization_panel, data),
            "Fig14d": make_standalone_plot(draw_nyquist_panel, data),
            "Fig14e": make_standalone_plot(draw_bode_modulus_panel, data),
            "Fig14f": make_standalone_plot(draw_bode_phase_panel, data),
        }
        for stem, figure in standalone_figures.items():
            save_figure(figure, output_dir / f"{stem}.png")
        save_figure(make_legend_strip(), output_dir / "Fig14_legend.png")

    raster_messages = validate_png_outputs(planned)

    print("Figure 14 generation completed successfully.")
    print(f"Output directory: {output_dir}")
    for message in raster_messages:
        print(f"  {message}")


if __name__ == "__main__":
    main()
