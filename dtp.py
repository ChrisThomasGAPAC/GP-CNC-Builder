import csv
import copy
import json
import math
import os
import re
import sys
import tkinter as tk
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Set, Tuple

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None


APP_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = getattr(sys, "_MEIPASS", APP_DIR)
USER_DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", APP_DIR), "GP CNC Builder")
PROFILE_ROOT = os.path.join(USER_DATA_DIR, "profiles")


PRODUCT_PRESETS = {
    "ToughRock 48 x 96 x 5/8": {"product": "ToughRock", "width": "48", "height": "96", "thickness": "0.625"},
    "DensGlass 48 x 96 x 1/2": {"product": "DensGlass", "width": "48", "height": "96", "thickness": "0.5"},
    "DensShield 48 x 96 x 1/2": {"product": "DensShield", "width": "48", "height": "96", "thickness": "0.5"},
    "DensDeck Prime 48 x 96": {"product": "DensDeck Prime", "width": "48", "height": "96", "thickness": "0.625"},
    "Custom Board": {"product": "", "width": "48", "height": "96", "thickness": "0.625"},
}

SNAP_INCREMENTS = ["0.125", "0.25", "0.5", "1", "6", "12"]
SAMPLE_REQUEST_ROWS = 8
SHOP_EDGE_INSET_IN = 2.0


def resource_path(filename: str, fallback_path: str) -> str:
    bundled_path = os.path.join(BUNDLE_DIR, "assets", filename)
    if os.path.exists(bundled_path):
        return bundled_path
    local_path = os.path.join(APP_DIR, "assets", filename)
    if os.path.exists(local_path):
        return local_path
    return fallback_path


def sample_requires_formed_edge(sample: "SampleRect") -> bool:
    return sample.metadata.get("formed_edge") == "true"


def sample_uses_shop_edge_inset(sample: "SampleRect") -> bool:
    if sample.metadata.get("full_board_fixture") == "true":
        return False
    if sample_requires_formed_edge(sample):
        return False
    if sample.metadata.get("min_margin"):
        return False
    if sample.sample_id.startswith("DTP-11"):
        return False
    return True


GP_LOGO_PNG = resource_path(
    "georgia-pacific-300x300.png",
    r"C:\Users\chris\OneDrive\Pictures\georgia-pacific-300x300.png",
)
GP_LOGO_SVG = resource_path(
    "georgia-pacific-logo-svg-vector.svg",
    r"C:\Users\chris\OneDrive\Pictures\georgia-pacific-logo-svg-vector.svg",
)
GP_LOGIN_LOGO_JPG = resource_path(
    "GP Blog, Case Study & Press Room Thumbnail.jpg",
    r"C:\Users\chris\OneDrive\Pictures\GP Blog, Case Study & Press Room Thumbnail.jpg",
)


# Desktop MVP for a DTP -> CNC cut planning GUI.
# Uses Tkinter so it runs without extra GUI packages.
# Initial supported procedures:
# - DTP-11 Z-Directional Pull (Hot Melt)
# - DTP-13 Pull-Through
# - DTP-15 Abrasion


RULES_JSON = {
    "DTP-11": {
        "dtp_id": "DTP-11",
        "test_name": "Z-Directional Pull (Hot Melt)",
        "category": "cut+setup",
        "sample_width_in": 5.5,
        "sample_height_in": 5.5,
        "quantity_default": 3,
        "layout_mode": "fixed_procedure_layout",
        "side_offset_in": 3.0,
        "end_offset_in": 10.0,
        "track_code_side": True,
        "post_cut_module": "zdt_hotmelt_block_fixture",
        "notes": "Uses standard procedure layout with three 5.5 x 5.5 specimens per panel."
    },
    "DTP-13": {
        "dtp_id": "DTP-13",
        "test_name": "Pull-Through",
        "category": "cut+drill+setup",
        "sample_width_in": 14.0,
        "sample_height_in": 14.0,
        "quantity_default": 1,
        "layout_mode": "centered_square_with_drill",
        "drill_required": True,
        "drill_pattern": {
            "type": "single_center_hole",
            "diameter_in": 0.25
        },
        "orientation_options": ["Face Up", "Face Down"],
        "post_cut_module": "pullthrough_disc_fastener_cradle",
        "notes": "Square cut with a single center hole."
    },
    "DTP-15": {
        "dtp_id": "DTP-15",
        "test_name": "Abrasion",
        "category": "cut+setup",
        "sample_width_in": 3.5,
        "sample_height_in": 14.0,
        "quantity_default": 1,
        "layout_mode": "strip_nesting",
        "md_required": True,
        "long_axis_must_match_md": True,
        "post_cut_module": "abrasion_tester_setup",
        "notes": "Long dimension must follow machine direction."
    },
    "DTP-16": {
        "dtp_id": "DTP-16",
        "test_name": "Surface Indentation",
        "category": "cut+setup",
        "sample_width_in": 6.0,
        "sample_height_in": 6.0,
        "quantity_default": 10,
        "layout_mode": "margin_grid",
        "min_margin_in": 4.0,
        "layer": "DTP16",
        "notes": "Cut 6 x 6 samples a minimum of 4 in from each end or edge."
    },
    "DTP-17": {
        "dtp_id": "DTP-17",
        "test_name": "Soft Body Impact",
        "category": "full_board_fixture",
        "quantity_default": 1,
        "layout_mode": "full_board_fixture",
        "layer": "DTP17",
        "notes": "Uses a wall/sample fixture representative of installed use; build with 4 x 8 if possible."
    },
    "STP308": {
        "dtp_id": "STP308",
        "test_name": "Humid Bond",
        "category": "cut+score+setup",
        "sample_width_in": 6.0,
        "sample_height_in": 4.25,
        "quantity_default": 6,
        "layout_mode": "edge_with_score",
        "score_offset_in": 1.25,
        "layer": "STP308",
        "notes": "Cut 6.0 x 4.25 samples from paper bound edge with a score line 1.25 in from the leading end."
    },
    "STP312": {
        "dtp_id": "STP312",
        "test_name": "Transverse Strength - Flexural",
        "category": "cut+setup",
        "sample_width_in": 12.0,
        "sample_height_in": 16.0,
        "quantity_default": 4,
        "layout_mode": "flexural_split_cd_md",
        "min_margin_in": 4.0,
        "layer": "STP312",
        "notes": "Cut 12 x 16 specimens; split into CD and MD orientations, at least 4 in from board ends/edges."
    },
    "STP311": {
        "dtp_id": "STP311",
        "test_name": "Nail Pull",
        "category": "cut+setup",
        "sample_width_in": 6.0,
        "sample_height_in": 6.0,
        "quantity_default": 10,
        "layout_mode": "grid",
        "layer": "STP311",
        "notes": "Cut 6 x 6 nail pull samples."
    },
    "STP315": {
        "dtp_id": "STP315",
        "test_name": "Humidified Deflection (SAG)",
        "category": "cut+setup",
        "md_length_in": 12.0,
        "cross_length_in": 24.0,
        "quantity_default": 2,
        "layout_mode": "centered_md_rect",
        "min_end_in": 12.0,
        "centered_across_width": True,
        "layer": "STP315",
        "notes": "Cut 12 x 24 samples with the 12 in dimension in machine direction, centered across board width, at least 12 in from board ends."
    },
    "STP318": {
        "dtp_id": "STP318",
        "test_name": "Edge Compression Shear",
        "category": "cut+setup",
        "sample_width_in": 4.0,
        "sample_height_in": 6.0,
        "quantity_default": 6,
        "layout_mode": "formed_edge",
        "min_end_in": 6.0,
        "formed_edge_side_in": 4.0,
        "layer": "STP318",
        "notes": "Cut 4 x 6 specimens with the 4 in side on a formed edge, at least 6 in from board ends."
    }
}


PROCEDURE_STEPS = {
    "DTP-11": [
        "Cut 5.5 x 5.5 inch specimens using the standard layout.",
        "Place specimen face side up on the bench.",
        "Apply hot melt adhesive to the 3.5 x 3.5 inch wooden block in a zig-zag pattern.",
        "Center the block on the specimen face and press firmly.",
        "Allow adhesive to harden before conditioning and testing."
    ],
    "DTP-13": [
        "Cut a 14 x 14 inch sample.",
        "Drill a 0.25 inch hole in the exact center.",
        "Insert the fastener through the roofing disc and sample.",
        "Clamp the fastener into the upper jaw.",
        "Verify face-up or face-down orientation before testing."
    ],
    "DTP-15": [
        "Cut a 3.5 x 14 inch strip.",
        "Ensure the 14 inch direction follows machine direction.",
        "Mark the abrasion test area and zero the depth gauge.",
        "Secure the sample in the abrasion tester slots.",
        "Run 50 cycles and re-measure the marked points."
    ],
    "DTP-16": [
        "Cut 6 x 6 inch samples.",
        "Keep every sample at least 4 inches from each end or edge.",
        "Mark the center test area and zero the depth gauge.",
        "Place one sample on the Gardner impact pad.",
        "Drop the specified weight and measure the deepest indentation."
    ],
    "DTP-17": [
        "Build a representative wall/sample fixture, using 4 x 8 if possible.",
        "Fasten the board horizontally to 16 inch O.C. studs.",
        "Impact between studs with the soft body bag.",
        "Record impact energy and measure deflection.",
        "Repeat with a new sample and then repeat with vertical attachment."
    ],
    "STP308": [
        "Cut 6.0 x 4.25 inch humid bond samples from the paper bound edge.",
        "Score through the paper 1.25 inches from the leading end across the 6 inch length.",
        "Prepare three face and three back samples for each set.",
        "Condition in humidity chamber at the specified settings.",
        "Break along the score line and test immediately in the bond jig."
    ],
    "STP312": [
        "Cut 12 x 16 inch flexural specimens.",
        "Cut CD samples with the 16 inch dimension parallel to the board edge.",
        "Cut MD samples with the 16 inch dimension perpendicular to the board edge.",
        "Keep all specimens at least 4 inches from the end or edge.",
        "Label face-up/face-down and MD/CD orientation before testing."
    ],
    "STP311": [
        "Cut 6 x 6 inch nail pull samples.",
        "Handle samples carefully and avoid cracked or damaged pieces.",
        "Label each specimen before testing."
    ],
    "STP315": [
        "Cut 12 x 24 inch humidified deflection samples.",
        "Keep the 12 inch dimension in machine direction.",
        "Cut samples at least 12 inches from the board ends.",
        "Center samples across the board width.",
        "Condition and measure deflection across the 24 inch dimension."
    ],
    "STP318": [
        "Cut 4 x 6 inch edge compression shear specimens.",
        "Keep the 4 inch side on the formed edge.",
        "Cut specimens at least 6 inches from the board end.",
        "Prepare a minimum of three specimens from each side when required.",
        "Center the specimen in the holder and under the load bar."
    ]
}


MANUAL_OBJECTS = {
    "DTP-11": {
        "label": "DTP-11 Z Pull",
        "width": 5.5,
        "height": 5.5,
        "layer": "CUT",
        "metadata": {"dtp_id": "DTP-11"},
    },
    "DTP-13": {
        "label": "DTP-13 Pull Through",
        "width": 14.0,
        "height": 14.0,
        "layer": "CUT",
        "drill_diameter": 0.25,
        "metadata": {"dtp_id": "DTP-13"},
    },
    "DTP-15": {
        "label": "DTP-15 Abrasion",
        "md_length": 14.0,
        "cross_length": 3.5,
        "layer": "DTP15",
        "metadata": {"dtp_id": "DTP-15"},
    },
    "STP315": {
        "label": "STP315 Humidified Deflection",
        "md_length": 12.0,
        "cross_length": 24.0,
        "layer": "STP315",
    },
    "STP318": {
        "label": "STP318 Edge Shear",
        "width": 4.0,
        "height": 6.0,
        "layer": "STP318",
        "metadata": {"formed_edge": "true", "min_end": "6"},
    },
    "DTP-17": {
        "label": "DTP-17 Soft Body Impact",
        "width": 48.0,
        "height": 96.0,
        "layer": "DTP17",
        "metadata": {"dtp_id": "DTP-17", "full_board_fixture": "true"},
    },
    "STP308": {
        "label": "STP308 Humid Bond",
        "width": 6.0,
        "height": 4.25,
        "layer": "STP308",
        "score_offset": 1.25,
        "metadata": {"min_margin": "2"},
    },
    "STP312-CD": {
        "label": "STP312 Flexural CD",
        "width": 12.0,
        "height": 16.0,
        "layer": "STP312",
        "metadata": {"min_margin": "4"},
    },
    "STP312-MD": {
        "label": "STP312 Flexural MD",
        "width": 16.0,
        "height": 12.0,
        "layer": "STP312",
        "metadata": {"min_margin": "4"},
    },
    "STP311": {
        "label": "STP311 Nail Pull",
        "width": 6.0,
        "height": 6.0,
        "layer": "STP311",
    },
    "DTP16": {
        "label": "DTP-16 Surface Indentation",
        "width": 6.0,
        "height": 6.0,
        "layer": "DTP16",
        "metadata": {"min_margin": "4"},
    },
}

MANUAL_OBJECT_ALIASES = {
    "dtp-11": "DTP-11",
    "dtp11": "DTP-11",
    "dtp 11": "DTP-11",
    "dtp-13": "DTP-13",
    "dtp13": "DTP-13",
    "dtp 13": "DTP-13",
    "pull through": "DTP-13",
    "pull-through": "DTP-13",
    "dtp-15": "DTP-15",
    "dtp15": "DTP-15",
    "dtp 15": "DTP-15",
    "abrasion": "DTP-15",
    "abrasions": "DTP-15",
    "stp315": "STP315",
    "stp 315": "STP315",
    "humidified deflection": "STP315",
    "humidified deflection sag": "STP315",
    "sag": "STP315",
    "stp318": "STP318",
    "stp 318": "STP318",
    "edge shear": "STP318",
    "edge shears": "STP318",
    "edge compression shear": "STP318",
    "stp308": "STP308",
    "stp 308": "STP308",
    "humid bond": "STP308",
    "humid bonds": "STP308",
    "humid bond samples": "STP308",
    "stp312 cd": "STP312-CD",
    "stp312-cd": "STP312-CD",
    "cd flexural": "STP312-CD",
    "flexural cd": "STP312-CD",
    "stp312 md": "STP312-MD",
    "stp312-md": "STP312-MD",
    "md flexural": "STP312-MD",
    "flexural md": "STP312-MD",
    "stp312": "STP312",
    "stp 312": "STP312",
    "flexural": "STP312",
    "flexurals": "STP312",
    "transverse strength": "STP312",
    "stp311": "STP311",
    "stp 311": "STP311",
    "nail pull": "STP311",
    "nail pulls": "STP311",
    "nil pull": "STP311",
    "dtp16": "DTP-16",
    "dtp-16": "DTP-16",
    "dtp 16": "DTP-16",
    "surface indentation": "DTP-16",
    "surface indent": "DTP-16",
    "indentation": "DTP-16",
    "indentations": "DTP-16",
    "dtp17": "DTP-17",
    "dtp-17": "DTP-17",
    "dtp 17": "DTP-17",
    "soft body impact": "DTP-17",
}

BOT_NUMBER_WORDS = {
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
    "ten": "10",
    "eleven": "11",
    "twelve": "12",
    "thirteen": "13",
    "fourteen": "14",
    "fifteen": "15",
    "sixteen": "16",
    "seventeen": "17",
    "eighteen": "18",
    "nineteen": "19",
    "twenty": "20",
}


@dataclass
class SampleRect:
    sample_id: str
    x_in: float
    y_in: float
    width_in: float
    height_in: float
    rotation_deg: int = 0
    drill_centers: List[Tuple[float, float, float]] = field(default_factory=list)  # x, y, diameter
    metadata: Dict[str, str] = field(default_factory=dict)


@dataclass
class LineEntity:
    layer: str
    x1_in: float
    y1_in: float
    x2_in: float
    y2_in: float


@dataclass
class TextEntity:
    layer: str
    x_in: float
    y_in: float
    height_in: float
    text: str


@dataclass
class LayoutResult:
    board_width_in: float
    board_height_in: float
    dtp_id: str
    samples: List[SampleRect]
    scrap_zones: List[Tuple[float, float, float, float]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    line_entities: List[LineEntity] = field(default_factory=list)
    text_entities: List[TextEntity] = field(default_factory=list)
    metadata: Dict[str, str] = field(default_factory=dict)


@dataclass
class LayoutRequest:
    dtp_id: str
    quantity: int
    request_index: int


class RuleEngine:
    def __init__(self, rules: Dict[str, Dict]):
        self.rules = rules

    def generate_layout(
        self,
        dtp_id: str,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        machine_direction: str,
        orientation: str,
        code_side: str,
    ) -> LayoutResult:
        if dtp_id not in self.rules:
            raise ValueError(f"Unsupported DTP: {dtp_id}")

        rule = self.rules[dtp_id]
        warnings: List[str] = []

        if board_width_in <= 0 or board_height_in <= 0:
            raise ValueError("Board dimensions must be greater than zero.")
        if quantity <= 0:
            raise ValueError("Quantity must be at least 1.")

        if rule.get("md_required") and machine_direction not in {"Horizontal", "Vertical"}:
            raise ValueError("Machine direction is required for this DTP.")

        if dtp_id == "DTP-11":
            return self._generate_dtp11(rule, board_width_in, board_height_in, quantity, warnings, code_side)
        if dtp_id == "DTP-13":
            return self._generate_grid_layout(rule, board_width_in, board_height_in, quantity, warnings, orientation)
        if dtp_id == "DTP-15":
            return self._generate_dtp15(rule, board_width_in, board_height_in, quantity, warnings, machine_direction)
        if dtp_id == "DTP-16":
            return self._generate_margin_grid(rule, board_width_in, board_height_in, quantity, warnings)
        if dtp_id == "DTP-17":
            return self._generate_full_board_fixture(rule, board_width_in, board_height_in, quantity, warnings)
        if dtp_id == "STP308":
            return self._generate_stp308(rule, board_width_in, board_height_in, quantity, warnings)
        if dtp_id == "STP312":
            return self._generate_stp312(rule, board_width_in, board_height_in, quantity, warnings)
        if dtp_id == "STP311":
            return self._generate_stp311(rule, board_width_in, board_height_in, quantity, warnings)
        if dtp_id == "STP315":
            return self._generate_stp315(rule, board_width_in, board_height_in, quantity, warnings, machine_direction)
        if dtp_id == "STP318":
            return self._generate_stp318(rule, board_width_in, board_height_in, quantity, warnings)

        raise ValueError(f"No generator implemented for {dtp_id}")

    def generate_combined_layout(
        self,
        requests: List[LayoutRequest],
        board_width_in: float,
        board_height_in: float,
        machine_direction: str,
        orientation: str,
        code_side: str,
    ) -> LayoutResult:
        if board_width_in <= 0 or board_height_in <= 0:
            raise ValueError("Board dimensions must be greater than zero.")
        if not requests:
            raise ValueError("Select at least one sample type.")
        if any(request.dtp_id == "DTP-17" for request in requests) and len(requests) > 1:
            raise ValueError("DTP-17 is a full-board fixture and cannot be mixed with other sample cuts on one sheet.")

        warnings: List[str] = []
        samples: List[SampleRect] = []
        has_dtp11 = any(request.dtp_id == "DTP-11" for request in requests)
        scrap_zones: List[Tuple[float, float, float, float]] = []
        x_min = 0.0
        y_min = 0.0
        x_max = board_width_in
        y_max = board_height_in

        if has_dtp11:
            rule = self.rules["DTP-11"]
            side_offset = rule["side_offset_in"]
            end_offset = rule["end_offset_in"]
            x_min = side_offset
            y_min = end_offset
            x_max = board_width_in - side_offset
            y_max = board_height_in - end_offset
            if x_max <= x_min or y_max <= y_min:
                raise ValueError("Board is too small after applying DTP-11 discard zones.")
            scrap_zones = [
                (0, 0, board_width_in, end_offset),
                (0, board_height_in - end_offset, board_width_in, end_offset),
                (0, 0, side_offset, board_height_in),
                (board_width_in - side_offset, 0, side_offset, board_height_in),
            ]
            warnings.append("DTP-11 is included, so the mixed layout starts inside the DTP-11 discard limits.")

        for request in requests:
            samples.extend(
                self._build_unplaced_samples(
                    request,
                    machine_direction,
                    orientation,
                    code_side,
                    warnings,
                )
            )

        self._pack_samples(samples, x_min, y_min, x_max, y_max)
        self._add_post_pack_features(samples, warnings)
        summary = ", ".join(f"{request.dtp_id} x {request.quantity}" for request in requests)
        warnings.append(f"Mixed sample plan generated for: {summary}.")
        layout = LayoutResult(board_width_in, board_height_in, "MIXED", samples, scrap_zones, warnings)
        self._add_stp308_score_lines(layout)
        return layout

    def _build_unplaced_samples(
        self,
        request: LayoutRequest,
        machine_direction: str,
        orientation: str,
        code_side: str,
        warnings: List[str],
    ) -> List[SampleRect]:
        if request.dtp_id not in self.rules:
            raise ValueError(f"Unsupported DTP: {request.dtp_id}")
        if request.quantity <= 0:
            raise ValueError(f"{request.dtp_id} quantity must be at least 1.")

        rule = self.rules[request.dtp_id]
        samples: List[SampleRect] = []
        metadata_base = {
            "dtp_id": request.dtp_id,
            "request": str(request.request_index),
        }

        if rule.get("md_required") and machine_direction not in {"Horizontal", "Vertical"}:
            raise ValueError(f"Machine direction is required for {request.dtp_id}.")

        if request.dtp_id == "DTP-11":
            per_panel = 3
            if request.quantity < per_panel or request.quantity % per_panel != 0:
                raise ValueError("DTP-11 quantity must be entered as complete three-sample sets: 3, 6, 9, etc.")
            for sample_index in range(request.quantity):
                panel_index = sample_index // per_panel + 1
                sample_in_set = sample_index % per_panel + 1
                samples.append(
                    SampleRect(
                        sample_id=f"DTP-11-R{request.request_index}-P{panel_index}-S{sample_in_set}",
                        x_in=0,
                        y_in=0,
                        width_in=rule["sample_width_in"],
                        height_in=rule["sample_height_in"],
                        metadata={
                            **metadata_base,
                            "code_side": code_side,
                            "panel_index": str(panel_index),
                        },
                    )
                )
            warnings.append(f"DTP-11 request {request.request_index} uses complete three-sample set(s).")
            return samples

        if request.dtp_id == "DTP-13":
            diameter = rule["drill_pattern"]["diameter_in"]
            for sample_index in range(request.quantity):
                sample_w = rule["sample_width_in"]
                sample_h = rule["sample_height_in"]
                samples.append(
                    SampleRect(
                        sample_id=f"DTP-13-R{request.request_index}-S{sample_index + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=sample_w,
                        height_in=sample_h,
                        drill_centers=[(sample_w / 2, sample_h / 2, diameter)],
                        metadata={**metadata_base, "orientation": orientation},
                    )
                )
            warnings.append(f"DTP-13 request {request.request_index} includes a center hole for every sample.")
            return samples

        if request.dtp_id == "DTP-15":
            sample_w = rule["sample_width_in"]
            sample_h = rule["sample_height_in"]
            if machine_direction == "Horizontal":
                piece_w, piece_h = sample_h, sample_w
                rotation = 0
            else:
                piece_w, piece_h = sample_w, sample_h
                rotation = 90
            for sample_index in range(request.quantity):
                samples.append(
                    SampleRect(
                        sample_id=f"DTP-15-R{request.request_index}-S{sample_index + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=piece_w,
                        height_in=piece_h,
                        rotation_deg=rotation,
                        metadata={**metadata_base, "machine_direction": machine_direction},
                    )
                )
            warnings.append(f"DTP-15 request {request.request_index} keeps the long axis aligned to machine direction.")
            return samples

        if request.dtp_id in {"DTP-16", "STP308", "STP312", "STP311", "STP315", "STP318", "DTP-17"}:
            return self._build_generic_unplaced_samples(request, machine_direction, warnings)

        raise ValueError(f"No generator implemented for {request.dtp_id}")

    def _build_generic_unplaced_samples(
        self,
        request: LayoutRequest,
        machine_direction: str,
        warnings: List[str],
    ) -> List[SampleRect]:
        rule = self.rules[request.dtp_id]
        samples: List[SampleRect] = []

        if request.dtp_id == "DTP-16":
            for i in range(request.quantity):
                samples.append(
                    SampleRect(
                        sample_id=f"DTP-16-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=rule["sample_width_in"],
                        height_in=rule["sample_height_in"],
                        metadata={"dtp_id": "DTP-16", "layer": "DTP16", "min_margin": "4"},
                    )
                )
            warnings.append(f"DTP-16 request {request.request_index} keeps samples 4 in from all edges.")
            return samples

        if request.dtp_id == "STP308":
            for i in range(request.quantity):
                samples.append(
                    SampleRect(
                        sample_id=f"STP308-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=rule["sample_width_in"],
                        height_in=rule["sample_height_in"],
                        metadata={"dtp_id": "STP308", "layer": "STP308", "min_margin": "2", "score_offset": "1.25"},
                    )
                )
            warnings.append(f"STP308 request {request.request_index} includes 1.25 in score lines and stays at least 2 in off the board edge.")
            return samples

        if request.dtp_id == "STP312":
            cd_qty = request.quantity // 2
            md_qty = request.quantity - cd_qty
            for i in range(cd_qty):
                samples.append(
                    SampleRect(
                        sample_id=f"STP312-CD-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=12,
                        height_in=16,
                        metadata={"dtp_id": "STP312", "layer": "STP312", "min_margin": "4", "orientation": "CD"},
                    )
                )
            for i in range(md_qty):
                samples.append(
                    SampleRect(
                        sample_id=f"STP312-MD-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=16,
                        height_in=12,
                        metadata={"dtp_id": "STP312", "layer": "STP312", "min_margin": "4", "orientation": "MD"},
                    )
                )
            warnings.append(f"STP312 request {request.request_index} split quantity between CD and MD samples.")
            return samples

        if request.dtp_id == "STP311":
            for i in range(request.quantity):
                samples.append(
                    SampleRect(
                        sample_id=f"STP311-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=rule["sample_width_in"],
                        height_in=rule["sample_height_in"],
                        metadata={"dtp_id": "STP311", "layer": "STP311"},
                    )
                )
            warnings.append(f"STP311 request {request.request_index} uses 6 x 6 nail pull samples.")
            return samples

        if request.dtp_id == "STP315":
            if machine_direction == "Vertical":
                piece_w, piece_h = rule["cross_length_in"], rule["md_length_in"]
            else:
                piece_w, piece_h = rule["md_length_in"], rule["cross_length_in"]
            for i in range(request.quantity):
                samples.append(
                    SampleRect(
                        sample_id=f"STP315-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=piece_w,
                        height_in=piece_h,
                        metadata={
                            "dtp_id": "STP315",
                            "layer": "STP315",
                            "min_end": str(rule["min_end_in"]),
                            "centered_width": "true",
                            "machine_direction": machine_direction,
                        },
                    )
                )
            warnings.append(f"STP315 request {request.request_index} keeps 12 in dimension in machine direction.")
            return samples

        if request.dtp_id == "STP318":
            for i in range(request.quantity):
                samples.append(
                    SampleRect(
                        sample_id=f"STP318-R{request.request_index}-S{i + 1}",
                        x_in=0,
                        y_in=0,
                        width_in=rule["sample_width_in"],
                        height_in=rule["sample_height_in"],
                        metadata={"dtp_id": "STP318", "layer": "STP318", "formed_edge": "true", "min_end": "6"},
                    )
                )
            warnings.append(f"STP318 request {request.request_index} keeps the 4 in side on a formed edge.")
            return samples

        raise ValueError(f"{request.dtp_id} cannot be mixed with other sample cuts.")

    def _add_post_pack_features(self, samples: List[SampleRect], warnings: List[str]) -> None:
        if any(sample.metadata.get("dtp_id") == "STP308" for sample in samples):
            warnings.append("STP308 score lines are drawn 1.25 in from the leading end.")

    def _add_stp308_score_lines(self, layout: LayoutResult) -> None:
        for sample in layout.samples:
            if sample.metadata.get("dtp_id") == "STP308" or sample.sample_id.startswith("STP308"):
                offset = float(sample.metadata.get("score_offset", "1.25"))
                score_y = sample.y_in + offset
                layout.line_entities.append(
                    LineEntity("STP308_SCORE", sample.x_in, score_y, sample.x_in + sample.width_in, score_y)
                )

    def _pack_samples(
        self,
        samples: List[SampleRect],
        x_min: float,
        y_min: float,
        x_max: float,
        y_max: float,
    ) -> None:
        gap = 0.25
        ordered_samples = sorted(samples, key=lambda sample: (-sample.height_in, -sample.width_in, sample.sample_id))
        placed: List[SampleRect] = []

        for sample in ordered_samples:
            if sample.width_in > x_max - x_min or sample.height_in > y_max - y_min:
                raise ValueError(f"{sample.sample_id} is too large for the available board area.")

            placed_sample = False
            candidate_ys = self._candidate_axis_positions(y_min, y_max, sample.height_in, gap)
            for y in candidate_ys:
                candidate_xs = self._candidate_x_positions(sample, x_min, x_max, y_max, gap)
                for x in candidate_xs:
                    if self._sample_position_allowed(sample, x, y, x_min, y_min, x_max, y_max, placed):
                        dx = x - sample.x_in
                        dy = y - sample.y_in
                        sample.x_in = x
                        sample.y_in = y
                        sample.drill_centers = [(cx + dx, cy + dy, dia) for cx, cy, dia in sample.drill_centers]
                        placed.append(sample)
                        placed_sample = True
                        break
                if placed_sample:
                    break

            if not placed_sample:
                raise ValueError(f"The selected sample types do not fit together on this board. Could not place {sample.sample_id}.")

    @staticmethod
    def _candidate_axis_positions(start: float, stop: float, length: float, step: float) -> List[float]:
        positions = []
        current = start
        while current + length <= stop + 1e-6:
            positions.append(round(current, 4))
            current += step
        return positions

    def _candidate_x_positions(
        self,
        sample: SampleRect,
        x_min: float,
        x_max: float,
        board_height: float,
        step: float,
    ) -> List[float]:
        if sample.metadata.get("centered_width") == "true":
            centered_x = (x_max - sample.width_in) / 2
            return [round(centered_x, 4)]
        if sample_requires_formed_edge(sample):
            if board_height >= x_max:
                return [x for x in [0.0, x_max - sample.width_in] if x_min - 1e-6 <= x <= x_max - sample.width_in + 1e-6]
        return self._candidate_axis_positions(x_min, x_max, sample.width_in, step)

    def _sample_position_allowed(
        self,
        sample: SampleRect,
        x: float,
        y: float,
        x_min: float,
        y_min: float,
        x_max: float,
        y_max: float,
        placed: List[SampleRect],
    ) -> bool:
        eps = 1e-6
        if x < x_min - eps or y < y_min - eps or x + sample.width_in > x_max + eps or y + sample.height_in > y_max + eps:
            return False

        min_margin = sample.metadata.get("min_margin")
        if min_margin:
            margin = float(min_margin)
            if x < margin - eps or y < margin - eps or x + sample.width_in > x_max - margin + eps or y + sample.height_in > y_max - margin + eps:
                return False

        if sample_uses_shop_edge_inset(sample):
            margin = SHOP_EDGE_INSET_IN
            if x < margin - eps or y < margin - eps or x + sample.width_in > x_max - margin + eps or y + sample.height_in > y_max - margin + eps:
                return False

        min_end = sample.metadata.get("min_end")
        if min_end:
            margin = float(min_end)
            if y_max >= x_max:
                if y < margin - eps or y + sample.height_in > y_max - margin + eps:
                    return False
            elif x < margin - eps or x + sample.width_in > x_max - margin + eps:
                return False

        if sample_requires_formed_edge(sample):
            if y_max >= x_max:
                on_edge = abs(x) <= 0.05 or abs(x + sample.width_in - x_max) <= 0.05
            else:
                on_edge = abs(y) <= 0.05 or abs(y + sample.height_in - y_max) <= 0.05
            if not on_edge:
                return False

        if sample.metadata.get("centered_width") == "true":
            centered_x = (x_max - sample.width_in) / 2
            if abs(x - centered_x) > 0.05:
                return False

        for other in placed:
            if self._rects_overlap_static(x, y, sample.width_in, sample.height_in, other.x_in, other.y_in, other.width_in, other.height_in):
                return False
        return True

    @staticmethod
    def _rects_overlap_static(ax: float, ay: float, aw: float, ah: float, bx: float, by: float, bw: float, bh: float) -> bool:
        eps = 1e-6
        return ax < bx + bw - eps and ax + aw > bx + eps and ay < by + bh - eps and ay + ah > by + eps

    def _generate_margin_grid(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
    ) -> LayoutResult:
        samples = [
            SampleRect(
                sample_id=f"{rule['dtp_id']}-S{i + 1}",
                x_in=0,
                y_in=0,
                width_in=rule["sample_width_in"],
                height_in=rule["sample_height_in"],
                metadata={
                    "dtp_id": rule["dtp_id"],
                    "layer": rule.get("layer", "CUT"),
                    "min_margin": str(rule.get("min_margin_in", 0)),
                },
            )
            for i in range(quantity)
        ]
        margin = rule.get("min_margin_in", 0.0)
        self._pack_samples(samples, 0, 0, board_width_in, board_height_in)
        warnings.append(f"{rule['dtp_id']} samples are kept at least {margin:g} in from every board edge/end.")
        return LayoutResult(board_width_in, board_height_in, rule["dtp_id"], samples, [], warnings)

    def _generate_full_board_fixture(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
    ) -> LayoutResult:
        if quantity != 1:
            raise ValueError("DTP-17 uses one full-board/wall fixture per sheet. Save additional sheets for additional fixtures.")
        sample = SampleRect(
            sample_id="DTP-17-FIXTURE-1",
            x_in=0,
            y_in=0,
            width_in=board_width_in,
            height_in=board_height_in,
            metadata={"dtp_id": "DTP-17", "layer": "DTP17", "full_board_fixture": "true"},
        )
        warnings.append("DTP-17 is shown as a full-board wall fixture; use separate sheets for duplicate test walls.")
        layout = LayoutResult(board_width_in, board_height_in, "DTP-17", [sample], [], warnings)
        layout.text_entities.append(TextEntity("LABELS", 1, 1, 0.45, "DTP-17 Soft Body Impact Fixture"))
        return layout

    def _generate_stp308(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
    ) -> LayoutResult:
        samples = [
            SampleRect(
                sample_id=f"STP308-S{i + 1}",
                x_in=0,
                y_in=0,
                width_in=rule["sample_width_in"],
                height_in=rule["sample_height_in"],
                metadata={"dtp_id": "STP308", "layer": "STP308", "min_margin": "2", "score_offset": str(rule["score_offset_in"])},
            )
            for i in range(quantity)
        ]
        self._pack_samples(samples, 0, 0, board_width_in, board_height_in)
        layout = LayoutResult(board_width_in, board_height_in, "STP308", samples, [], warnings)
        self._add_stp308_score_lines(layout)
        warnings.append("STP308 samples stay at least 2 in off the board edge and include 1.25 in score lines.")
        return layout

    def _generate_stp312(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
    ) -> LayoutResult:
        samples = self._build_generic_unplaced_samples(LayoutRequest("STP312", quantity, 1), "Horizontal", warnings)
        margin = rule.get("min_margin_in", 4.0)
        self._pack_samples(samples, 0, 0, board_width_in, board_height_in)
        warnings.append("STP312 quantity was split between CD and MD orientations.")
        return LayoutResult(board_width_in, board_height_in, "STP312", samples, [], warnings)

    def _generate_stp311(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
    ) -> LayoutResult:
        samples = [
            SampleRect(
                sample_id=f"STP311-S{i + 1}",
                x_in=0,
                y_in=0,
                width_in=rule["sample_width_in"],
                height_in=rule["sample_height_in"],
                metadata={"dtp_id": "STP311", "layer": "STP311"},
            )
            for i in range(quantity)
        ]
        self._pack_samples(samples, 0, 0, board_width_in, board_height_in)
        warnings.append("STP311 generated 6 x 6 nail pull samples.")
        return LayoutResult(board_width_in, board_height_in, "STP311", samples, [], warnings)

    def _generate_stp315(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
        machine_direction: str,
    ) -> LayoutResult:
        samples = self._build_generic_unplaced_samples(LayoutRequest("STP315", quantity, 1), machine_direction, warnings)
        self._pack_samples(samples, 0, 0, board_width_in, board_height_in)
        warnings.append("STP315 samples are centered across board width and at least 12 in from board ends.")
        return LayoutResult(board_width_in, board_height_in, "STP315", samples, [], warnings)

    def _generate_stp318(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
    ) -> LayoutResult:
        samples = [
            SampleRect(
                sample_id=f"STP318-S{i + 1}",
                x_in=0,
                y_in=0,
                width_in=rule["sample_width_in"],
                height_in=rule["sample_height_in"],
                metadata={"dtp_id": "STP318", "layer": "STP318", "formed_edge": "true", "min_end": str(rule["min_end_in"])},
            )
            for i in range(quantity)
        ]
        self._pack_samples(samples, 0, 0, board_width_in, board_height_in)
        warnings.append("STP318 samples keep the 4 in side on a formed edge and stay 6 in from board ends.")
        return LayoutResult(board_width_in, board_height_in, "STP318", samples, [], warnings)

    def _generate_dtp11(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
        code_side: str,
    ) -> LayoutResult:
        sample_w = rule["sample_width_in"]
        sample_h = rule["sample_height_in"]
        side_offset = rule["side_offset_in"]
        end_offset = rule["end_offset_in"]
        per_panel = 3

        usable_w = board_width_in - 2 * side_offset
        usable_h = board_height_in - 2 * end_offset
        if usable_w <= 0 or usable_h <= 0:
            raise ValueError("Board is too small after applying DTP-11 discard zones.")

        cols = max(1, int(usable_w // sample_w))
        rows = max(1, int(usable_h // sample_h))
        capacity = cols * rows
        if capacity < per_panel:
            raise ValueError("Board is too small for the standard DTP-11 three-sample layout.")
        if quantity < per_panel or quantity % per_panel != 0:
            raise ValueError("DTP-11 quantity must be entered as complete three-sample sets: 3, 6, 9, etc.")
        if capacity < quantity:
            raise ValueError(f"Board capacity is {capacity} DTP-11 samples, but {quantity} were requested.")

        samples: List[SampleRect] = []
        scrap_zones = [
            (0, 0, board_width_in, end_offset),
            (0, board_height_in - end_offset, board_width_in, end_offset),
            (0, 0, side_offset, board_height_in),
            (board_width_in - side_offset, 0, side_offset, board_height_in),
        ]

        for sample_index in range(quantity):
            row = sample_index // cols
            col = sample_index % cols
            x = side_offset + col * sample_w
            y = end_offset + row * sample_h
            if x + sample_w > board_width_in - side_offset or y + sample_h > board_height_in - end_offset:
                raise ValueError("Not enough usable room for the requested DTP-11 samples.")
            panel_index = sample_index // per_panel + 1
            sample_in_set = sample_index % per_panel + 1
            sample_id = f"{rule['dtp_id']}-P{panel_index}-S{sample_in_set}"
            samples.append(
                SampleRect(
                    sample_id=sample_id,
                    x_in=x,
                    y_in=y,
                    width_in=sample_w,
                    height_in=sample_h,
                    metadata={
                        "code_side": code_side,
                        "panel_index": str(panel_index),
                    },
                )
            )

        warnings.append(
            f"DTP-11 generated {quantity} specimens as {quantity // per_panel} complete three-sample set(s), "
            "with 10 in end discard and 3 in side discard."
        )
        return LayoutResult(board_width_in, board_height_in, "DTP-11", samples, scrap_zones, warnings)

    def _generate_grid_layout(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
        orientation: str,
    ) -> LayoutResult:
        sample_w = rule["sample_width_in"]
        sample_h = rule["sample_height_in"]
        gap = 0.25
        edge_inset = SHOP_EDGE_INSET_IN

        usable_w = board_width_in - 2 * edge_inset
        usable_h = board_height_in - 2 * edge_inset
        cols = max(1, int((usable_w + gap) // (sample_w + gap)))
        rows = max(1, int((usable_h + gap) // (sample_h + gap)))
        capacity = cols * rows
        if capacity < quantity:
            raise ValueError(f"Board capacity is {capacity} samples, but {quantity} were requested.")

        samples: List[SampleRect] = []
        x0 = edge_inset
        y0 = edge_inset
        count = 0
        for r in range(rows):
            for c in range(cols):
                if count >= quantity:
                    break
                x = x0 + c * (sample_w + gap)
                y = y0 + r * (sample_h + gap)
                cx = x + sample_w / 2
                cy = y + sample_h / 2
                diameter = rule["drill_pattern"]["diameter_in"]
                count += 1
                samples.append(
                    SampleRect(
                        sample_id=f"{rule['dtp_id']}-S{count}",
                        x_in=x,
                        y_in=y,
                        width_in=sample_w,
                        height_in=sample_h,
                        drill_centers=[(cx, cy, diameter)],
                        metadata={"dtp_id": rule["dtp_id"], "orientation": orientation},
                    )
                )
            if count >= quantity:
                break

        warnings.append("Center hole is included for every DTP-13 sample; samples stay at least 2 in off board edges.")
        return LayoutResult(board_width_in, board_height_in, rule["dtp_id"], samples, [], warnings)

    def _generate_dtp15(
        self,
        rule: Dict,
        board_width_in: float,
        board_height_in: float,
        quantity: int,
        warnings: List[str],
        machine_direction: str,
    ) -> LayoutResult:
        sample_w = rule["sample_width_in"]
        sample_h = rule["sample_height_in"]
        gap = 0.25
        edge_inset = SHOP_EDGE_INSET_IN

        # Long axis must match machine direction.
        if machine_direction == "Horizontal":
            piece_w, piece_h = sample_h, sample_w
            rotation = 0
        else:
            piece_w, piece_h = sample_w, sample_h
            rotation = 90

        usable_w = board_width_in - 2 * edge_inset
        usable_h = board_height_in - 2 * edge_inset
        cols = max(1, int((usable_w + gap) // (piece_w + gap)))
        rows = max(1, int((usable_h + gap) // (piece_h + gap)))
        capacity = cols * rows
        if capacity < quantity:
            raise ValueError(f"Board capacity is {capacity} abrasion samples, but {quantity} were requested.")

        samples: List[SampleRect] = []
        count = 0
        for r in range(rows):
            for c in range(cols):
                if count >= quantity:
                    break
                x = edge_inset + c * (piece_w + gap)
                y = edge_inset + r * (piece_h + gap)
                count += 1
                samples.append(
                    SampleRect(
                        sample_id=f"{rule['dtp_id']}-S{count}",
                        x_in=x,
                        y_in=y,
                        width_in=piece_w,
                        height_in=piece_h,
                        rotation_deg=rotation,
                        metadata={"dtp_id": rule["dtp_id"], "machine_direction": machine_direction},
                    )
                )
            if count >= quantity:
                break

        warnings.append("DTP-15 long axis was aligned to the selected machine direction; samples stay at least 2 in off board edges.")
        return LayoutResult(board_width_in, board_height_in, rule["dtp_id"], samples, [], warnings)


class SvgExporter:
    @staticmethod
    def export(path: str, layout: LayoutResult) -> None:
        scale = 20  # px per inch
        width_px = layout.board_width_in * scale
        height_px = layout.board_height_in * scale

        lines = [
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{width_px}" height="{height_px}" viewBox="0 0 {width_px} {height_px}">',
            '<rect x="0" y="0" width="100%" height="100%" fill="white" stroke="black" stroke-width="2" />',
        ]

        for x, y, w, h in layout.scrap_zones:
            lines.append(
                f'<rect x="{x * scale}" y="{y * scale}" width="{w * scale}" height="{h * scale}" '
                'fill="#f8d7da" stroke="#dc3545" stroke-width="1" opacity="0.5" />'
            )

        for sample in layout.samples:
            lines.append(
                f'<rect x="{sample.x_in * scale}" y="{sample.y_in * scale}" width="{sample.width_in * scale}" '
                f'height="{sample.height_in * scale}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="2" />'
            )
            text_x = (sample.x_in + 0.15) * scale
            text_y = (sample.y_in + 0.45) * scale
            lines.append(
                f'<text x="{text_x}" y="{text_y}" font-size="10" fill="black">{sample.sample_id}</text>'
            )
            for cx, cy, dia in sample.drill_centers:
                r = dia * scale / 2
                lines.append(
                    f'<circle cx="{cx * scale}" cy="{cy * scale}" r="{r}" fill="none" stroke="#111827" stroke-width="2" />'
                )

        lines.append('</svg>')
        with open(path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))


class DxfExporter:
    @staticmethod
    def export(path: str, layout: LayoutResult) -> None:
        sheet_layer = layout.metadata.get("sheet_layer", "BOARD")
        label_layer = layout.metadata.get("label_layer", "LABEL")
        configured_layers = {layer.strip() for layer in layout.metadata.get("layers", "").split(",") if layer.strip()}
        if configured_layers:
            layer_names = set(configured_layers)
            layer_names.update({sheet_layer, label_layer})
        else:
            layer_names = {sheet_layer, label_layer, "CUT", "DRILL", "SCRAP"}
        layer_names.update(sample.metadata.get("layer", "CUT") for sample in layout.samples)
        layer_names.update(line.layer for line in layout.line_entities)
        layer_names.update(text.layer for text in layout.text_entities)

        lines = [
            "0", "SECTION",
            "2", "HEADER",
            "9", "$INSUNITS",
            "70", "1",
            "0", "ENDSEC",
            "0", "SECTION",
            "2", "TABLES",
            "0", "TABLE",
            "2", "LAYER",
            "70", str(len(layer_names)),
        ]

        layer_colors = {
            "SHEET": "7",
            "BOARD": "7",
            "MD_ARROW": "2",
            "LABELS": "3",
            "LABEL": "3",
            "FORMED_EDGE": "6",
            "SCRAP": "1",
            "DRILL": "1",
            "CUT": "5",
            "STP315": "5",
            "STP318": "4",
            "STP308": "2",
            "STP308_SCORE": "1",
            "STP312": "5",
            "DTP15": "6",
            "STP311": "4",
            "DTP16": "2",
            "DTP17": "1",
        }

        for name in sorted(layer_names):
            lines.extend([
                "0", "LAYER",
                "2", name,
                "70", "0",
                "62", layer_colors.get(name, "7"),
                "6", "CONTINUOUS",
            ])

        lines.extend([
            "0", "ENDTAB",
            "0", "ENDSEC",
            "0", "SECTION",
            "2", "ENTITIES",
        ])

        DxfExporter._add_rect(lines, sheet_layer, 0, 0, layout.board_width_in, layout.board_height_in, layout.board_height_in)

        for x, y, w, h in layout.scrap_zones:
            DxfExporter._add_rect(lines, "SCRAP", x, y, w, h, layout.board_height_in)

        for sample in layout.samples:
            sample_layer = sample.metadata.get("layer", "CUT")
            DxfExporter._add_rect(
                lines,
                sample_layer,
                sample.x_in,
                sample.y_in,
                sample.width_in,
                sample.height_in,
                layout.board_height_in,
            )
            DxfExporter._add_text(
                lines,
                label_layer,
                sample.x_in + 0.15,
                layout.board_height_in - sample.y_in - 0.45,
                0.25,
                sample.sample_id,
            )
            for cx, cy, dia in sample.drill_centers:
                DxfExporter._add_circle(lines, "DRILL", cx, layout.board_height_in - cy, dia / 2)

        for line in layout.line_entities:
            DxfExporter._add_line(
                lines,
                line.layer,
                line.x1_in,
                layout.board_height_in - line.y1_in,
                line.x2_in,
                layout.board_height_in - line.y2_in,
            )

        for text in layout.text_entities:
            DxfExporter._add_text(
                lines,
                text.layer,
                text.x_in,
                layout.board_height_in - text.y_in,
                text.height_in,
                text.text,
            )

        lines.extend(["0", "ENDSEC", "0", "EOF"])

        with open(path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))

    @staticmethod
    def _add_rect(lines: List[str], layer: str, x: float, y: float, w: float, h: float, board_h: float) -> None:
        left = x
        right = x + w
        top = board_h - y
        bottom = board_h - y - h
        points = [(left, bottom), (right, bottom), (right, top), (left, top)]
        for i, (x1, y1) in enumerate(points):
            x2, y2 = points[(i + 1) % len(points)]
            DxfExporter._add_line(lines, layer, x1, y1, x2, y2)

    @staticmethod
    def _add_line(lines: List[str], layer: str, x1: float, y1: float, x2: float, y2: float) -> None:
        lines.extend([
            "0", "LINE",
            "8", layer,
            "10", f"{x1:.4f}",
            "20", f"{y1:.4f}",
            "30", "0.0",
            "11", f"{x2:.4f}",
            "21", f"{y2:.4f}",
            "31", "0.0",
        ])

    @staticmethod
    def _add_circle(lines: List[str], layer: str, cx: float, cy: float, radius: float) -> None:
        lines.extend([
            "0", "CIRCLE",
            "8", layer,
            "10", f"{cx:.4f}",
            "20", f"{cy:.4f}",
            "30", "0.0",
            "40", f"{radius:.4f}",
        ])

    @staticmethod
    def _add_text(lines: List[str], layer: str, x: float, y: float, height: float, text: str) -> None:
        safe_text = text.replace("\n", " ")
        lines.extend([
            "0", "TEXT",
            "8", layer,
            "10", f"{x:.4f}",
            "20", f"{y:.4f}",
            "30", "0.0",
            "40", f"{height:.4f}",
            "1", safe_text,
        ])


class CsvExporter:
    @staticmethod
    def export(path: str, layout: LayoutResult) -> None:
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([
                "sample_id", "x_in", "y_in", "width_in", "height_in",
                "rotation_deg", "drill_count", "metadata"
            ])
            for sample in layout.samples:
                writer.writerow([
                    sample.sample_id,
                    sample.x_in,
                    sample.y_in,
                    sample.width_in,
                    sample.height_in,
                    sample.rotation_deg,
                    len(sample.drill_centers),
                    json.dumps(sample.metadata),
                ])


class PdfReportExporter:
    @staticmethod
    def export(path: str, layout: LayoutResult, validation_issues: List[Tuple[str, str]]) -> None:
        lines = PdfReportExporter._report_lines(layout, validation_issues)
        content_parts = ["BT", "/F1 10 Tf", "50 760 Td"]
        first = True
        for line in lines[:58]:
            if not first:
                content_parts.append("0 -13 Td")
            content_parts.append(f"({PdfReportExporter._pdf_escape(line)}) Tj")
            first = False
        content_parts.append("ET")
        stream = "\n".join(content_parts).encode("latin-1", errors="replace")
        objects = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream",
        ]
        pdf = bytearray(b"%PDF-1.4\n")
        offsets = []
        for index, obj in enumerate(objects, start=1):
            offsets.append(len(pdf))
            pdf.extend(f"{index} 0 obj\n".encode("ascii"))
            pdf.extend(obj)
            pdf.extend(b"\nendobj\n")
        xref_at = len(pdf)
        pdf.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
        for offset in offsets:
            pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
        pdf.extend(f"trailer << /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_at}\n%%EOF\n".encode("ascii"))
        with open(path, "wb") as f:
            f.write(pdf)

    @staticmethod
    def _report_lines(layout: LayoutResult, validation_issues: List[Tuple[str, str]]) -> List[str]:
        used_area = sum(sample.width_in * sample.height_in for sample in layout.samples)
        board_area = layout.board_width_in * layout.board_height_in
        lines = [
            "GP CNC Builder Job Report",
            f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            f"Sheet: {layout.metadata.get('sheet_name', '')}",
            f"Project: {layout.metadata.get('project_id', '')}",
            f"Product: {layout.metadata.get('product', '')}",
            f"Operator: {layout.metadata.get('operator', '')}",
            f"Board: {layout.board_width_in:g} x {layout.board_height_in:g} in",
            f"Machine Direction: {layout.metadata.get('machine_direction', '')}",
            f"Samples: {len(layout.samples)}",
            f"Used Area: {used_area:.2f} sq in ({(used_area / board_area * 100 if board_area else 0):.1f}%)",
            "",
            "Validation:",
        ]
        if validation_issues:
            lines.extend(f"- {sample_id}: {issue}" for sample_id, issue in validation_issues[:12])
        else:
            lines.append("- No placement issues found.")
        lines.extend(["", "Cut List:"])
        for sample in layout.samples[:35]:
            lines.append(
                f"{sample.sample_id}: X {sample.x_in:.2f}, Y {sample.y_in:.2f}, "
                f"W {sample.width_in:.2f}, H {sample.height_in:.2f}, Layer {sample.metadata.get('layer', 'CUT')}"
            )
        layers = sorted({sample.metadata.get("layer", "CUT") for sample in layout.samples})
        lines.extend(["", f"Layers: {', '.join(layers)}"])
        return lines

    @staticmethod
    def _pdf_escape(text: str) -> str:
        return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


class RuleSheetBuilder:
    REQUIRED_LAYERS = (
        "SHEET",
        "MD_ARROW",
        "LABELS",
        "FORMED_EDGE",
        "STP315",
        "STP318",
        "STP308",
        "STP308_SCORE",
        "STP312",
        "DTP15",
        "STP311",
        "DTP16",
    )

    @classmethod
    def build_all(cls, board_w: float, board_h: float, machine_direction: str) -> List[LayoutResult]:
        return [
            cls.build_board1(board_w, board_h, machine_direction),
            cls.build_board2(board_w, board_h, machine_direction),
            cls.build_board3(board_w, board_h, machine_direction),
        ]

    @classmethod
    def build_board1(cls, board_w: float, board_h: float, machine_direction: str) -> LayoutResult:
        layout = cls._new_layout(board_w, board_h, "board1", machine_direction)

        stp315_w, stp315_h = cls._md_oriented_size(machine_direction, 12, 24)
        stp315_positions = []
        if 12 + 4 * stp315_h + 3 <= board_h - 12:
            x = (board_w - stp315_w) / 2
            stp315_positions = [(x, 12 + i * (stp315_h + 1)) for i in range(4)]
        else:
            group_w = 2 * stp315_w + 1
            group_h = 2 * stp315_h + 1
            if group_w > board_w or 12 + group_h > board_h - 12:
                raise ValueError("STP315 parts do not fit with the selected board size and machine direction.")
            start_x = (board_w - group_w) / 2
            stp315_positions = [
                (start_x + col * (stp315_w + 1), 12 + row * (stp315_h + 1))
                for row in range(2)
                for col in range(2)
            ]

        stp315_centered = len({round(pos[0], 4) for pos in stp315_positions}) == 1
        for i, (x, y) in enumerate(stp315_positions):
            metadata = {"min_end": "12"}
            if stp315_centered:
                metadata["centered_width"] = "true"
            cls._add_part(
                layout,
                "STP315",
                i + 1,
                x,
                y,
                stp315_w,
                stp315_h,
                "STP315",
                metadata,
            )

        if board_h >= board_w:
            edge_specs = [(SHOP_EDGE_INSET_IN, 6), (board_w - SHOP_EDGE_INSET_IN - 6, 6)]
            shear_specs = [(0, 4), (board_w - 4, 4)]
        else:
            edge_specs = [(6, SHOP_EDGE_INSET_IN), (6, board_h - SHOP_EDGE_INSET_IN - 4.25)]
            shear_specs = [(4, 0), (4, board_h - 4)]

        for edge_index, (x0, y0) in enumerate(edge_specs):
            for i in range(5):
                part_x = x0
                part_y = y0 + i * 5.25 if board_h >= board_w else y0
                if board_h < board_w:
                    part_x = x0 + i * 7
                sample = cls._add_part(
                    layout,
                    "STP308",
                    edge_index * 5 + i + 1,
                    part_x,
                    part_y,
                    6,
                    4.25,
                    "STP308",
                    {"min_margin": "2"},
                )
                score_y = sample.y_in + 1.25
                layout.line_entities.append(LineEntity("STP308_SCORE", sample.x_in, score_y, sample.x_in + 6, score_y))

        for edge_index, (x0, width) in enumerate(shear_specs):
            for i in range(3):
                y = max(6, board_h - 6 - (3 - i) * 7)
                cls._add_part(
                    layout,
                    "STP318",
                    edge_index * 3 + i + 1,
                    x0,
                    y,
                    width,
                    6,
                    "STP318",
                    {"formed_edge": "true", "min_end": "6"},
                )

        layout.warnings.append("Board 1 rule sheet: STP315, STP318, and STP308 placed from manual board dimensions. Only STP318 remains on the formed edge.")
        cls._finalize(layout)
        return layout

    @classmethod
    def build_board2(cls, board_w: float, board_h: float, machine_direction: str) -> LayoutResult:
        layout = cls._new_layout(board_w, board_h, "board2", machine_direction)

        cls._add_grid(layout, "STP312-CD", "STP312", 5, 4, 4, 12, 16, 1, {"min_margin": "4"})
        cls._add_grid(layout, "STP312-MD", "STP312", 5, 4, 38, 16, 12, 1, {"min_margin": "4"})

        dtp_w, dtp_h = cls._md_oriented_size(machine_direction, 14, 3.5)
        if machine_direction == "Horizontal":
            cols = 3
            start_x = max(SHOP_EDGE_INSET_IN, board_w - SHOP_EDGE_INSET_IN - cols * dtp_w - (cols - 1) * 1)
            start_y = board_h - 12
            for i in range(6):
                col = i % cols
                row = i // cols
                x = start_x + col * (dtp_w + 1)
                y = start_y + row * (dtp_h + 1)
                cls._add_part(layout, "DTP15", i + 1, x, y, dtp_w, dtp_h, "DTP15")
        else:
            x = max(SHOP_EDGE_INSET_IN, board_w - dtp_w - SHOP_EDGE_INSET_IN)
            for i in range(6):
                y = 4 + i * (dtp_h + 0.5)
                cls._add_part(layout, "DTP15", i + 1, x, y, dtp_w, dtp_h, "DTP15")

        layout.warnings.append("Board 2 rule sheet: STP312 and DTP-15 placed from manual board dimensions.")
        cls._finalize(layout)
        return layout

    @classmethod
    def build_board3(cls, board_w: float, board_h: float, machine_direction: str) -> LayoutResult:
        layout = cls._new_layout(board_w, board_h, "board3", machine_direction)

        cls._add_grid(layout, "STP311", "STP311", 10, SHOP_EDGE_INSET_IN, SHOP_EDGE_INSET_IN, 6, 6, 1)
        cls._add_grid(layout, "DTP16", "DTP16", 10, 4, 20, 6, 6, 1, {"min_margin": "4"})

        layout.warnings.append("Board 3 rule sheet: STP311 and DTP-16 placed from manual board dimensions.")
        cls._finalize(layout)
        return layout

    @classmethod
    def _new_layout(cls, board_w: float, board_h: float, sheet_name: str, machine_direction: str) -> LayoutResult:
        if board_w <= 0 or board_h <= 0:
            raise ValueError("Board dimensions must be greater than zero.")
        layout = LayoutResult(
            board_width_in=board_w,
            board_height_in=board_h,
            dtp_id="RULE-SHEET",
            samples=[],
            metadata={
                "sheet_name": sheet_name,
                "sheet_layer": "SHEET",
                "label_layer": "LABELS",
                "layers": ",".join(cls.REQUIRED_LAYERS),
                "machine_direction": machine_direction,
            },
        )
        cls._add_formed_edges(layout)
        cls._add_md_arrow(layout, machine_direction)
        return layout

    @staticmethod
    def _md_oriented_size(machine_direction: str, md_len: float, cross_len: float) -> Tuple[float, float]:
        if machine_direction == "Vertical":
            return cross_len, md_len
        return md_len, cross_len

    @staticmethod
    def _add_part(
        layout: LayoutResult,
        prefix: str,
        index: int,
        x: float,
        y: float,
        width: float,
        height: float,
        layer: str,
        extra_metadata: Optional[Dict[str, str]] = None,
    ) -> SampleRect:
        sample = SampleRect(
            sample_id=f"{prefix}-{index}",
            x_in=x,
            y_in=y,
            width_in=width,
            height_in=height,
            metadata={"layer": layer, **(extra_metadata or {})},
        )
        layout.samples.append(sample)
        return sample

    @classmethod
    def _add_grid(
        cls,
        layout: LayoutResult,
        prefix: str,
        layer: str,
        quantity: int,
        start_x: float,
        start_y: float,
        width: float,
        height: float,
        gap: float,
        extra_metadata: Optional[Dict[str, str]] = None,
    ) -> None:
        usable_w = max(width, layout.board_width_in - start_x - 4)
        cols = max(1, int((usable_w + gap) // (width + gap)))
        for i in range(quantity):
            col = i % cols
            row = i // cols
            x = start_x + col * (width + gap)
            y = start_y + row * (height + gap)
            cls._add_part(layout, prefix, i + 1, x, y, width, height, layer, extra_metadata)

    @staticmethod
    def _add_formed_edges(layout: LayoutResult) -> None:
        if layout.board_height_in >= layout.board_width_in:
            layout.line_entities.append(LineEntity("FORMED_EDGE", 0, 0, 0, layout.board_height_in))
            layout.line_entities.append(LineEntity("FORMED_EDGE", layout.board_width_in, 0, layout.board_width_in, layout.board_height_in))
            layout.text_entities.append(TextEntity("LABELS", 1, 3, 0.3, "FORMED EDGE"))
            layout.text_entities.append(TextEntity("LABELS", layout.board_width_in - 7, 3, 0.3, "FORMED EDGE"))
        else:
            layout.line_entities.append(LineEntity("FORMED_EDGE", 0, 0, layout.board_width_in, 0))
            layout.line_entities.append(LineEntity("FORMED_EDGE", 0, layout.board_height_in, layout.board_width_in, layout.board_height_in))
            layout.text_entities.append(TextEntity("LABELS", 3, 1, 0.3, "FORMED EDGE"))
            layout.text_entities.append(TextEntity("LABELS", 3, layout.board_height_in - 1, 0.3, "FORMED EDGE"))

    @staticmethod
    def _add_md_arrow(layout: LayoutResult, machine_direction: str) -> None:
        if machine_direction == "Vertical":
            x = layout.board_width_in + 2
            y_end = max(8, layout.board_height_in - 8)
            layout.line_entities.append(LineEntity("MD_ARROW", x, y_end, x, 8))
            layout.line_entities.append(LineEntity("MD_ARROW", x, 8, x - 1.5, 11))
            layout.line_entities.append(LineEntity("MD_ARROW", x, 8, x + 1.5, 11))
            layout.text_entities.append(TextEntity("MD_ARROW", x + 1, 10, 0.45, "MD"))
        else:
            y = layout.board_height_in + 2
            x_end = max(8, layout.board_width_in - 8)
            layout.line_entities.append(LineEntity("MD_ARROW", 8, y, x_end, y))
            layout.line_entities.append(LineEntity("MD_ARROW", x_end, y, x_end - 3, y - 1.5))
            layout.line_entities.append(LineEntity("MD_ARROW", x_end, y, x_end - 3, y + 1.5))
            layout.text_entities.append(TextEntity("MD_ARROW", max(1, x_end - 4), y + 1, 0.45, "MD"))

    @classmethod
    def _finalize(cls, layout: LayoutResult) -> None:
        cls._validate_bounds_and_spacing(layout)
        layout.text_entities.append(TextEntity("LABELS", 1, layout.board_height_in - 1, 0.35, layout.metadata.get("sheet_name", "sheet")))

    @staticmethod
    def _validate_bounds_and_spacing(layout: LayoutResult) -> None:
        for sample in layout.samples:
            if sample.x_in < -1e-6 or sample.y_in < -1e-6:
                raise ValueError(f"{sample.sample_id} is outside the board.")
            if sample.x_in + sample.width_in > layout.board_width_in + 1e-6:
                raise ValueError(f"{sample.sample_id} does not fit within the board width.")
            if sample.y_in + sample.height_in > layout.board_height_in + 1e-6:
                raise ValueError(f"{sample.sample_id} does not fit within the board height.")

        for i, sample in enumerate(layout.samples):
            for other in layout.samples[i + 1:]:
                overlap = (
                    sample.x_in < other.x_in + other.width_in - 1e-6
                    and sample.x_in + sample.width_in > other.x_in + 1e-6
                    and sample.y_in < other.y_in + other.height_in - 1e-6
                    and sample.y_in + sample.height_in > other.y_in + 1e-6
                )
                if overlap:
                    raise ValueError(f"{sample.sample_id} overlaps {other.sample_id}.")


class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=10)
        self.master = master
        self.engine = RuleEngine(RULES_JSON)
        self.layout_result: Optional[LayoutResult] = None
        self.selected_sample_id: Optional[str] = None
        self.selected_sample_ids: Set[str] = set()
        self.drag_sample_id: Optional[str] = None
        self.drag_last_xy: Optional[Tuple[int, int]] = None
        self.last_drag_error: Optional[str] = None
        self.marquee_start_canvas: Optional[Tuple[float, float]] = None
        self.marquee_rect_id: Optional[int] = None
        self.marquee_dragging = False
        self.canvas_scale = 1.0
        self.canvas_origin: Tuple[float, float] = (0.0, 0.0)
        self.view_zoom = 1.0
        self.measure_mode_var = tk.BooleanVar(value=False)
        self.highlight_mode_var = tk.BooleanVar(value=True)
        self.measure_start: Optional[Tuple[float, float]] = None
        self.updating_part_tree = False
        self.updating_sheet_tabs = False
        self.panning_canvas = False
        self.saved_sheets: List[LayoutResult] = []
        self.undo_stack: List[Dict] = []
        self.redo_stack: List[Dict] = []
        self.drag_undo_recorded = False
        self.show_grid_var = tk.BooleanVar(value=True)
        self.show_labels_var = tk.BooleanVar(value=True)
        self.snap_enabled_var = tk.BooleanVar(value=True)
        self.snap_increment_var = tk.StringVar(value="0.25")
        self.board_preset_var = tk.StringVar(value="ToughRock 48 x 96 x 5/8")
        self.selected_x_var = tk.StringVar()
        self.selected_y_var = tk.StringVar()
        self.selected_w_var = tk.StringVar()
        self.selected_h_var = tk.StringVar()
        self.edge_side_var = tk.StringVar(value="Left")
        self.edge_distance_var = tk.StringVar(value="1")
        self.selection_count_var = tk.StringVar(value="Selected: 0")
        self.sample_filter_var = tk.StringVar()
        self.manual_object_var = tk.StringVar(value="DTP-15")
        self.manual_qty_var = tk.StringVar(value="1")
        self.custom_w_var = tk.StringVar(value="6")
        self.custom_h_var = tk.StringVar(value="6")
        self.custom_label_var = tk.StringVar(value="CUSTOM")
        self.bot_prompt_var = tk.StringVar()
        self.usage_pct_var = tk.DoubleVar(value=0.0)
        self.usage_used_var = tk.StringVar(value="Used: --")
        self.usage_waste_var = tk.StringVar(value="Waste: --")
        self.usage_parts_var = tk.StringVar(value="Parts: --")
        self.usage_board_var = tk.StringVar(value="Board: --")
        self.bot_status_var = tk.StringVar(value="Ready. Type a request and press Enter / Build Board.")
        self.bot_placeholder_text = "Type here..."
        self.bot_placeholder_active = False
        self.bot_log_height = 0
        self.bot_input_height = 9
        self.insert_text_var = tk.StringVar(value="NOTE")
        self.insert_text_x_var = tk.StringVar(value="1")
        self.insert_text_y_var = tk.StringVar(value="1")
        self.insert_text_height_var = tk.StringVar(value="0.35")
        self.app_icon_image: Optional[tk.PhotoImage] = None
        self.header_logo_image: Optional[tk.PhotoImage] = None
        self.login_logo_image = None
        self.setup_preview_dtp_id = "DTP-11"
        self.current_user_email = ""
        self.profile_dir = ""
        self.profile_path = ""
        self.last_export_folder = ""

        master.title("GP CNC Builder")
        master.report_callback_exception = self.report_callback_exception
        self.load_brand_assets()
        screen_w = master.winfo_screenwidth()
        screen_h = master.winfo_screenheight()
        start_w = min(1760, max(1200, int(screen_w * 0.9)))
        start_h = min(960, max(720, int(screen_h * 0.88)))
        master.geometry(f"{start_w}x{start_h}")
        master.minsize(1050, 680)
        master.resizable(True, True)
        self.prompt_for_login()
        if master.state() != "withdrawn":
            try:
                master.state("zoomed")
            except tk.TclError:
                pass
        self.pack(fill=tk.BOTH, expand=True)

        self._build_ui()
        self._bind_shortcuts()
        self.load_user_profile()

    def load_brand_assets(self) -> None:
        if not os.path.exists(GP_LOGO_PNG):
            return
        try:
            self.app_icon_image = tk.PhotoImage(file=GP_LOGO_PNG)
            self.master.iconphoto(True, self.app_icon_image)
            self.header_logo_image = self.app_icon_image.subsample(5, 5)
        except tk.TclError:
            self.app_icon_image = None
            self.header_logo_image = None

    def prompt_for_login(self) -> None:
        if self.master.state() == "withdrawn":
            self.set_current_user("guest@local.profile")
            return

        dialog = tk.Toplevel(self.master)
        dialog.title("GP CNC Builder Sign In")
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()

        container = ttk.Frame(dialog, padding=18)
        container.pack(fill=tk.BOTH, expand=True)

        logo_loaded = self._load_login_logo()
        if logo_loaded:
            tk.Label(container, image=self.login_logo_image, borderwidth=0).pack(pady=(0, 12))

        ttk.Label(container, text="GP CNC Builder", font=("Segoe UI", 15, "bold")).pack()
        ttk.Label(container, text="Sign in with your email to load your saved sheets.").pack(pady=(4, 14))

        email_var = tk.StringVar()
        status_var = tk.StringVar()
        entry = ttk.Entry(container, textvariable=email_var, width=36)
        entry.pack(fill=tk.X)
        ttk.Label(container, textvariable=status_var, foreground="#b91c1c").pack(anchor="w", pady=(6, 8))

        buttons = ttk.Frame(container)
        buttons.pack(fill=tk.X)

        def finish_with_email() -> None:
            email = email_var.get().strip().lower()
            if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
                status_var.set("Enter a valid email address.")
                entry.focus_set()
                return
            self.set_current_user(email)
            dialog.destroy()

        def use_guest() -> None:
            self.set_current_user("guest@local.profile")
            dialog.destroy()

        ttk.Button(buttons, text="Sign In", command=finish_with_email).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Continue as Guest", command=use_guest).pack(side=tk.LEFT, padx=(8, 0))
        dialog.bind("<Return>", lambda _event: finish_with_email())
        dialog.protocol("WM_DELETE_WINDOW", use_guest)

        dialog.update_idletasks()
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        x = self.master.winfo_screenwidth() // 2 - width // 2
        y = self.master.winfo_screenheight() // 2 - height // 2
        dialog.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")
        entry.focus_set()
        self.master.wait_window(dialog)

    def _load_login_logo(self) -> bool:
        if Image is not None and ImageTk is not None and os.path.exists(GP_LOGIN_LOGO_JPG):
            try:
                image = Image.open(GP_LOGIN_LOGO_JPG)
                image.thumbnail((260, 130))
                self.login_logo_image = ImageTk.PhotoImage(image)
                return True
            except Exception:
                self.login_logo_image = None

        if os.path.exists(GP_LOGO_PNG):
            try:
                image = tk.PhotoImage(file=GP_LOGO_PNG)
                self.login_logo_image = image.subsample(3, 3)
                return True
            except tk.TclError:
                self.login_logo_image = None
        return False

    def set_current_user(self, email: str) -> None:
        self.current_user_email = email.strip().lower() or "guest@local.profile"
        profile_id = self._safe_profile_id(self.current_user_email)
        self.profile_dir = os.path.join(PROFILE_ROOT, profile_id)
        self.profile_path = os.path.join(self.profile_dir, "sheets.json")

    @staticmethod
    def _safe_profile_id(email: str) -> str:
        profile_id = re.sub(r"[^a-z0-9_.-]+", "_", email.lower()).strip("._-")
        return profile_id or "guest"

    def report_callback_exception(self, exc_type, exc_value, exc_traceback) -> None:
        details = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        try:
            with open("gp_cnc_builder_error.log", "w", encoding="utf-8") as f:
                f.write(details)
        except OSError:
            pass
        messagebox.showerror(
            "GP CNC Builder Error",
            f"{exc_value}\n\nDetails were written to gp_cnc_builder_error.log.",
        )

    def _bind_shortcuts(self) -> None:
        self.master.bind_all("<Control-z>", lambda _event: self.undo())
        self.master.bind_all("<Control-y>", lambda _event: self.redo())
        self.master.bind_all("<Control-s>", lambda _event: self.save_current_sheet())
        self.master.bind_all("<Control-e>", lambda _event: self.export_dxf())
        self.master.bind_all("<Delete>", self._delete_shortcut)
        self.master.bind_all("<Control-d>", lambda _event: self.duplicate_selected_sample())

    def _delete_shortcut(self, _event=None) -> None:
        focus = self.master.focus_get()
        if isinstance(focus, (tk.Entry, tk.Text, ttk.Entry, ttk.Combobox)):
            return
        self.delete_selected_sample()

    def _build_ui(self) -> None:
        left_shell = ttk.Frame(self, width=370)
        center = ttk.Frame(self)
        right_shell = ttk.Frame(self, width=340)

        left_shell.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_shell.pack_propagate(False)
        center.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        right_shell.pack(side=tk.LEFT, fill=tk.Y)
        right_shell.pack_propagate(False)

        left = self._make_scrollable_panel(left_shell, bind_child_wheel=True)
        right = self._make_scrollable_panel(right_shell, bind_child_wheel=True)

        self._build_brand_header(left)
        self._build_project_panel(left)
        self._build_preview_panel(center)
        self._build_setup_panel(right)

    def _make_scrollable_panel(self, parent: ttk.Frame, bind_child_wheel: bool = False) -> ttk.Frame:
        canvas = tk.Canvas(parent, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        content = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def update_scrollregion(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def sync_width(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        content.bind("<Configure>", update_scrollregion)
        canvas.bind("<Configure>", sync_width)
        self._bind_panel_mousewheel(canvas, content, bind_child_wheel=bind_child_wheel)
        return content

    def _bind_panel_mousewheel(self, canvas: tk.Canvas, content: ttk.Frame, bind_child_wheel: bool = False) -> None:
        def on_mousewheel(event) -> None:
            delta = getattr(event, "delta", 0)
            if delta == 0:
                button_num = getattr(event, "num", None)
                if button_num == 4:
                    delta = 120
                elif button_num == 5:
                    delta = -120
                else:
                    return
            canvas.yview_scroll(-1 if delta > 0 else 1, "units")
            return "break"

        for widget in (canvas, content):
            widget.bind("<MouseWheel>", on_mousewheel, add="+")
            widget.bind("<Button-4>", on_mousewheel, add="+")
            widget.bind("<Button-5>", on_mousewheel, add="+")

        if bind_child_wheel:
            def bind_children(_event=None) -> None:
                pending = list(content.winfo_children())
                while pending:
                    child = pending.pop()
                    if not getattr(child, "_panel_mousewheel_bound", False):
                        child.bind("<MouseWheel>", on_mousewheel, add="+")
                        child.bind("<Button-4>", on_mousewheel, add="+")
                        child.bind("<Button-5>", on_mousewheel, add="+")
                        child._panel_mousewheel_bound = True
                    pending.extend(child.winfo_children())

            content.bind("<Configure>", bind_children, add="+")

    def _bind_widget_mousewheel(self, widget) -> None:
        def on_mousewheel(event) -> str:
            delta = getattr(event, "delta", 0)
            if delta == 0:
                button_num = getattr(event, "num", None)
                if button_num == 4:
                    delta = 120
                elif button_num == 5:
                    delta = -120
                else:
                    return "break"
            widget.yview_scroll(-1 if delta > 0 else 1, "units")
            return "break"

        widget.bind("<MouseWheel>", on_mousewheel, add="+")
        widget.bind("<Button-4>", on_mousewheel, add="+")
        widget.bind("<Button-5>", on_mousewheel, add="+")

    def _build_brand_header(self, parent: ttk.Frame) -> None:
        brand = ttk.Frame(parent)
        brand.pack(fill=tk.X, pady=(0, 10))
        if self.header_logo_image:
            tk.Label(brand, image=self.header_logo_image, borderwidth=0).pack(side=tk.LEFT, padx=(0, 8))
        text_frame = ttk.Frame(brand)
        text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(text_frame, text="GP CNC Builder", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(text_frame, text="Georgia-Pacific sample cut planning", font=("Segoe UI", 8)).pack(anchor="w")
        ttk.Label(text_frame, text=f"Signed in: {self.current_user_email}", font=("Segoe UI", 8)).pack(anchor="w")

    def _build_project_panel(self, parent: ttk.Frame) -> None:
        project = ttk.LabelFrame(parent, text="Project Setup", padding=10)
        project.pack(fill=tk.X, pady=(0, 10))

        self.operator_var = tk.StringVar()
        self.project_var = tk.StringVar()
        self.product_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar(value="Sheet 1")
        self.board_w_var = tk.StringVar(value="48")
        self.board_h_var = tk.StringVar(value="96")
        self.thickness_var = tk.StringVar(value="0.625")
        self.md_var = tk.StringVar(value="Horizontal")
        self.code_side_var = tk.StringVar(value="Yes")
        self.dtp_var = tk.StringVar(value="DTP-11")
        self.quantity_var = tk.StringVar(value=str(RULES_JSON["DTP-11"]["quantity_default"]))
        self.orientation_var = tk.StringVar(value="Face Up")
        self.hole_size_var = tk.StringVar(value="0.25")
        self.additional_dtp_vars: List[tk.StringVar] = []
        self.additional_qty_vars: List[tk.StringVar] = []
        self.additional_dtp_combos: List[ttk.Combobox] = []

        preset_row = ttk.Frame(project)
        preset_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Combobox(
            preset_row,
            textvariable=self.board_preset_var,
            values=list(PRODUCT_PRESETS.keys()),
            state="readonly",
            width=24,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(preset_row, text="Apply Preset", command=self.apply_board_preset).pack(side=tk.LEFT, padx=(6, 0))

        rows = [
            ("Operator", self.operator_var),
            ("Project ID", self.project_var),
            ("Product", self.product_var),
            ("Sheet Name", self.sheet_name_var),
            ("Board Width (in)", self.board_w_var),
            ("Board Height (in)", self.board_h_var),
            ("Thickness (in)", self.thickness_var),
        ]
        for i, (label, var) in enumerate(rows):
            row_index = i + 1
            ttk.Label(project, text=label).grid(row=row_index, column=0, sticky="w", pady=2)
            if label == "Product":
                ttk.Combobox(
                    project,
                    textvariable=var,
                    values=["ToughRock", "DensGlass", "DensShield", "DensDeck Prime", "Other"],
                    width=18,
                ).grid(row=row_index, column=1, sticky="ew", pady=2)
            else:
                ttk.Entry(project, textvariable=var, width=18).grid(row=row_index, column=1, sticky="ew", pady=2)

        ttk.Label(project, text="Machine Direction").grid(row=8, column=0, sticky="w", pady=2)
        ttk.Combobox(project, textvariable=self.md_var, values=["Horizontal", "Vertical"], state="readonly", width=15).grid(row=8, column=1, sticky="ew", pady=2)

        ttk.Label(project, text="Code Side").grid(row=9, column=0, sticky="w", pady=2)
        ttk.Combobox(project, textvariable=self.code_side_var, values=["Yes", "No"], state="readonly", width=15).grid(row=9, column=1, sticky="ew", pady=2)
        project.columnconfigure(1, weight=1)

        select = ttk.LabelFrame(parent, text="Sample Requests", padding=10)
        select.pack(fill=tk.X, pady=(0, 10))

        self.dtp_values_all = [f"{k} - {v['test_name']}" for k, v in RULES_JSON.items()]
        dtp_values = self.dtp_values_all[:]
        ttk.Label(select, text="Search").grid(row=0, column=0, sticky="w", pady=(0, 3))
        search_row = ttk.Frame(select)
        search_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Entry(search_row, textvariable=self.sample_filter_var, width=20).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(search_row, text="Filter", command=self.filter_sample_types).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(select, text="Sample Type").grid(row=2, column=0, sticky="w", pady=(0, 3))
        ttk.Label(select, text="Qty").grid(row=2, column=1, sticky="w", padx=(6, 0), pady=(0, 3))

        self.dtp_combo = ttk.Combobox(select, state="readonly", values=dtp_values, width=34)
        self.dtp_combo.current(0)
        self.dtp_combo.grid(row=3, column=0, sticky="ew", pady=2)
        ttk.Entry(select, textvariable=self.quantity_var, width=7).grid(row=3, column=1, sticky="ew", padx=(6, 0), pady=2)
        self.dtp_combo.bind("<<ComboboxSelected>>", self._on_dtp_changed)
        select.columnconfigure(0, weight=1)

        for i in range(SAMPLE_REQUEST_ROWS - 1):
            dtp_var = tk.StringVar(value="None")
            qty_var = tk.StringVar(value="0")
            self.additional_dtp_vars.append(dtp_var)
            self.additional_qty_vars.append(qty_var)
            combo = ttk.Combobox(select, textvariable=dtp_var, state="readonly", values=["None"] + dtp_values, width=34)
            self.additional_dtp_combos.append(combo)
            combo.grid(row=i + 4, column=0, sticky="ew", pady=2)
            combo.bind("<<ComboboxSelected>>", lambda _event, index=i: self._on_additional_dtp_changed(index))
            ttk.Entry(select, textvariable=qty_var, width=7).grid(row=i + 4, column=1, sticky="ew", padx=(6, 0), pady=2)

        params = ttk.LabelFrame(parent, text="Dynamic Parameters", padding=10)
        params.pack(fill=tk.X)
        self.params_frame = params
        self._rebuild_params()

        manual = ttk.LabelFrame(parent, text="Add One Part", padding=10)
        manual.pack(fill=tk.X, pady=(10, 0))
        self.manual_object_combo = ttk.Combobox(
            manual,
            textvariable=self.manual_object_var,
            values=list(MANUAL_OBJECTS.keys()) + ["CUSTOM"],
            state="readonly",
            width=18,
        )
        self.manual_object_combo.grid(row=0, column=0, columnspan=2, sticky="ew", pady=2)
        self.manual_object_combo.bind("<<ComboboxSelected>>", self._on_manual_object_changed)
        ttk.Label(manual, text="Qty").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(manual, textvariable=self.manual_qty_var, width=8).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Label(manual, text="Custom Label").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(manual, textvariable=self.custom_label_var, width=10).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(manual, text="W x H").grid(row=3, column=0, sticky="w", pady=2)
        size_row = ttk.Frame(manual)
        size_row.grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Entry(size_row, textvariable=self.custom_w_var, width=5).pack(side=tk.LEFT)
        ttk.Label(size_row, text=" x ").pack(side=tk.LEFT)
        ttk.Entry(size_row, textvariable=self.custom_h_var, width=5).pack(side=tk.LEFT)
        ttk.Button(manual, text="Add To Sheet", command=self.add_manual_parts).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        manual.columnconfigure(1, weight=1)

        buttons = ttk.Frame(parent)
        buttons.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(buttons, text="Generate Layout", command=self.generate_layout).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Fill Board", command=self.fill_board_current_layout).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Save Current Sheet", command=self.save_current_sheet).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Build 3 Rule Sheets", command=self.build_rule_sheet_set).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Export SVG", command=self.export_svg).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Export CSV", command=self.export_csv).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Export DXF", command=self.export_dxf).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Export Saved DXFs", command=self.export_saved_dxfs).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Save Project", command=self.save_project_file).pack(fill=tk.X, pady=3)
        ttk.Button(buttons, text="Load Project", command=self.load_project_file).pack(fill=tk.X, pady=3)

    def _rebuild_params(self) -> None:
        for child in self.params_frame.winfo_children():
            child.destroy()

        dtp_id = self.current_dtp_id
        rule = RULES_JSON[dtp_id]

        ttk.Label(self.params_frame, text=f"{dtp_id}: {rule['test_name']}").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(self.params_frame, text="Quantities are set in the sample request rows above.").grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 6),
        )

        row = 2
        if dtp_id == "DTP-13":
            ttk.Label(self.params_frame, text="Orientation").grid(row=row, column=0, sticky="w", pady=2)
            ttk.Combobox(
                self.params_frame,
                textvariable=self.orientation_var,
                values=rule.get("orientation_options", ["Face Up"]),
                state="readonly",
                width=14,
            ).grid(row=row, column=1, sticky="ew", pady=2)
            row += 1

            ttk.Label(self.params_frame, text="Hole Size (in)").grid(row=row, column=0, sticky="w", pady=2)
            ttk.Entry(self.params_frame, textvariable=self.hole_size_var, width=16).grid(row=row, column=1, sticky="ew", pady=2)
            row += 1

        ttk.Label(self.params_frame, text="Notes").grid(row=row, column=0, sticky="nw", pady=(8, 2))
        note = tk.Text(self.params_frame, height=5, width=28, wrap="word")
        note.grid(row=row, column=1, sticky="ew", pady=(8, 2))
        note.insert("1.0", rule.get("notes", ""))
        note.configure(state="disabled")

    @property
    def current_dtp_id(self) -> str:
        selected = self.dtp_combo.get().strip()
        return selected.split(" - ", 1)[0]

    def _on_dtp_changed(self, _event=None) -> None:
        default_qty = RULES_JSON[self.current_dtp_id].get("quantity_default", 1)
        self.quantity_var.set(str(default_qty))
        self.setup_preview_dtp_id = self.current_dtp_id
        self._rebuild_params()
        self.update_procedure_steps()

    def _on_additional_dtp_changed(self, index: int) -> None:
        selected = self.additional_dtp_vars[index].get()
        if selected == "None":
            self.additional_qty_vars[index].set("0")
        else:
            dtp_id = selected.split(" - ", 1)[0]
            default_qty = RULES_JSON[dtp_id].get("quantity_default", 1)
            self.additional_qty_vars[index].set(str(default_qty))
            self.setup_preview_dtp_id = dtp_id
        self.update_procedure_steps()

    def _active_dtp_ids(self) -> List[str]:
        dtp_ids = [self.current_dtp_id]
        for dtp_var in self.additional_dtp_vars:
            selected = dtp_var.get()
            if selected and selected != "None":
                dtp_id = selected.split(" - ", 1)[0]
                if dtp_id not in dtp_ids:
                    dtp_ids.append(dtp_id)
        return dtp_ids

    def _collect_layout_requests(self) -> List[LayoutRequest]:
        requests = [
            LayoutRequest(
                dtp_id=self.current_dtp_id,
                quantity=int(self.quantity_var.get()),
                request_index=1,
            )
        ]

        for index, (dtp_var, qty_var) in enumerate(zip(self.additional_dtp_vars, self.additional_qty_vars), start=2):
            selected = dtp_var.get()
            if not selected or selected == "None":
                continue
            quantity = int(qty_var.get())
            if quantity <= 0:
                continue
            requests.append(
                LayoutRequest(
                    dtp_id=selected.split(" - ", 1)[0],
                    quantity=quantity,
                    request_index=index,
                )
            )

        return requests

    def apply_board_preset(self) -> None:
        preset = PRODUCT_PRESETS.get(self.board_preset_var.get())
        if not preset:
            return
        self.product_var.set(preset["product"])
        self.board_w_var.set(preset["width"])
        self.board_h_var.set(preset["height"])
        self.thickness_var.set(preset["thickness"])

    def filter_sample_types(self) -> None:
        query = self.sample_filter_var.get().strip().lower()
        values = [value for value in self.dtp_values_all if query in value.lower()] if query else self.dtp_values_all[:]
        if not values:
            values = self.dtp_values_all[:]
        self.dtp_combo.configure(values=values)
        for combo in self.additional_dtp_combos:
            combo.configure(values=["None"] + values)

    def _sync_project_fields_from_layout(self, layout: LayoutResult) -> None:
        self.sheet_name_var.set(layout.metadata.get("sheet_name", self.sheet_name_var.get()))
        self.product_var.set(layout.metadata.get("product", self.product_var.get()))
        self.project_var.set(layout.metadata.get("project_id", self.project_var.get()))
        self.operator_var.set(layout.metadata.get("operator", self.operator_var.get()))
        self.md_var.set(layout.metadata.get("machine_direction", self.md_var.get()))

    def _on_manual_object_changed(self, _event=None) -> None:
        selected = self.manual_object_var.get()
        if selected in MANUAL_OBJECTS:
            width, height = self._manual_object_size(selected)
            self.custom_label_var.set(selected)
            self.custom_w_var.set(f"{width:g}")
            self.custom_h_var.set(f"{height:g}")
            preview_id = selected
            if selected.startswith("STP312"):
                preview_id = "STP312"
            elif selected == "DTP16":
                preview_id = "DTP-16"
            if preview_id in RULES_JSON:
                self.setup_preview_dtp_id = preview_id
                self.draw_animation_placeholder()

    def _ensure_working_layout(self) -> LayoutResult:
        if self.layout_result:
            return self.layout_result
        board_width = float(self.board_w_var.get())
        board_height = float(self.board_h_var.get())
        self.layout_result = LayoutResult(
            board_width_in=board_width,
            board_height_in=board_height,
            dtp_id="MANUAL",
            samples=[],
            metadata={
                "sheet_name": self.sheet_name_var.get().strip() or "manual_sheet",
                "sheet_layer": "SHEET",
                "label_layer": "LABELS",
                "machine_direction": self.md_var.get(),
            },
        )
        RuleSheetBuilder._add_md_arrow(self.layout_result, self.md_var.get())
        self._stamp_layout_metadata(self.layout_result)
        return self.layout_result

    def _manual_object_size(self, object_key: str) -> Tuple[float, float]:
        object_key = self._normalize_manual_object_key(object_key)
        if object_key == "CUSTOM":
            return float(self.custom_w_var.get()), float(self.custom_h_var.get())
        if object_key == "DTP-17":
            return float(self.board_w_var.get()), float(self.board_h_var.get())
        spec = MANUAL_OBJECTS[object_key]
        if "md_length" in spec:
            return RuleSheetBuilder._md_oriented_size(
                self.md_var.get(),
                float(spec["md_length"]),
                float(spec["cross_length"]),
            )
        return float(spec["width"]), float(spec["height"])

    def _make_manual_sample(self, object_key: str, index: int) -> SampleRect:
        original_key = object_key
        object_key = self._normalize_manual_object_key(object_key)
        width, height = self._manual_object_size(object_key)
        if object_key == "CUSTOM":
            label = self.custom_label_var.get().strip() or "CUSTOM"
            layer = "CUT"
            metadata = {"layer": layer, "manual": "true"}
        else:
            spec = MANUAL_OBJECTS[object_key]
            label = original_key
            layer = str(spec.get("layer", "CUT"))
            metadata = {"layer": layer, "manual": "true"}
            metadata.update(spec.get("metadata", {}))

        sample = SampleRect(
            sample_id=f"{label}-{self._next_part_number(label)}",
            x_in=0,
            y_in=0,
            width_in=width,
            height_in=height,
            metadata=metadata,
        )
        if object_key == "DTP-13":
            diameter = float(MANUAL_OBJECTS["DTP-13"].get("drill_diameter", 0.25))
            sample.drill_centers.append((width / 2, height / 2, diameter))
        return sample

    def _normalize_manual_object_key(self, object_key: str) -> str:
        return {"DTP-16": "DTP16"}.get(object_key, object_key)

    def _next_part_number(self, label: str) -> int:
        if not self.layout_result:
            return 1
        count = 0
        prefix = f"{label}-"
        for sample in self.layout_result.samples:
            if sample.sample_id.startswith(prefix):
                count += 1
        return count + 1

    def add_manual_parts(self) -> None:
        try:
            layout = self._ensure_working_layout()
            object_key = self.manual_object_var.get()
            quantity = max(1, int(self.manual_qty_var.get()))
            self.push_undo("Add manual parts")
            added = []
            for i in range(quantity):
                sample = self._make_manual_sample(object_key, i)
                self._place_sample_first_fit(layout, sample)
                layout.samples.append(sample)
                self._add_object_extras(layout, sample, object_key)
                added.append(sample.sample_id)
            self.draw_layout()
            self.show_warnings()
            self.sample_info.delete("1.0", tk.END)
            self.sample_info.insert("1.0", f"Added {len(added)} part(s): {', '.join(added)}.")
        except Exception as exc:
            messagebox.showerror("Add Part Error", str(exc))

    def _add_object_extras(self, layout: LayoutResult, sample: SampleRect, object_key: str) -> None:
        object_key = self._normalize_manual_object_key(object_key)
        if object_key == "STP308":
            score_y = sample.y_in + float(MANUAL_OBJECTS["STP308"]["score_offset"])
            layout.line_entities.append(LineEntity("STP308_SCORE", sample.x_in, score_y, sample.x_in + sample.width_in, score_y))

    def _place_sample_first_fit(self, layout: LayoutResult, sample: SampleRect) -> None:
        gap = 0.25
        x_start = self._preferred_start_x(layout, sample)
        y_start = self._preferred_start_y(layout, sample)
        step = 0.25
        y = y_start
        while y + sample.height_in <= layout.board_height_in + 1e-6:
            x = x_start
            while x + sample.width_in <= layout.board_width_in + 1e-6:
                if self._validate_sample_position(sample, x, y) is None:
                    self._move_sample(sample, x, y)
                    return
                x += step
            y += step
        raise ValueError(f"No open space found for {sample.sample_id}.")

    def _preferred_start_x(self, layout: LayoutResult, sample: SampleRect) -> float:
        if sample.sample_id.startswith("DTP-11"):
            return float(RULES_JSON["DTP-11"]["side_offset_in"])
        if sample_requires_formed_edge(sample):
            return 0.0
        if sample.metadata.get("min_margin"):
            return float(sample.metadata["min_margin"])
        if sample_uses_shop_edge_inset(sample):
            return SHOP_EDGE_INSET_IN
        return 0.0

    def _preferred_start_y(self, layout: LayoutResult, sample: SampleRect) -> float:
        if sample.sample_id.startswith("DTP-11"):
            return float(RULES_JSON["DTP-11"]["end_offset_in"])
        if sample.metadata.get("min_margin"):
            return float(sample.metadata["min_margin"])
        if sample.metadata.get("min_end") and layout.board_height_in >= layout.board_width_in:
            return float(sample.metadata["min_end"])
        if sample_uses_shop_edge_inset(sample):
            return SHOP_EDGE_INSET_IN
        return 0.0

    def run_build_bot(self) -> None:
        if hasattr(self, "bot_input"):
            prompt = self.bot_input.get("1.0", tk.END).strip()
            if self.bot_placeholder_active or prompt == self.bot_placeholder_text:
                prompt = ""
        else:
            prompt = self.bot_prompt_var.get().strip()
        if not prompt:
            messagebox.showerror(
                "Build Bot Error",
                "Type a board request first.\n\nExample: 3 nail pulls, 3 humid bonds, 2 edge shears.",
            )
            return
        try:
            requests = self._parse_builder_prompt(prompt)
            if not requests:
                messagebox.showerror(
                    "Build Bot Error",
                    "I could not find a quantity with a known sample name.\n\n"
                    "Use a number plus the sample name, like:\n"
                    "- 10 nail pulls\n"
                    "- 6 humid bonds\n"
                    "- 4 abrasions\n"
                    "- 3 edge shears\n"
                    "- 5 flexural CD and 5 flexural MD",
                )
                return
            unknown_terms = self._find_unknown_builder_terms(prompt)
            if unknown_terms:
                messagebox.showerror(
                    "Build Bot Error",
                    "I found a quantity with a sample name I do not recognize:\n\n"
                    + "\n".join(f"- {term}" for term in unknown_terms)
                    + "\n\nTry names like nail pulls, humid bonds, abrasions, edge shears, surface indentation, "
                    "humidified deflection, flexural CD, or flexural MD.",
                )
                return
            self._sync_sample_requests_from_bot(requests)
            self._apply_bot_project_options(prompt)
            self.push_undo("Build Bot")
            layout = self._build_from_object_counts(requests)
            self.layout_result = layout
            self._stamp_layout_metadata(self.layout_result)
            self._spread_generated_layout()
            self.selected_sample_id = None
            self.selected_sample_ids.clear()
            self.selection_count_var.set("Selected: 0")
            self.drag_sample_id = None
            self.drag_last_xy = None
            self.last_drag_error = None
            self.draw_layout()
            self.show_warnings()
            self.sample_info.delete("1.0", tk.END)
            summary = ", ".join(f"{self._friendly_object_name(key)} x {qty}" for key, qty in requests)
            self.sample_info.insert("1.0", f"Build Bot generated {len(layout.samples)} part(s): {summary}.")
            self._bot_say(f"Built {len(layout.samples)} part(s): {summary}.")
            if "save" in prompt.lower():
                self._save_current_sheet_silent()
                self._bot_say("Saved the generated sheet.")
            self.bot_prompt_var.set("")
            if hasattr(self, "bot_input"):
                self.bot_input.delete("1.0", tk.END)
                self._set_bot_placeholder()
        except Exception as exc:
            messagebox.showerror("Build Bot Error", self._build_bot_error_message(exc, prompt))

    def _bot_say(self, text: str) -> None:
        if hasattr(self, "bot_status_var"):
            self.bot_status_var.set(text)
        if hasattr(self, "bot_log"):
            self.bot_log.insert(tk.END, f"\n{text}")
            self.bot_log.see(tk.END)

    def _build_bot_error_message(self, exc: Exception, prompt: str) -> str:
        raw = str(exc).strip()
        clean = raw.strip("'\"")
        if clean in {"DTP-16", "DTP16"}:
            return (
                "Surface Indentation was recognized, but the builder could not map it to the cut object.\n\n"
                "Recommendation: use surface indentation or indentations in the Build Bot request. "
                "This has been corrected so DTP-16 maps to the 6 x 6 surface indentation sample."
            )
        if "DTP-11" in raw and "three-sample" in raw:
            return raw + "\n\nRecommendation: request DTP-11 nail pull in groups of 3, such as 3, 6, 9, or 12."
        return (
            f"{raw}\n\n"
            "Recommendation: use a quantity plus a supported sample name, such as:\n"
            "- 3 nail pulls\n"
            "- 3 humid bonds\n"
            "- 2 edge shears\n"
            "- 2 indentations\n"
            "- 2 abrasions\n"
            "- 3 flexurals"
        )

    def _friendly_object_name(self, object_key: str) -> str:
        names = {
            "DTP-11": "Z pull",
            "DTP-13": "pull through",
            "DTP-15": "abrasion",
            "DTP-16": "surface indentation",
            "DTP-17": "soft body impact",
            "STP308": "humid bond",
            "STP312": "flexural",
            "STP312-CD": "flexural CD",
            "STP312-MD": "flexural MD",
            "STP311": "nail pull",
            "STP315": "humidified deflection",
            "STP318": "edge shear",
        }
        return names.get(object_key, object_key)

    def _send_bot_on_enter(self, event):
        if event.state & 0x0001:
            return None
        self.run_build_bot()
        return "break"

    def resize_build_bot(self, delta: int) -> None:
        self.bot_log_height = max(0, min(32, self.bot_log_height + delta))
        self.bot_input_height = max(5, min(18, self.bot_input_height + delta))
        if hasattr(self, "bot_log"):
            self.bot_log.configure(height=self.bot_log_height)
        if hasattr(self, "bot_input"):
            self.bot_input.configure(height=self.bot_input_height)

    def _set_bot_placeholder(self) -> None:
        if not hasattr(self, "bot_input"):
            return
        self.bot_placeholder_active = True
        self.bot_input.delete("1.0", tk.END)
        self.bot_input.insert("1.0", self.bot_placeholder_text)
        self.bot_input.configure(foreground="#6b7280")

    def _clear_bot_placeholder(self, _event=None) -> None:
        if not getattr(self, "bot_placeholder_active", False):
            return
        self.bot_placeholder_active = False
        self.bot_input.delete("1.0", tk.END)
        self.bot_input.configure(foreground="#111827")

    def _restore_bot_placeholder_if_empty(self, _event=None) -> None:
        if not hasattr(self, "bot_input"):
            return
        if self.bot_input.get("1.0", tk.END).strip():
            return
        self._set_bot_placeholder()

    def _parse_builder_prompt(self, prompt: str) -> List[Tuple[str, int]]:
        text = prompt.lower()
        for word, number in BOT_NUMBER_WORDS.items():
            text = re.sub(rf"\b{word}\b", number, text)
        normalized = re.sub(r"[^a-z0-9\-\s]", " ", text)
        results: Dict[str, int] = {}
        consumed_spans: List[Tuple[int, int]] = []

        for alias, object_key in sorted(MANUAL_OBJECT_ALIASES.items(), key=lambda item: -len(item[0])):
            alias_pattern = re.escape(alias)
            patterns = [
                rf"(\d+)\s*(?:x|samples?|parts?)?\s*(?:of)?\s*{alias_pattern}\b",
                rf"{alias_pattern}\s*(?:x|of|qty|quantity)\s*(\d+)\b",
            ]
            for pattern in patterns:
                for match in re.finditer(pattern, normalized):
                    if any(match.start() < end and match.end() > start for start, end in consumed_spans):
                        continue
                    qty = int(match.group(1))
                    results[object_key] = results.get(object_key, 0) + qty
                    consumed_spans.append((match.start(), match.end()))

        expanded: Dict[str, int] = {}
        for key, qty in results.items():
            if key == "STP312":
                cd_qty = qty // 2
                md_qty = qty - cd_qty
                expanded["STP312-CD"] = expanded.get("STP312-CD", 0) + cd_qty
                expanded["STP312-MD"] = expanded.get("STP312-MD", 0) + md_qty
            else:
                expanded[key] = expanded.get(key, 0) + qty

        return list(expanded.items())

    def _find_unknown_builder_terms(self, prompt: str) -> List[str]:
        text = prompt.lower()
        for word, number in BOT_NUMBER_WORDS.items():
            text = re.sub(rf"\b{word}\b", number, text)
        normalized = re.sub(r"[^a-z0-9\-\s,;]", " ", text)
        pieces = re.split(r",|;|\n|\band\b", normalized)
        unknown: List[str] = []
        for piece in pieces:
            piece = re.sub(r"\s+", " ", piece).strip()
            if not piece or not re.search(r"\d+", piece):
                continue
            if "board" in piece or re.search(r"\d+\s*x\s*\d+", piece):
                continue
            if any(re.search(rf"\b{re.escape(alias)}\b", piece) for alias in MANUAL_OBJECT_ALIASES):
                continue
            unknown.append(piece)
        return unknown

    def _sync_sample_requests_from_bot(self, requests: List[Tuple[str, int]]) -> None:
        normalized: Dict[str, int] = {}
        for dtp_id, quantity in requests:
            request_id = "STP312" if dtp_id in {"STP312-CD", "STP312-MD"} else dtp_id
            if request_id not in RULES_JSON:
                continue
            normalized[request_id] = normalized.get(request_id, 0) + quantity

        rows = list(normalized.items())[:SAMPLE_REQUEST_ROWS]
        if not rows:
            return

        def display_value(dtp_id: str) -> str:
            return f"{dtp_id} - {RULES_JSON[dtp_id]['test_name']}"

        first_id, first_qty = rows[0]
        self.dtp_combo.set(display_value(first_id))
        self.quantity_var.set(str(first_qty))

        for dtp_var, qty_var in zip(self.additional_dtp_vars, self.additional_qty_vars):
            dtp_var.set("None")
            qty_var.set("0")

        for (dtp_id, quantity), dtp_var, qty_var in zip(rows[1:], self.additional_dtp_vars, self.additional_qty_vars):
            dtp_var.set(display_value(dtp_id))
            qty_var.set(str(quantity))

        self.setup_preview_dtp_id = first_id
        self._rebuild_params()
        self.update_procedure_steps()

    def _apply_bot_project_options(self, prompt: str) -> None:
        text = prompt.lower()
        board_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:x|by)\s*(\d+(?:\.\d+)?)\s*(?:board|sheet)?", text)
        if board_match:
            self.board_w_var.set(board_match.group(1))
            self.board_h_var.set(board_match.group(2))
        if "vertical" in text:
            self.md_var.set("Vertical")
        elif "horizontal" in text:
            self.md_var.set("Horizontal")
        for product in ["ToughRock", "DensGlass", "DensShield", "DensDeck Prime"]:
            if product.lower() in text:
                self.product_var.set(product)
                break
        label_match = re.search(r"(?:label|text)\s+([a-z0-9 _.-]{2,40})", text)
        if label_match:
            self.insert_text_var.set(label_match.group(1).strip())

    def _build_from_object_counts(self, object_counts: List[Tuple[str, int]]) -> LayoutResult:
        board_width = float(self.board_w_var.get())
        board_height = float(self.board_h_var.get())
        engine_supported = all(key in RULES_JSON for key, _qty in object_counts)

        if engine_supported:
            requests = [
                LayoutRequest(dtp_id=key, quantity=qty, request_index=index)
                for index, (key, qty) in enumerate(object_counts, start=1)
            ]
            if any(request.dtp_id == "DTP-13" for request in requests):
                RULES_JSON["DTP-13"]["drill_pattern"]["diameter_in"] = float(self.hole_size_var.get())
            if len(requests) == 1:
                request = requests[0]
                layout = self.engine.generate_layout(
                    dtp_id=request.dtp_id,
                    board_width_in=board_width,
                    board_height_in=board_height,
                    quantity=request.quantity,
                    machine_direction=self.md_var.get(),
                    orientation=self.orientation_var.get(),
                    code_side=self.code_side_var.get(),
                )
            else:
                layout = self.engine.generate_combined_layout(
                    requests=requests,
                    board_width_in=board_width,
                    board_height_in=board_height,
                    machine_direction=self.md_var.get(),
                    orientation=self.orientation_var.get(),
                    code_side=self.code_side_var.get(),
                )
            layout.metadata["sheet_name"] = self.sheet_name_var.get().strip() or "bot_sheet"
            return layout

        layout = LayoutResult(
            board_width_in=board_width,
            board_height_in=board_height,
            dtp_id="BOT-BUILD",
            samples=[],
            metadata={
                "sheet_name": self.sheet_name_var.get().strip() or "bot_sheet",
                "sheet_layer": "SHEET",
                "label_layer": "LABELS",
                "machine_direction": self.md_var.get(),
                "layers": ",".join(RuleSheetBuilder.REQUIRED_LAYERS),
            },
        )
        RuleSheetBuilder._add_md_arrow(layout, self.md_var.get())
        self.layout_result = layout
        for object_key, quantity in object_counts:
            if object_key == "DTP-11" and quantity % 3 != 0:
                raise ValueError("DTP-11 must be requested in complete three-sample sets: 3, 6, 9, etc.")
            for i in range(quantity):
                sample = self._make_manual_sample(object_key, i)
                self._place_sample_first_fit(layout, sample)
                layout.samples.append(sample)
                self._add_object_extras(layout, sample, object_key)
        layout.warnings.append("Build Bot generated this board from your text request.")
        return layout

    def _build_preview_panel(self, parent: ttk.Frame) -> None:
        preview = ttk.LabelFrame(parent, text="Cut Layout", padding=8)
        preview.pack(fill=tk.BOTH, expand=True)

        export_bar = ttk.Frame(preview)
        export_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(export_bar, text="Export Current DXF", command=self.export_dxf).pack(side=tk.LEFT)
        ttk.Button(export_bar, text="Export Saved DXFs", command=self.export_saved_dxfs).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(export_bar, text="Save Current Sheet", command=self.save_current_sheet).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(export_bar, text="Job Package", command=self.export_job_package).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(export_bar, text="PDF Report", command=self.export_pdf_report).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(export_bar, text="Open Output", command=self.open_last_export_folder).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(export_bar, text="DXF exports use the current/moved part positions.").pack(side=tk.LEFT, padx=(10, 0))

        toolbar = ttk.Frame(preview)
        toolbar.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(toolbar, text="Grid", variable=self.show_grid_var, command=self.draw_layout).pack(side=tk.LEFT)
        ttk.Checkbutton(toolbar, text="Labels", variable=self.show_labels_var, command=self.draw_layout).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Checkbutton(toolbar, text="Highlight", variable=self.highlight_mode_var, command=self.draw_layout).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Checkbutton(toolbar, text="Snap", variable=self.snap_enabled_var).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Combobox(toolbar, textvariable=self.snap_increment_var, values=SNAP_INCREMENTS, width=5).pack(side=tk.LEFT, padx=(3, 0))
        ttk.Button(toolbar, text="Fit View", command=self.fit_view).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Button(toolbar, text="Zoom +", command=lambda: self.zoom_view(1.2)).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(toolbar, text="Zoom -", command=lambda: self.zoom_view(1 / 1.2)).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(toolbar, text="Undo", command=self.undo).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(toolbar, text="Redo", command=self.redo).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Checkbutton(toolbar, text="Measure", variable=self.measure_mode_var, command=self.toggle_measure_mode).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Label(toolbar, text="Text").pack(side=tk.LEFT, padx=(12, 2))
        ttk.Entry(toolbar, textvariable=self.insert_text_var, width=12).pack(side=tk.LEFT)
        ttk.Label(toolbar, text="X").pack(side=tk.LEFT, padx=(4, 1))
        ttk.Entry(toolbar, textvariable=self.insert_text_x_var, width=5).pack(side=tk.LEFT)
        ttk.Label(toolbar, text="Y").pack(side=tk.LEFT, padx=(4, 1))
        ttk.Entry(toolbar, textvariable=self.insert_text_y_var, width=5).pack(side=tk.LEFT)
        ttk.Label(toolbar, text="H").pack(side=tk.LEFT, padx=(4, 1))
        ttk.Entry(toolbar, textvariable=self.insert_text_height_var, width=5).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Insert Text", command=self.insert_text_note).pack(side=tk.LEFT, padx=(6, 0))
        self.cursor_label = ttk.Label(toolbar, text="X: --  Y: --")
        self.cursor_label.pack(side=tk.RIGHT)

        usage = ttk.LabelFrame(preview, text="Board Usage", padding=8)
        usage.pack(fill=tk.X, pady=(0, 6))
        usage_top = ttk.Frame(usage)
        usage_top.pack(fill=tk.X)
        for label_var in [self.usage_board_var, self.usage_parts_var, self.usage_used_var, self.usage_waste_var]:
            ttk.Label(usage_top, textvariable=label_var).pack(side=tk.LEFT, padx=(0, 14))
        self.usage_progress = ttk.Progressbar(
            usage,
            orient=tk.HORIZONTAL,
            mode="determinate",
            maximum=100,
            variable=self.usage_pct_var,
        )
        self.usage_progress.pack(fill=tk.X, pady=(6, 0))

        canvas_frame = ttk.Frame(preview)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(canvas_frame, bg="#d1d5db", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas_v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas_h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas_v_scroll.grid(row=0, column=1, sticky="ns")
        self.canvas_h_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(xscrollcommand=self.canvas_h_scroll.set, yscrollcommand=self.canvas_v_scroll.set)
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<ButtonPress-2>", self.on_canvas_pan_start)
        self.canvas.bind("<B2-Motion>", self.on_canvas_pan_drag)
        self.canvas.bind("<ButtonRelease-2>", self.on_canvas_pan_end)
        self.canvas.bind("<ButtonPress-3>", self.on_canvas_pan_start)
        self.canvas.bind("<B3-Motion>", self.on_canvas_pan_drag)
        self.canvas.bind("<ButtonRelease-3>", self.on_canvas_pan_end)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        self.canvas.bind("<MouseWheel>", self.on_canvas_mousewheel)
        self.canvas.bind("<Button-4>", self.on_canvas_mousewheel)
        self.canvas.bind("<Button-5>", self.on_canvas_mousewheel)

        self.part_tree = ttk.Treeview(
            preview,
            columns=("layer", "x", "y", "w", "h"),
            show="headings",
            height=6,
        )
        for column, label, width in [
            ("layer", "Layer", 90),
            ("x", "X", 60),
            ("y", "Y", 60),
            ("w", "W", 60),
            ("h", "H", 60),
        ]:
            self.part_tree.heading(column, text=label)
            self.part_tree.column(column, width=width, anchor=tk.CENTER)
        self.part_tree.pack(fill=tk.X, pady=(8, 0))
        self.part_tree.bind("<<TreeviewSelect>>", self.on_part_tree_selected)

        edit = ttk.LabelFrame(preview, text="Selected Part", padding=6)
        edit.pack(fill=tk.X, pady=(8, 0))
        for col, (label, var) in enumerate([
            ("X", self.selected_x_var),
            ("Y", self.selected_y_var),
            ("W", self.selected_w_var),
            ("H", self.selected_h_var),
        ]):
            ttk.Label(edit, text=label).grid(row=0, column=col * 2, sticky="w", padx=(0, 2))
            ttk.Entry(edit, textvariable=var, width=7).grid(row=0, column=col * 2 + 1, sticky="w", padx=(0, 6))
        ttk.Button(edit, text="Apply", command=self.apply_selected_part_edits).grid(row=0, column=8, padx=(4, 0))
        ttk.Button(edit, text="Duplicate", command=self.duplicate_selected_sample).grid(row=0, column=9, padx=(4, 0))
        ttk.Button(edit, text="Delete", command=self.delete_selected_sample).grid(row=0, column=10, padx=(4, 0))
        ttk.Button(edit, text="Rotate", command=self.rotate_selected_sample).grid(row=0, column=11, padx=(4, 0))
        ttk.Button(edit, text="Lock", command=self.toggle_selected_lock).grid(row=0, column=12, padx=(4, 0))
        ttk.Button(edit, text="Auto Arrange", command=self.auto_arrange_current_layout).grid(row=0, column=13, padx=(4, 0))
        ttk.Button(edit, text="Fill Board", command=self.fill_board_current_layout).grid(row=0, column=14, padx=(4, 0))
        ttk.Button(edit, text="Check Board", command=self.check_board).grid(row=0, column=15, padx=(4, 0))
        ttk.Label(edit, textvariable=self.selection_count_var).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(edit, text="Edge").grid(row=1, column=3, sticky="e", pady=(6, 0))
        ttk.Combobox(
            edit,
            textvariable=self.edge_side_var,
            values=["Left", "Right", "Top", "Bottom"],
            state="readonly",
            width=8,
        ).grid(row=1, column=4, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(edit, text="Distance").grid(row=1, column=6, sticky="e", pady=(6, 0))
        ttk.Entry(edit, textvariable=self.edge_distance_var, width=7).grid(row=1, column=7, sticky="w", pady=(6, 0))
        ttk.Button(edit, text="Set From Edge", command=self.set_selected_distance_from_edge).grid(
            row=1,
            column=8,
            columnspan=3,
            sticky="w",
            padx=(4, 0),
            pady=(6, 0),
        )
        ttk.Button(edit, text="Clear Selection", command=self.clear_selection).grid(row=1, column=11, columnspan=2, sticky="w", padx=(4, 0), pady=(6, 0))

        self._build_build_bot_panel(preview)

        warning_frame = ttk.Frame(preview)
        warning_frame.pack(fill=tk.X, pady=(10, 0))
        self.warning_box = tk.Text(warning_frame, height=7, wrap="word")
        warning_scroll = ttk.Scrollbar(warning_frame, orient=tk.VERTICAL, command=self.warning_box.yview)
        self.warning_box.configure(yscrollcommand=warning_scroll.set)
        self.warning_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        warning_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._bind_widget_mousewheel(self.warning_box)

    def _build_setup_panel(self, parent: ttk.Frame) -> None:
        setup = ttk.LabelFrame(parent, text="Post-Cut Setup Viewer", padding=10)
        setup.pack(fill=tk.BOTH, expand=True)

        self.step_title = ttk.Label(setup, text="Procedure Steps", font=("Segoe UI", 10, "bold"))
        self.step_title.pack(anchor="w")

        step_frame = ttk.Frame(setup)
        step_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 8))
        self.step_list = tk.Listbox(step_frame, height=10)
        step_scroll = ttk.Scrollbar(step_frame, orient=tk.VERTICAL, command=self.step_list.yview)
        self.step_list.configure(yscrollcommand=step_scroll.set)
        self.step_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        step_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._bind_widget_mousewheel(self.step_list)
        self.step_list.bind("<<ListboxSelect>>", self.on_step_selected)

        self.animation_canvas = tk.Canvas(setup, width=260, height=220, bg="#f8fafc")
        self.animation_canvas.pack(fill=tk.X)

        sample_info_frame = ttk.Frame(setup)
        sample_info_frame.pack(fill=tk.BOTH, expand=False, pady=(8, 0))
        self.sample_info = tk.Text(sample_info_frame, height=7, wrap="word")
        sample_info_scroll = ttk.Scrollbar(sample_info_frame, orient=tk.VERTICAL, command=self.sample_info.yview)
        self.sample_info.configure(yscrollcommand=sample_info_scroll.set)
        self.sample_info.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sample_info_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._bind_widget_mousewheel(self.sample_info)

        sheets = ttk.LabelFrame(setup, text="Saved Sheets", padding=8)
        sheets.pack(fill=tk.BOTH, expand=False, pady=(8, 0))

        self.sheet_notebook = ttk.Notebook(sheets, height=34)
        self.sheet_notebook.pack(fill=tk.X, pady=(0, 6))
        self.sheet_notebook.bind("<<NotebookTabChanged>>", self.on_sheet_tab_changed)

        saved_frame = ttk.Frame(sheets)
        saved_frame.pack(fill=tk.BOTH, expand=True)
        self.saved_sheet_list = tk.Listbox(saved_frame, height=5)
        saved_scroll = ttk.Scrollbar(saved_frame, orient=tk.VERTICAL, command=self.saved_sheet_list.yview)
        self.saved_sheet_list.configure(yscrollcommand=saved_scroll.set)
        self.saved_sheet_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        saved_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._bind_widget_mousewheel(self.saved_sheet_list)
        self.saved_sheet_list.bind("<<ListboxSelect>>", self.on_saved_sheet_selected)

        sheet_buttons = ttk.Frame(sheets)
        sheet_buttons.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(sheet_buttons, text="Load", command=self.load_selected_sheet).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))
        ttk.Button(sheet_buttons, text="Remove", command=self.remove_selected_sheet).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))
        ttk.Button(sheet_buttons, text="New", command=self.new_blank_sheet).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))

        self.update_procedure_steps()

    def _build_build_bot_panel(self, parent: ttk.Frame) -> None:
        bot = ttk.LabelFrame(parent, text="Build Bot", padding=8)
        bot.pack(fill=tk.BOTH, expand=False, pady=(8, 0))
        bot_header = ttk.Frame(bot)
        bot_header.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            bot_header,
            text="Ask for sample names directly. Example: 10 nail pull, 6 humid bond, 4 abrasion, 3 edge shear.",
            wraplength=760,
        ).pack(side=tk.LEFT, anchor="w", fill=tk.X, expand=True)
        ttk.Button(bot_header, text="Taller", command=lambda: self.resize_build_bot(3)).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(bot_header, text="Shorter", command=lambda: self.resize_build_bot(-3)).pack(side=tk.RIGHT)
        ttk.Label(bot, textvariable=self.bot_status_var, wraplength=1000).pack(fill=tk.X, pady=(0, 6))
        bot_entry = ttk.LabelFrame(bot, text="Type here", padding=6)
        bot_entry.pack(fill=tk.BOTH, expand=False)
        self.bot_input = tk.Text(bot_entry, height=self.bot_input_height, wrap="word")
        bot_input_scroll = ttk.Scrollbar(bot_entry, orient=tk.VERTICAL, command=self.bot_input.yview)
        self.bot_input.configure(yscrollcommand=bot_input_scroll.set)
        self.bot_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        bot_input_scroll.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))
        self._bind_widget_mousewheel(self.bot_input)
        self.bot_input.bind("<Return>", self._send_bot_on_enter)
        self.bot_input.bind("<FocusIn>", self._clear_bot_placeholder)
        self.bot_input.bind("<FocusOut>", self._restore_bot_placeholder_if_empty)
        self._set_bot_placeholder()
        ttk.Button(bot_entry, text="Enter / Build Board", command=self.run_build_bot).pack(side=tk.RIGHT, fill=tk.Y)

    def update_procedure_steps(self) -> None:
        self.step_list.delete(0, tk.END)
        for dtp_id in self._active_dtp_ids():
            self.step_list.insert(tk.END, f"{dtp_id}: {RULES_JSON[dtp_id]['test_name']}")
            for i, step in enumerate(PROCEDURE_STEPS.get(dtp_id, []), start=1):
                self.step_list.insert(tk.END, f"  {i}. {step}")
        self.draw_animation_placeholder()

    def on_step_selected(self, _event=None) -> None:
        selected = self.step_list.curselection()
        if not selected:
            return
        row_text = self.step_list.get(selected[0])
        match = re.match(r"^(DTP-\d+|STP\d+):", row_text)
        if match:
            self.setup_preview_dtp_id = match.group(1)
            self.draw_animation_placeholder()

    def generate_layout(self) -> None:
        try:
            board_width = float(self.board_w_var.get())
            board_height = float(self.board_h_var.get())
            machine_direction = self.md_var.get()
            orientation = self.orientation_var.get()
            code_side = self.code_side_var.get()
            requests = self._collect_layout_requests()

            if any(request.dtp_id == "DTP-13" for request in requests):
                hole_size = float(self.hole_size_var.get())
                RULES_JSON["DTP-13"]["drill_pattern"]["diameter_in"] = hole_size

            self.push_undo("Generate layout")
            if len(requests) == 1:
                request = requests[0]
                self.layout_result = self.engine.generate_layout(
                    dtp_id=request.dtp_id,
                    board_width_in=board_width,
                    board_height_in=board_height,
                    quantity=request.quantity,
                    machine_direction=machine_direction,
                    orientation=orientation,
                    code_side=code_side,
                )
            else:
                self.layout_result = self.engine.generate_combined_layout(
                    requests=requests,
                    board_width_in=board_width,
                    board_height_in=board_height,
                    machine_direction=machine_direction,
                    orientation=orientation,
                    code_side=code_side,
                )
            self._stamp_layout_metadata(self.layout_result)
            self._spread_generated_layout()
            self.selected_sample_id = None
            self.selected_sample_ids.clear()
            self.selection_count_var.set("Selected: 0")
            self.drag_sample_id = None
            self.drag_last_xy = None
            self.last_drag_error = None
            self.draw_layout()
            self.show_warnings()
            self.sample_info.delete("1.0", tk.END)
            summary = ", ".join(f"{request.dtp_id} x {request.quantity}" for request in requests)
            self.sample_info.insert(
                "1.0",
                f"Generated {len(self.layout_result.samples)} sample(s): {summary}.\n"
                "Click and drag a sample in the preview to move it."
            )
        except Exception as exc:
            messagebox.showerror("Layout Error", str(exc))

    def _stamp_layout_metadata(self, layout: LayoutResult) -> None:
        layout.metadata.setdefault("sheet_name", self.sheet_name_var.get().strip() or f"Sheet {len(self.saved_sheets) + 1}")
        layout.metadata["project_id"] = self.project_var.get().strip()
        layout.metadata["product"] = self.product_var.get().strip()
        layout.metadata["operator"] = self.operator_var.get().strip()
        layout.metadata["machine_direction"] = self.md_var.get()
        layout.metadata["user_email"] = self.current_user_email

    def load_user_profile(self) -> None:
        self.saved_sheets = []
        if not self.profile_path or not os.path.exists(self.profile_path):
            self.refresh_saved_sheet_list()
            return
        try:
            with open(self.profile_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            sheets = data.get("saved_sheets", [])
            self.saved_sheets = [self.layout_from_dict(sheet_data) for sheet_data in sheets]
            self.refresh_saved_sheet_list()
        except Exception as exc:
            self.saved_sheets = []
            self.refresh_saved_sheet_list()
            messagebox.showwarning(
                "Profile Load Warning",
                f"Could not load saved sheets for {self.current_user_email}.\n\n{exc}",
            )

    def save_user_profile(self) -> None:
        if not self.profile_path:
            return
        os.makedirs(self.profile_dir, exist_ok=True)
        data = {
            "email": self.current_user_email,
            "saved_sheets": [self.layout_to_dict(sheet) for sheet in self.saved_sheets],
        }
        with open(self.profile_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def _layout_snapshot(self) -> Optional[Dict]:
        if not self.layout_result:
            return None
        return {
            "layout": self.layout_to_dict(self.layout_result),
            "selected_sample_id": self.selected_sample_id,
            "selected_sample_ids": sorted(self.selected_sample_ids),
            "sheet_name": self.sheet_name_var.get(),
        }

    def _restore_layout_snapshot(self, snapshot: Dict) -> None:
        self.layout_result = self.layout_from_dict(snapshot["layout"])
        self.selected_sample_id = snapshot.get("selected_sample_id")
        self.selected_sample_ids = set(snapshot.get("selected_sample_ids", []))
        self.sheet_name_var.set(snapshot.get("sheet_name") or self.layout_result.metadata.get("sheet_name", "Sheet 1"))
        self.drag_sample_id = None
        self.drag_last_xy = None
        self.last_drag_error = None
        self.draw_layout()
        self.show_warnings()
        if self.selected_sample_id:
            self.select_sample(self.selected_sample_id)
        elif self.selected_sample_ids:
            self.select_samples(self.selected_sample_ids)

    def push_undo(self, _label: str = "") -> None:
        snapshot = self._layout_snapshot()
        if not snapshot:
            return
        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self) -> None:
        if not self.undo_stack:
            return
        current = self._layout_snapshot()
        snapshot = self.undo_stack.pop()
        if current:
            self.redo_stack.append(current)
        self._restore_layout_snapshot(snapshot)

    def redo(self) -> None:
        if not self.redo_stack:
            return
        current = self._layout_snapshot()
        snapshot = self.redo_stack.pop()
        if current:
            self.undo_stack.append(current)
        self._restore_layout_snapshot(snapshot)

    @staticmethod
    def layout_to_dict(layout: LayoutResult) -> Dict:
        return {
            "board_width_in": layout.board_width_in,
            "board_height_in": layout.board_height_in,
            "dtp_id": layout.dtp_id,
            "samples": [
                {
                    "sample_id": sample.sample_id,
                    "x_in": sample.x_in,
                    "y_in": sample.y_in,
                    "width_in": sample.width_in,
                    "height_in": sample.height_in,
                    "rotation_deg": sample.rotation_deg,
                    "drill_centers": [list(center) for center in sample.drill_centers],
                    "metadata": dict(sample.metadata),
                }
                for sample in layout.samples
            ],
            "scrap_zones": [list(zone) for zone in layout.scrap_zones],
            "warnings": list(layout.warnings),
            "line_entities": [
                {
                    "layer": entity.layer,
                    "x1_in": entity.x1_in,
                    "y1_in": entity.y1_in,
                    "x2_in": entity.x2_in,
                    "y2_in": entity.y2_in,
                }
                for entity in layout.line_entities
            ],
            "text_entities": [
                {
                    "layer": entity.layer,
                    "x_in": entity.x_in,
                    "y_in": entity.y_in,
                    "height_in": entity.height_in,
                    "text": entity.text,
                }
                for entity in layout.text_entities
            ],
            "metadata": dict(layout.metadata),
        }

    @staticmethod
    def layout_from_dict(data: Dict) -> LayoutResult:
        return LayoutResult(
            board_width_in=float(data["board_width_in"]),
            board_height_in=float(data["board_height_in"]),
            dtp_id=str(data.get("dtp_id", "COMBINED")),
            samples=[
                SampleRect(
                    sample_id=str(sample.get("sample_id", "")),
                    x_in=float(sample.get("x_in", 0)),
                    y_in=float(sample.get("y_in", 0)),
                    width_in=float(sample.get("width_in", 0)),
                    height_in=float(sample.get("height_in", 0)),
                    rotation_deg=int(sample.get("rotation_deg", 0)),
                    drill_centers=[
                        (float(center[0]), float(center[1]), float(center[2]))
                        for center in sample.get("drill_centers", [])
                    ],
                    metadata=dict(sample.get("metadata", {})),
                )
                for sample in data.get("samples", [])
            ],
            scrap_zones=[
                (float(zone[0]), float(zone[1]), float(zone[2]), float(zone[3]))
                for zone in data.get("scrap_zones", [])
            ],
            warnings=[str(warning) for warning in data.get("warnings", [])],
            line_entities=[
                LineEntity(
                    layer=str(entity.get("layer", "SHEET")),
                    x1_in=float(entity.get("x1_in", 0)),
                    y1_in=float(entity.get("y1_in", 0)),
                    x2_in=float(entity.get("x2_in", 0)),
                    y2_in=float(entity.get("y2_in", 0)),
                )
                for entity in data.get("line_entities", [])
            ],
            text_entities=[
                TextEntity(
                    layer=str(entity.get("layer", "LABELS")),
                    x_in=float(entity.get("x_in", 0)),
                    y_in=float(entity.get("y_in", 0)),
                    height_in=float(entity.get("height_in", 0.35)),
                    text=str(entity.get("text", "")),
                )
                for entity in data.get("text_entities", [])
            ],
            metadata=dict(data.get("metadata", {})),
        )

    def save_current_sheet(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Nothing to save", "Generate a layout first.")
            return

        sheet = self._save_current_sheet_silent()
        messagebox.showinfo("Sheet Saved", f"Saved {sheet.metadata['sheet_name']}.")

    def _save_current_sheet_silent(self) -> LayoutResult:
        if not self.layout_result:
            raise ValueError("Generate a layout first.")
        sheet = copy.deepcopy(self.layout_result)
        self._stamp_layout_metadata(sheet)
        sheet.metadata["sheet_name"] = self.sheet_name_var.get().strip() or f"Sheet {len(self.saved_sheets) + 1}"
        self.saved_sheets.append(sheet)
        self.refresh_saved_sheet_list()
        self.save_user_profile()
        return sheet

    def refresh_saved_sheet_list(self) -> None:
        self.saved_sheet_list.delete(0, tk.END)
        for index, sheet in enumerate(self.saved_sheets, start=1):
            name = sheet.metadata.get("sheet_name", f"Sheet {index}")
            self.saved_sheet_list.insert(tk.END, f"{index}. {name} ({len(sheet.samples)} parts)")
        if hasattr(self, "sheet_notebook"):
            self.updating_sheet_tabs = True
            try:
                for tab_id in self.sheet_notebook.tabs():
                    self.sheet_notebook.forget(tab_id)
                for index, sheet in enumerate(self.saved_sheets, start=1):
                    frame = ttk.Frame(self.sheet_notebook)
                    name = sheet.metadata.get("sheet_name", f"Sheet {index}")
                    self.sheet_notebook.add(frame, text=name[:18] or f"Sheet {index}")
            finally:
                self.updating_sheet_tabs = False

    def selected_saved_sheet_index(self) -> Optional[int]:
        selected = self.saved_sheet_list.curselection()
        if not selected:
            return None
        return int(selected[0])

    def on_sheet_tab_changed(self, _event=None) -> None:
        if self.updating_sheet_tabs or not hasattr(self, "sheet_notebook"):
            return
        selected = self.sheet_notebook.select()
        if not selected:
            return
        tabs = list(self.sheet_notebook.tabs())
        if selected not in tabs:
            return
        index = tabs.index(selected)
        if index >= len(self.saved_sheets):
            return
        self.saved_sheet_list.selection_clear(0, tk.END)
        self.saved_sheet_list.selection_set(index)
        self.load_selected_sheet()

    def new_blank_sheet(self) -> None:
        try:
            self.push_undo("New blank sheet")
            board_width = float(self.board_w_var.get())
            board_height = float(self.board_h_var.get())
            self.layout_result = LayoutResult(
                board_width_in=board_width,
                board_height_in=board_height,
                dtp_id="MANUAL",
                samples=[],
                metadata={
                    "sheet_name": self.sheet_name_var.get().strip() or f"Sheet {len(self.saved_sheets) + 1}",
                    "sheet_layer": "SHEET",
                    "label_layer": "LABELS",
                    "machine_direction": self.md_var.get(),
                },
            )
            RuleSheetBuilder._add_md_arrow(self.layout_result, self.md_var.get())
            self._stamp_layout_metadata(self.layout_result)
            self.selected_sample_id = None
            self.selected_sample_ids.clear()
            self.selection_count_var.set("Selected: 0")
            self.draw_layout()
            self.show_warnings()
        except Exception as exc:
            messagebox.showerror("New Sheet Error", str(exc))

    def on_saved_sheet_selected(self, _event=None) -> None:
        self.load_selected_sheet()

    def load_selected_sheet(self) -> None:
        index = self.selected_saved_sheet_index()
        if index is None:
            return
        self.layout_result = copy.deepcopy(self.saved_sheets[index])
        self._sync_project_fields_from_layout(self.layout_result)
        if not self.sheet_name_var.get().strip():
            self.sheet_name_var.set(f"Sheet {index + 1}")
        self.selected_sample_id = None
        self.selected_sample_ids.clear()
        self.selection_count_var.set("Selected: 0")
        self.drag_sample_id = None
        self.drag_last_xy = None
        self.last_drag_error = None
        self.draw_layout()
        self.show_warnings()
        self.sample_info.delete("1.0", tk.END)
        self.sample_info.insert(
            "1.0",
            f"Loaded {self.layout_result.metadata.get('sheet_name', 'saved sheet')}.\n"
            "Click and drag a sample in the preview to move it."
        )

    def remove_selected_sheet(self) -> None:
        index = self.selected_saved_sheet_index()
        if index is None:
            return
        del self.saved_sheets[index]
        self.refresh_saved_sheet_list()
        self.save_user_profile()

    def export_saved_dxfs(self) -> None:
        if not self.saved_sheets:
            messagebox.showinfo("Nothing to export", "Save at least one sheet first.")
            return
        folder = filedialog.askdirectory(title="Choose folder for DXF files")
        if not folder:
            return
        try:
            for index, sheet in enumerate(self.saved_sheets, start=1):
                name = sheet.metadata.get("sheet_name", f"sheet{index}")
                path = os.path.join(folder, f"{self._safe_file_stem(name) or f'sheet{index}'}.dxf")
                DxfExporter.export(path, sheet)
            self.last_export_folder = folder
            messagebox.showinfo("Export Complete", f"Exported {len(self.saved_sheets)} DXF file(s).")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    @staticmethod
    def _safe_file_stem(name: str) -> str:
        allowed = []
        for char in name.lower().strip():
            if char.isalnum():
                allowed.append(char)
            elif char in {" ", "-", "_"}:
                allowed.append("_")
        stem = "".join(allowed).strip("_")
        while "__" in stem:
            stem = stem.replace("__", "_")
        return stem

    def build_rule_sheet_set(self) -> None:
        try:
            board_width = float(self.board_w_var.get())
            board_height = float(self.board_h_var.get())
            machine_direction = self.md_var.get()
            sheets = RuleSheetBuilder.build_all(board_width, board_height, machine_direction)
            for sheet in sheets:
                self._stamp_layout_metadata(sheet)
                sheet.metadata["project_id"] = self.project_var.get().strip()
                sheet.metadata["product"] = self.product_var.get().strip()
                sheet.metadata["operator"] = self.operator_var.get().strip()
            self.saved_sheets.extend(sheets)
            self.refresh_saved_sheet_list()
            self.save_user_profile()
            self.layout_result = copy.deepcopy(sheets[0])
            self.sheet_name_var.set(self.layout_result.metadata.get("sheet_name", "board1"))
            self.draw_layout()
            self.show_warnings()
            self.sample_info.delete("1.0", tk.END)
            self.sample_info.insert(
                "1.0",
                "Built and saved board1, board2, and board3 from the rule sheet set.\n"
                "Use Export Saved DXFs when ready."
            )
        except Exception as exc:
            messagebox.showerror("Rule Sheet Error", str(exc))

    def draw_layout(self) -> None:
        self.canvas.delete("all")
        if not self.layout_result:
            self.refresh_part_tree()
            return

        w = max(self.canvas.winfo_width(), 400)
        h = max(self.canvas.winfo_height(), 300)
        margin = 40
        scale_x = (w - 2 * margin) / self.layout_result.board_width_in
        scale_y = (h - 2 * margin) / self.layout_result.board_height_in
        scale = min(scale_x, scale_y) * self.view_zoom
        md_pad = 5 * scale
        drawing_w = self.layout_result.board_width_in * scale + 2 * margin + md_pad
        drawing_h = self.layout_result.board_height_in * scale + 2 * margin + md_pad
        scroll_w = max(w, drawing_w)
        scroll_h = max(h, drawing_h)

        bx0 = margin
        by0 = margin
        bx1 = bx0 + self.layout_result.board_width_in * scale
        by1 = by0 + self.layout_result.board_height_in * scale
        self.canvas_scale = scale
        self.canvas_origin = (bx0, by0)
        self.canvas.configure(scrollregion=(0, 0, scroll_w, scroll_h))
        self.canvas.create_rectangle(0, 0, scroll_w, scroll_h, fill="#d1d5db", outline="")
        self.canvas.create_rectangle(bx0, by0, bx1, by1, outline="#111827", width=2, fill="#ffffff")
        self.canvas.create_text(bx0, by0 - 16, text="0,0", anchor="w", fill="#374151", font=("Segoe UI", 8))
        self.canvas.create_text(bx1, by0 - 16, text=f"{self.layout_result.board_width_in:g} in", anchor="e", fill="#374151", font=("Segoe UI", 8))
        self.canvas.create_text(bx0 - 8, by1, text=f"{self.layout_result.board_height_in:g} in", anchor="e", fill="#374151", font=("Segoe UI", 8))
        self._draw_rule_zones(bx0, by0, bx1, by1, scale)

        if self.show_grid_var.get():
            self._draw_grid(bx0, by0, bx1, by1, scale)

        for x, y, sw, sh in self.layout_result.scrap_zones:
            self.canvas.create_rectangle(
                bx0 + x * scale,
                by0 + y * scale,
                bx0 + (x + sw) * scale,
                by0 + (y + sh) * scale,
                fill="#fecaca",
                outline="#ef4444",
                stipple="gray25",
            )

        layer_colors = {
            "FORMED_EDGE": "#7c3aed",
            "MD_ARROW": "#16a34a",
            "STP308_SCORE": "#dc2626",
        }
        for line in self.layout_result.line_entities:
            self.canvas.create_line(
                bx0 + line.x1_in * scale,
                by0 + line.y1_in * scale,
                bx0 + line.x2_in * scale,
                by0 + line.y2_in * scale,
                fill=layer_colors.get(line.layer, "#111827"),
                width=2,
            )

        # MD arrow
        md = self.md_var.get()
        if not self.layout_result.line_entities and md == "Horizontal":
            self.canvas.create_line(bx0 + 20, by1 + 15, bx0 + 120, by1 + 15, arrow=tk.LAST, width=3)
            self.canvas.create_text(bx0 + 70, by1 + 30, text="MD", font=("Segoe UI", 9, "bold"))
        elif not self.layout_result.line_entities:
            self.canvas.create_line(bx1 + 15, by1 - 20, bx1 + 15, by1 - 120, arrow=tk.LAST, width=3)
            self.canvas.create_text(bx1 + 30, by1 - 70, text="MD", font=("Segoe UI", 9, "bold"))

        invalid_ids = {sample_id for sample_id, _reason in self._board_validation_issues(self.layout_result)}
        for sample in self.layout_result.samples:
            x0 = bx0 + sample.x_in * scale
            y0 = by0 + sample.y_in * scale
            x1 = bx0 + (sample.x_in + sample.width_in) * scale
            y1 = by0 + (sample.y_in + sample.height_in) * scale
            tag = sample.sample_id
            is_selected = sample.sample_id in self.selected_sample_ids and self.highlight_mode_var.get()
            is_invalid = sample.sample_id in invalid_ids
            fill = "#fecaca" if is_invalid else ("#fde68a" if is_selected else "#bfdbfe")
            outline = "#dc2626" if is_invalid else ("#b45309" if is_selected else "#1d4ed8")
            line_width = 4 if is_selected else 2
            self.canvas.create_rectangle(
                x0,
                y0,
                x1,
                y1,
                fill=fill,
                outline=outline,
                width=line_width,
                tags=(tag, "sample", "sample_rect"),
            )
            if self.show_labels_var.get():
                self.canvas.create_text(
                    (x0 + x1) / 2,
                    (y0 + y1) / 2,
                    text=sample.sample_id,
                    width=max(60, (x1 - x0) - 8),
                    tags=(tag, "sample", "sample_label"),
                )
            for cx, cy, dia in sample.drill_centers:
                r = dia * scale / 2
                self.canvas.create_oval(
                    bx0 + cx * scale - r,
                    by0 + cy * scale - r,
                    bx0 + cx * scale + r,
                    by0 + cy * scale + r,
                    outline="#111827",
                    width=2,
                    tags=(tag, "sample", "sample_drill"),
                )

        for text in self.layout_result.text_entities:
            self.canvas.create_text(
                bx0 + text.x_in * scale,
                by0 + text.y_in * scale,
                text=text.text,
                fill="#111827",
                font=("Segoe UI", 8),
                anchor="w",
            )
        if self.measure_start:
            sx, sy = self.measure_start
            self.canvas.create_oval(
                bx0 + sx * scale - 4,
                by0 + sy * scale - 4,
                bx0 + sx * scale + 4,
                by0 + sy * scale + 4,
                fill="#dc2626",
                outline="#991b1b",
                tags=("measure",),
            )
        self.refresh_part_tree()

    def _draw_grid(self, bx0: float, by0: float, bx1: float, by1: float, scale: float) -> None:
        if not self.layout_result:
            return
        minor = 6.0
        x = minor
        while x < self.layout_result.board_width_in:
            canvas_x = bx0 + x * scale
            self.canvas.create_line(canvas_x, by0, canvas_x, by1, fill="#e5e7eb")
            if abs(x % 12) < 1e-6:
                self.canvas.create_text(canvas_x + 2, by0 + 10, text=f"{x:g}", anchor="w", fill="#9ca3af", font=("Segoe UI", 7))
            x += minor
        y = minor
        while y < self.layout_result.board_height_in:
            canvas_y = by0 + y * scale
            self.canvas.create_line(bx0, canvas_y, bx1, canvas_y, fill="#e5e7eb")
            if abs(y % 12) < 1e-6:
                self.canvas.create_text(bx0 + 2, canvas_y - 2, text=f"{y:g}", anchor="sw", fill="#9ca3af", font=("Segoe UI", 7))
            y += minor

    def _draw_rule_zones(self, bx0: float, by0: float, bx1: float, by1: float, scale: float) -> None:
        if not self.layout_result:
            return
        margins = []
        for sample in self.layout_result.samples:
            if sample.metadata.get("min_margin"):
                margins.append(float(sample.metadata["min_margin"]))
            elif sample_uses_shop_edge_inset(sample):
                margins.append(SHOP_EDGE_INSET_IN)
            if sample.sample_id.startswith("DTP-11"):
                margins.append(float(RULES_JSON["DTP-11"]["side_offset_in"]))
            if sample.sample_id.startswith("STP315"):
                margins.append(float(RULES_JSON["STP315"]["min_end_in"]))
        if not margins:
            return
        margin = min(max(margins), min(self.layout_result.board_width_in, self.layout_result.board_height_in) / 2)
        x0 = bx0 + margin * scale
        y0 = by0 + margin * scale
        x1 = bx1 - margin * scale
        y1 = by1 - margin * scale
        if x1 > x0 and y1 > y0:
            self.canvas.create_rectangle(x0, y0, x1, y1, outline="#f59e0b", dash=(6, 4), width=2)

    def refresh_part_tree(self) -> None:
        if not hasattr(self, "part_tree"):
            return
        self.updating_part_tree = True
        selected_ids = set(self.selected_sample_ids)
        try:
            for item in self.part_tree.get_children():
                self.part_tree.delete(item)
            if not self.layout_result:
                return
            for sample in self.layout_result.samples:
                self.part_tree.insert(
                    "",
                    tk.END,
                    iid=sample.sample_id,
                    values=(
                        sample.metadata.get("layer", "CUT"),
                        f"{sample.x_in:.2f}",
                        f"{sample.y_in:.2f}",
                        f"{sample.width_in:.2f}",
                        f"{sample.height_in:.2f}",
                    ),
                )
            existing_ids = [sample_id for sample_id in selected_ids if sample_id in self.part_tree.get_children()]
            if existing_ids:
                self.part_tree.selection_set(existing_ids)
                self.part_tree.see(existing_ids[-1])
        finally:
            self.updating_part_tree = False

    def show_warnings(self) -> None:
        self.warning_box.delete("1.0", tk.END)
        if not self.layout_result:
            self._update_board_usage_display(None)
            return
        text = []
        for item in self.layout_result.warnings:
            text.append(f"- {item}")
        for sample_id, issue in self._board_validation_issues(self.layout_result):
            text.append(f"- {sample_id}: {issue}")
        used_area, board_area, waste_area, used_pct = self._board_usage_stats(self.layout_result)
        if board_area > 0:
            text.append(f"- Used area: {used_area:.2f} sq in ({used_pct:.1f}% of board).")
            text.append(f"- Estimated waste/open area: {waste_area:.2f} sq in.")
        if not text:
            text.append("- No warnings.")
        self.warning_box.insert("1.0", "\n".join(text))
        self._update_board_usage_display(self.layout_result)

    def _board_usage_stats(self, layout: LayoutResult) -> Tuple[float, float, float, float]:
        used_area = sum(sample.width_in * sample.height_in for sample in layout.samples)
        board_area = layout.board_width_in * layout.board_height_in
        waste_area = max(0.0, board_area - used_area)
        used_pct = used_area / board_area * 100 if board_area > 0 else 0.0
        return used_area, board_area, waste_area, used_pct

    def _update_board_usage_display(self, layout: Optional[LayoutResult]) -> None:
        if not hasattr(self, "usage_pct_var"):
            return
        if not layout:
            self.usage_pct_var.set(0.0)
            self.usage_board_var.set("Board: --")
            self.usage_parts_var.set("Parts: --")
            self.usage_used_var.set("Used: --")
            self.usage_waste_var.set("Waste: --")
            return
        used_area, board_area, waste_area, used_pct = self._board_usage_stats(layout)
        self.usage_pct_var.set(min(100.0, max(0.0, used_pct)))
        self.usage_board_var.set(f"Board: {layout.board_width_in:g} x {layout.board_height_in:g} in")
        self.usage_parts_var.set(f"Parts: {len(layout.samples)}")
        self.usage_used_var.set(f"Used: {used_area:.1f} sq in ({used_pct:.1f}%)")
        self.usage_waste_var.set(f"Waste/Open: {waste_area:.1f} sq in")

    def _board_validation_issues(self, layout: LayoutResult) -> List[Tuple[str, str]]:
        issues = []
        seen = set()
        for sample in layout.samples:
            reason = self._validate_sample_position(sample, sample.x_in, sample.y_in, allow_locked=True)
            if reason:
                clean_reason = reason.replace("Move blocked: ", "")
                issues.append((sample.sample_id, clean_reason))
                seen.add(sample.sample_id)
        for i, sample in enumerate(layout.samples):
            for other in layout.samples[i + 1:]:
                if self._rects_overlap(
                    sample.x_in,
                    sample.y_in,
                    sample.width_in,
                    sample.height_in,
                    other.x_in,
                    other.y_in,
                    other.width_in,
                    other.height_in,
                ):
                    if sample.sample_id not in seen:
                        issues.append((sample.sample_id, f"Overlaps {other.sample_id}."))
                        seen.add(sample.sample_id)
                    if other.sample_id not in seen:
                        issues.append((other.sample_id, f"Overlaps {sample.sample_id}."))
                        seen.add(other.sample_id)
        return issues

    def check_board(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Check Board", "Generate or load a sheet first.")
            return
        self.show_warnings()
        issues = self._board_validation_issues(self.layout_result)
        if not issues:
            messagebox.showinfo("Check Board", "Board check passed. No placement issues found.")
            return
        detail = "\n".join(f"{sample_id}: {issue}" for sample_id, issue in issues[:12])
        if len(issues) > 12:
            detail += f"\n...and {len(issues) - 12} more issue(s)."
        messagebox.showwarning("Check Board", detail)

    def auto_arrange_current_layout(self) -> None:
        if not self.layout_result:
            return
        unlocked = [sample for sample in self.layout_result.samples if sample.metadata.get("locked") != "true"]
        if not unlocked:
            messagebox.showinfo("Auto Arrange", "All samples are locked.")
            return
        self.push_undo("Auto arrange")
        locked = [sample for sample in self.layout_result.samples if sample.metadata.get("locked") == "true"]
        self.layout_result.samples = locked[:]
        for sample in sorted(unlocked, key=lambda s: s.width_in * s.height_in, reverse=True):
            sample.drill_centers = [(cx - sample.x_in, cy - sample.y_in, dia) for cx, cy, dia in sample.drill_centers]
            sample.x_in = 0
            sample.y_in = 0
            self._place_sample_first_fit(self.layout_result, sample)
            self.layout_result.samples.append(sample)
        self._rebuild_score_lines()
        self.draw_layout()
        self.show_warnings()

    def fill_board_current_layout(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Fill Board", "Generate or load a sheet first.")
            return
        unlocked = [sample for sample in self.layout_result.samples if sample.metadata.get("locked") != "true"]
        if not unlocked:
            messagebox.showinfo("Fill Board", "All samples are locked.")
            return

        before = copy.deepcopy(self.layout_result)
        self.push_undo("Fill board")
        try:
            moved = self._spread_layout_samples(self.layout_result, preserve_locked=True)
        except Exception as exc:
            self.layout_result = before
            self.draw_layout()
            self.show_warnings()
            messagebox.showerror("Fill Board", str(exc))
            return

        if moved:
            note = "Fill Board spread movable samples across the usable board area while keeping rule limits."
            if note not in self.layout_result.warnings:
                self.layout_result.warnings.append(note)
        self.draw_layout()
        self.show_warnings()
        if not moved:
            messagebox.showinfo("Fill Board", "No movable samples needed repositioning.")

    def _spread_generated_layout(self) -> None:
        if not self.layout_result or len(self.layout_result.samples) < 2:
            return
        if any(sample.metadata.get("full_board_fixture") == "true" for sample in self.layout_result.samples):
            return

        before = copy.deepcopy(self.layout_result)
        try:
            moved = self._spread_layout_samples(self.layout_result, preserve_locked=False)
        except Exception:
            self.layout_result = before
            return
        if moved:
            note = "Layout was automatically spread across the board with rule-safe placement."
            if note not in self.layout_result.warnings:
                self.layout_result.warnings.append(note)

    def _spread_layout_samples(self, layout: LayoutResult, preserve_locked: bool = True) -> bool:
        movable = [
            sample
            for sample in layout.samples
            if sample.metadata.get("full_board_fixture") != "true"
            and (not preserve_locked or sample.metadata.get("locked") != "true")
        ]
        if len(movable) < 2:
            return False

        original_positions = {sample.sample_id: (sample.x_in, sample.y_in) for sample in layout.samples}
        movable_ids = {id(sample) for sample in movable}
        fixed = [sample for sample in layout.samples if id(sample) not in movable_ids]
        ordered = sorted(
            movable,
            key=lambda sample: (
                -self._spread_constraint_weight(sample),
                -(sample.width_in * sample.height_in),
                sample.sample_id,
            ),
        )
        group_indexes = self._spread_group_indexes(ordered)

        layout.samples = fixed[:]
        for sample in ordered:
            sample.drill_centers = [(cx - sample.x_in, cy - sample.y_in, dia) for cx, cy, dia in sample.drill_centers]
            sample.x_in = 0.0
            sample.y_in = 0.0
            target = self._spread_target_for_sample(layout, sample, *group_indexes[sample.sample_id])
            candidate = self._best_spread_candidate(layout, sample, target, require_spacing=True)
            if candidate is None:
                candidate = self._best_spread_candidate(layout, sample, target, require_spacing=False)
            if candidate is None:
                self._place_sample_first_fit(layout, sample)
            else:
                self._move_sample(sample, candidate[0], candidate[1])
            layout.samples.append(sample)

        self._rebuild_score_lines()
        issues = self._board_validation_issues(layout)
        if issues:
            detail = "; ".join(f"{sample_id}: {issue}" for sample_id, issue in issues[:3])
            raise ValueError(f"Fill Board could not keep all samples inside their procedure rules. {detail}")

        return any(
            abs(sample.x_in - original_positions.get(sample.sample_id, (sample.x_in, sample.y_in))[0]) > 1e-6
            or abs(sample.y_in - original_positions.get(sample.sample_id, (sample.x_in, sample.y_in))[1]) > 1e-6
            for sample in layout.samples
        )

    def _spread_group_indexes(self, samples: List[SampleRect]) -> Dict[str, Tuple[int, int]]:
        groups: Dict[str, List[SampleRect]] = {}
        for sample in samples:
            key = self._spread_group_key(sample)
            groups.setdefault(key, []).append(sample)

        indexes: Dict[str, Tuple[int, int]] = {}
        for group_samples in groups.values():
            total = len(group_samples)
            for index, sample in enumerate(group_samples):
                indexes[sample.sample_id] = (index, total)
        return indexes

    def _spread_group_key(self, sample: SampleRect) -> str:
        return "|".join(
            [
                sample.metadata.get("dtp_id", sample.sample_id.split("-S", 1)[0]),
                sample.metadata.get("orientation", ""),
                sample.metadata.get("formed_edge", ""),
                sample.metadata.get("edge_taken", ""),
                sample.metadata.get("centered_width", ""),
            ]
        )

    def _spread_constraint_weight(self, sample: SampleRect) -> int:
        weight = 0
        if sample.metadata.get("centered_width") == "true":
            weight += 40
        if sample_requires_formed_edge(sample):
            weight += 30
        if sample.metadata.get("min_end"):
            weight += 20
        if sample.metadata.get("min_margin"):
            weight += 10
        if sample.sample_id.startswith("DTP-11"):
            weight += 25
        return weight

    def _spread_target_for_sample(self, layout: LayoutResult, sample: SampleRect, index: int, total: int) -> Tuple[float, float]:
        x_min, y_min, x_max, y_max = self._spread_bounds_for_sample(layout, sample)
        usable_w = max(sample.width_in, x_max - x_min + sample.width_in)
        usable_h = max(sample.height_in, y_max - y_min + sample.height_in)
        aspect = max(0.25, usable_w / max(usable_h, 0.01))
        cols = max(1, min(total, math.ceil(math.sqrt(total * aspect))))
        rows = max(1, math.ceil(total / cols))
        col = index % cols
        row = index // cols
        target_x = self._grid_slot_center(x_min, x_max, sample.width_in, col, cols)
        target_y = self._grid_slot_center(y_min, y_max, sample.height_in, row, rows)
        return target_x, target_y

    @staticmethod
    def _grid_slot_center(start: float, stop: float, length: float, index: int, count: int) -> float:
        if count <= 1 or stop <= start:
            return start + length / 2
        return start + (stop - start) * index / (count - 1) + length / 2

    def _best_spread_candidate(
        self,
        layout: LayoutResult,
        sample: SampleRect,
        target: Tuple[float, float],
        require_spacing: bool,
    ) -> Optional[Tuple[float, float]]:
        candidates = self._spread_candidate_positions(layout, sample)
        best_score = None
        best_candidate = None
        for x, y in candidates:
            if self._validate_sample_position(sample, x, y) is not None:
                continue
            if require_spacing and not self._spread_has_clearance(sample, x, y, layout.samples):
                continue
            score = self._spread_candidate_score(layout, sample, x, y, target)
            if best_score is None or score > best_score:
                best_score = score
                best_candidate = (x, y)
        return best_candidate

    def _spread_candidate_positions(self, layout: LayoutResult, sample: SampleRect) -> List[Tuple[float, float]]:
        x_min, y_min, x_max, y_max = self._spread_bounds_for_sample(layout, sample)
        if x_max < x_min - 1e-6 or y_max < y_min - 1e-6:
            return []

        x_values = self._spread_axis_values(x_min, x_max, 9)
        y_values = self._spread_axis_values(y_min, y_max, 9)

        if sample.metadata.get("centered_width") == "true":
            x_values = [self._snap_for_fill((layout.board_width_in - sample.width_in) / 2)]

        if sample_requires_formed_edge(sample):
            if layout.board_height_in >= layout.board_width_in:
                edge_values = [0.0, layout.board_width_in - sample.width_in]
                x_values = [self._snap_for_fill(x) for x in edge_values if x_min - 1e-6 <= x <= x_max + 1e-6]
            else:
                edge_values = [0.0, layout.board_height_in - sample.height_in]
                y_values = [self._snap_for_fill(y) for y in edge_values if y_min - 1e-6 <= y <= y_max + 1e-6]

        candidates = []
        seen = set()
        for y in y_values:
            for x in x_values:
                key = (round(x, 4), round(y, 4))
                if key not in seen:
                    seen.add(key)
                    candidates.append(key)
        return candidates

    def _spread_bounds_for_sample(self, layout: LayoutResult, sample: SampleRect) -> Tuple[float, float, float, float]:
        x_min = 0.0
        y_min = 0.0
        x_max = layout.board_width_in - sample.width_in
        y_max = layout.board_height_in - sample.height_in

        for zone_x, zone_y, zone_w, zone_h in layout.scrap_zones:
            spans_height = zone_y <= 1e-6 and zone_y + zone_h >= layout.board_height_in - 1e-6
            spans_width = zone_x <= 1e-6 and zone_x + zone_w >= layout.board_width_in - 1e-6
            if spans_height and zone_x <= 1e-6:
                x_min = max(x_min, zone_x + zone_w)
            elif spans_height and zone_x + zone_w >= layout.board_width_in - 1e-6:
                x_max = min(x_max, zone_x - sample.width_in)
            elif spans_width and zone_y <= 1e-6:
                y_min = max(y_min, zone_y + zone_h)
            elif spans_width and zone_y + zone_h >= layout.board_height_in - 1e-6:
                y_max = min(y_max, zone_y - sample.height_in)

        min_margin = sample.metadata.get("min_margin")
        if min_margin:
            margin = float(min_margin)
            x_min = max(x_min, margin)
            y_min = max(y_min, margin)
            x_max = min(x_max, layout.board_width_in - margin - sample.width_in)
            y_max = min(y_max, layout.board_height_in - margin - sample.height_in)
        elif sample_uses_shop_edge_inset(sample):
            margin = SHOP_EDGE_INSET_IN
            x_min = max(x_min, margin)
            y_min = max(y_min, margin)
            x_max = min(x_max, layout.board_width_in - margin - sample.width_in)
            y_max = min(y_max, layout.board_height_in - margin - sample.height_in)

        min_end = sample.metadata.get("min_end")
        if min_end:
            margin = float(min_end)
            if layout.board_height_in >= layout.board_width_in:
                y_min = max(y_min, margin)
                y_max = min(y_max, layout.board_height_in - margin - sample.height_in)
            else:
                x_min = max(x_min, margin)
                x_max = min(x_max, layout.board_width_in - margin - sample.width_in)

        if sample.sample_id.startswith("DTP-11"):
            rule = RULES_JSON["DTP-11"]
            side_offset = float(rule["side_offset_in"])
            end_offset = float(rule["end_offset_in"])
            x_min = max(x_min, side_offset)
            x_max = min(x_max, layout.board_width_in - side_offset - sample.width_in)
            y_min = max(y_min, end_offset)
            y_max = min(y_max, layout.board_height_in - end_offset - sample.height_in)

        return x_min, y_min, x_max, y_max

    def _spread_axis_values(self, start: float, stop: float, count: int) -> List[float]:
        if stop < start - 1e-6:
            return []
        if stop <= start + 1e-6 or count <= 1:
            return [self._snap_for_fill(start)]
        return [self._snap_for_fill(start + (stop - start) * i / (count - 1)) for i in range(count)]

    def _snap_for_fill(self, value: float) -> float:
        return round(value * 4) / 4

    def _spread_has_clearance(self, sample: SampleRect, x: float, y: float, placed: List[SampleRect]) -> bool:
        gap = 0.25
        for other in placed:
            if self._rects_overlap(
                x - gap,
                y - gap,
                sample.width_in + gap * 2,
                sample.height_in + gap * 2,
                other.x_in,
                other.y_in,
                other.width_in,
                other.height_in,
            ):
                return False
        return True

    def _spread_candidate_score(
        self,
        layout: LayoutResult,
        sample: SampleRect,
        x: float,
        y: float,
        target: Tuple[float, float],
    ) -> float:
        cx = x + sample.width_in / 2
        cy = y + sample.height_in / 2
        if layout.samples:
            nearest = min(
                math.hypot(cx - (other.x_in + other.width_in / 2), cy - (other.y_in + other.height_in / 2))
                for other in layout.samples
            )
        else:
            nearest = min(layout.board_width_in, layout.board_height_in) / 2
        target_distance = math.hypot(cx - target[0], cy - target[1])
        board_center_distance = math.hypot(cx - layout.board_width_in / 2, cy - layout.board_height_in / 2)
        edge_balance = min(cx, layout.board_width_in - cx, cy, layout.board_height_in - cy)
        return nearest * 2.0 - target_distance * 0.65 - board_center_distance * 0.04 + edge_balance * 0.08

    def _rebuild_score_lines(self) -> None:
        if not self.layout_result:
            return
        self.layout_result.line_entities = [line for line in self.layout_result.line_entities if line.layer != "STP308_SCORE"]
        for sample in self.layout_result.samples:
            if sample.sample_id.startswith("STP308") or sample.metadata.get("edge_taken") == "true":
                score_y = sample.y_in + float(MANUAL_OBJECTS["STP308"]["score_offset"])
                self.layout_result.line_entities.append(
                    LineEntity("STP308_SCORE", sample.x_in, score_y, sample.x_in + sample.width_in, score_y)
                )

    def fit_view(self) -> None:
        self.view_zoom = 1.0
        self.measure_start = None
        self.draw_layout()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def zoom_view(self, factor: float) -> None:
        self.view_zoom = max(0.25, min(8.0, self.view_zoom * factor))
        self.draw_layout()

    def on_canvas_mousewheel(self, event) -> None:
        delta = getattr(event, "delta", 0)
        if delta == 0:
            button_num = getattr(event, "num", None)
            if button_num == 4:
                delta = 120
            elif button_num == 5:
                delta = -120
            else:
                return

        state = getattr(event, "state", 0)
        if state & 0x0001:
            self.canvas.xview_scroll(-1 if delta > 0 else 1, "units")
        else:
            self.zoom_view_at(1.12 if delta > 0 else 1 / 1.12, event)

    def zoom_view_at(self, factor: float, event) -> None:
        focus = self._event_to_board_inches(event)
        self.view_zoom = max(0.25, min(8.0, self.view_zoom * factor))
        self.draw_layout()
        if not focus:
            return
        target_x = self.canvas_origin[0] + focus[0] * self.canvas_scale
        target_y = self.canvas_origin[1] + focus[1] * self.canvas_scale
        scroll_region = self.canvas.cget("scrollregion")
        if not scroll_region:
            return
        try:
            _left, _top, right, bottom = [float(value) for value in str(scroll_region).split()]
        except ValueError:
            bbox = self.canvas.bbox("all")
            if not bbox:
                return
            _left, _top, right, bottom = [float(value) for value in bbox]
        if right > 0:
            self.canvas.xview_moveto(max(0.0, min(1.0, (target_x - event.x) / right)))
        if bottom > 0:
            self.canvas.yview_moveto(max(0.0, min(1.0, (target_y - event.y) / bottom)))

    def on_canvas_pan_start(self, event) -> None:
        self.panning_canvas = True
        self.drag_sample_id = None
        self.drag_last_xy = None
        self.canvas.scan_mark(event.x, event.y)
        try:
            self.canvas.configure(cursor="fleur")
        except tk.TclError:
            self.canvas.configure(cursor="hand2")

    def on_canvas_pan_drag(self, event) -> None:
        if not self.panning_canvas:
            return
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def on_canvas_pan_end(self, _event) -> None:
        self.panning_canvas = False
        self.canvas.configure(cursor="")

    def toggle_measure_mode(self) -> None:
        self.measure_start = None
        self.draw_layout()

    def _event_to_board_inches(self, event) -> Optional[Tuple[float, float]]:
        if not self.layout_result:
            return None
        bx0, by0 = self.canvas_origin
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        x_in = (canvas_x - bx0) / self.canvas_scale
        y_in = (canvas_y - by0) / self.canvas_scale
        if 0 <= x_in <= self.layout_result.board_width_in and 0 <= y_in <= self.layout_result.board_height_in:
            return x_in, y_in
        return None

    def handle_measure_click(self, event) -> None:
        point = self._event_to_board_inches(event)
        if not point:
            return
        if not self.measure_start:
            self.measure_start = point
            self.sample_info.delete("1.0", tk.END)
            self.sample_info.insert("1.0", f"Measure start: ({point[0]:.2f}, {point[1]:.2f}) in")
            self.draw_layout()
            return

        sx, sy = self.measure_start
        ex, ey = point
        dx = ex - sx
        dy = ey - sy
        distance = math.hypot(dx, dy)
        bx0, by0 = self.canvas_origin
        self.draw_layout()
        x1 = bx0 + sx * self.canvas_scale
        y1 = by0 + sy * self.canvas_scale
        x2 = bx0 + ex * self.canvas_scale
        y2 = by0 + ey * self.canvas_scale
        self.canvas.create_line(x1, y1, x2, y2, fill="#dc2626", width=3, dash=(6, 4), tags=("measure",))
        self.canvas.create_oval(x2 - 4, y2 - 4, x2 + 4, y2 + 4, fill="#dc2626", outline="#991b1b", tags=("measure",))
        self.canvas.create_text(
            (x1 + x2) / 2,
            (y1 + y2) / 2 - 10,
            text=f"{distance:.2f} in  DX {dx:.2f}  DY {dy:.2f}",
            fill="#991b1b",
            font=("Segoe UI", 9, "bold"),
            tags=("measure",),
        )
        self.sample_info.delete("1.0", tk.END)
        self.sample_info.insert(
            "1.0",
            f"Measured distance: {distance:.2f} in\n"
            f"Delta X: {dx:.2f} in\n"
            f"Delta Y: {dy:.2f} in\n"
            f"Start: ({sx:.2f}, {sy:.2f}) in\n"
            f"End: ({ex:.2f}, {ey:.2f}) in",
        )
        self.measure_start = None

    def insert_text_note(self) -> None:
        try:
            layout = self._ensure_working_layout()
            text = self.insert_text_var.get().strip()
            if not text:
                raise ValueError("Text cannot be blank.")
            x = float(self.insert_text_x_var.get())
            y = float(self.insert_text_y_var.get())
            height = float(self.insert_text_height_var.get())
            if x < 0 or y < 0 or x > layout.board_width_in or y > layout.board_height_in:
                raise ValueError("Text position must be inside the board.")
            self.push_undo("Insert text")
            layout.text_entities.append(TextEntity("LABELS", x, y, height, text))
            self.draw_layout()
            self.sample_info.delete("1.0", tk.END)
            self.sample_info.insert("1.0", f"Inserted text at ({x:.2f}, {y:.2f}) in: {text}")
        except Exception as exc:
            messagebox.showerror("Insert Text Error", str(exc))

    def on_canvas_motion(self, event) -> None:
        if not self.layout_result:
            self.cursor_label.configure(text="X: --  Y: --")
            return
        point = self._event_to_board_inches(event)
        if point:
            x_in, y_in = point
            self.cursor_label.configure(text=f"X: {x_in:.2f}  Y: {y_in:.2f}")
        else:
            self.cursor_label.configure(text="X: --  Y: --")

    def on_part_tree_selected(self, _event=None) -> None:
        if self.updating_part_tree:
            return
        selected = self.part_tree.selection()
        if not selected:
            return
        selected_ids = set(selected)
        if selected_ids == self.selected_sample_ids:
            return
        self.select_samples(selected_ids, sync_tree=False)

    def _get_sample(self, sample_id: str) -> Optional[SampleRect]:
        if not self.layout_result:
            return None
        return next((s for s in self.layout_result.samples if s.sample_id == sample_id), None)

    def on_canvas_press(self, event) -> None:
        try:
            if not self.layout_result:
                return
            if self.measure_mode_var.get():
                self.handle_measure_click(event)
                return
            sample_id = self._sample_id_at_event(event)
            if not sample_id:
                self.drag_sample_id = None
                self.drag_last_xy = None
                self.start_marquee_selection(event)
                return
            self.select_sample(sample_id, sync_tree=True)
            self.drag_sample_id = sample_id
            self.drag_last_xy = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
            self.last_drag_error = None
            self.drag_undo_recorded = False
        except Exception as exc:
            self.drag_sample_id = None
            self.drag_last_xy = None
            self.sample_info.delete("1.0", tk.END)
            self.sample_info.insert("1.0", f"Click error: {exc}")

    def start_marquee_selection(self, event) -> None:
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        self.marquee_start_canvas = (canvas_x, canvas_y)
        self.marquee_dragging = True
        if self.marquee_rect_id:
            self.canvas.delete(self.marquee_rect_id)
        self.marquee_rect_id = self.canvas.create_rectangle(
            canvas_x,
            canvas_y,
            canvas_x,
            canvas_y,
            outline="#0f766e",
            width=2,
            dash=(5, 3),
            tags=("marquee",),
        )

    def update_marquee_selection(self, event) -> None:
        if not self.marquee_dragging or not self.marquee_start_canvas or not self.marquee_rect_id:
            return
        start_x, start_y = self.marquee_start_canvas
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.marquee_rect_id, start_x, start_y, current_x, current_y)

    def finish_marquee_selection(self, event) -> None:
        if not self.marquee_dragging or not self.marquee_start_canvas:
            return
        start_x, start_y = self.marquee_start_canvas
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        if self.marquee_rect_id:
            self.canvas.delete(self.marquee_rect_id)
        self.marquee_start_canvas = None
        self.marquee_rect_id = None
        self.marquee_dragging = False

        if abs(end_x - start_x) < 4 and abs(end_y - start_y) < 4:
            self.clear_selection()
            return

        bx0, by0 = self.canvas_origin
        x1 = (min(start_x, end_x) - bx0) / self.canvas_scale
        y1 = (min(start_y, end_y) - by0) / self.canvas_scale
        x2 = (max(start_x, end_x) - bx0) / self.canvas_scale
        y2 = (max(start_y, end_y) - by0) / self.canvas_scale
        selected = self._sample_ids_in_rect(x1, y1, x2, y2)
        self.select_samples(selected)

    def _sample_ids_in_rect(self, x1: float, y1: float, x2: float, y2: float) -> Set[str]:
        if not self.layout_result:
            return set()
        selected = set()
        for sample in self.layout_result.samples:
            intersects = (
                sample.x_in < x2
                and sample.x_in + sample.width_in > x1
                and sample.y_in < y2
                and sample.y_in + sample.height_in > y1
            )
            if intersects:
                selected.add(sample.sample_id)
        return selected

    def _sample_id_at_event(self, event) -> Optional[str]:
        if not self.layout_result:
            return None
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        bx0, by0 = self.canvas_origin
        x_in = (canvas_x - bx0) / self.canvas_scale
        y_in = (canvas_y - by0) / self.canvas_scale

        for sample in reversed(self.layout_result.samples):
            if (
                sample.x_in <= x_in <= sample.x_in + sample.width_in
                and sample.y_in <= y_in <= sample.y_in + sample.height_in
            ):
                return sample.sample_id
        return None

    def select_sample(self, sample_id: str, sync_tree: bool = True) -> None:
        self.select_samples({sample_id}, sync_tree=sync_tree)

    def select_samples(self, sample_ids: Set[str], sync_tree: bool = True) -> None:
        if not self.layout_result:
            return
        sample_ids = {sample_id for sample_id in sample_ids if self._get_sample(sample_id)}
        self.selected_sample_ids = set(sample_ids)
        self.selected_sample_id = next(iter(sample_ids), None)
        self.selection_count_var.set(f"Selected: {len(sample_ids)}")
        if sync_tree and hasattr(self, "part_tree"):
            self.updating_part_tree = True
            try:
                existing_ids = [sample_id for sample_id in sample_ids if sample_id in self.part_tree.get_children()]
                self.part_tree.selection_set(existing_ids)
                if existing_ids:
                    self.part_tree.see(existing_ids[-1])
            finally:
                self.updating_part_tree = False
        if len(sample_ids) == 1 and self.selected_sample_id:
            self.show_sample_details(self.selected_sample_id)
            self.draw_animation_placeholder(self.selected_sample_id)
        elif sample_ids:
            self.show_multi_sample_details(sample_ids)
            self.draw_animation_placeholder()
        else:
            self.sample_info.delete("1.0", tk.END)
            self.draw_animation_placeholder()
        self.highlight_selected_sample()

    def clear_selection(self) -> None:
        self.selected_sample_id = None
        self.selected_sample_ids.clear()
        self.selection_count_var.set("Selected: 0")
        self.selected_x_var.set("")
        self.selected_y_var.set("")
        self.selected_w_var.set("")
        self.selected_h_var.set("")
        if hasattr(self, "part_tree"):
            self.updating_part_tree = True
            try:
                self.part_tree.selection_remove(self.part_tree.selection())
            finally:
                self.updating_part_tree = False
        self.highlight_selected_sample()

    def _legacy_select_sample(self, sample_id: str, sync_tree: bool = True) -> None:
        self.selected_sample_id = sample_id
        if sync_tree and hasattr(self, "part_tree") and sample_id in self.part_tree.get_children():
            self.updating_part_tree = True
            try:
                self.part_tree.selection_set(sample_id)
                self.part_tree.see(sample_id)
            finally:
                self.updating_part_tree = False
        self.show_sample_details(sample_id)
        self.draw_animation_placeholder(sample_id)
        self.highlight_selected_sample()

    def highlight_selected_sample(self) -> None:
        try:
            for item in self.canvas.find_withtag("sample_rect"):
                self.canvas.itemconfigure(item, fill="#bfdbfe", outline="#1d4ed8", width=2)
            for sample_id in self.selected_sample_ids:
                for item in self.canvas.find_withtag(sample_id):
                    tags = self.canvas.gettags(item)
                    if "sample_rect" in tags:
                        self.canvas.itemconfigure(item, fill="#fde68a", outline="#b45309", width=4)
        except tk.TclError:
            pass

    def show_sample_details(self, sample_id: str) -> None:
        if not self.layout_result:
            return
        sample = self._get_sample(sample_id)
        if not sample:
            return
        info = [
            f"Selected sample: {sample.sample_id}",
            f"Position: ({sample.x_in:.2f}, {sample.y_in:.2f}) in",
            f"Size: {sample.width_in:.2f} x {sample.height_in:.2f} in",
            f"Rotation: {sample.rotation_deg} deg",
        ]
        self.selected_x_var.set(f"{sample.x_in:.3f}")
        self.selected_y_var.set(f"{sample.y_in:.3f}")
        self.selected_w_var.set(f"{sample.width_in:.3f}")
        self.selected_h_var.set(f"{sample.height_in:.3f}")
        if sample.metadata.get("locked") == "true":
            info.append("Locked: Yes")
        if sample.drill_centers:
            for i, (cx, cy, dia) in enumerate(sample.drill_centers, start=1):
                info.append(f"Hole {i}: center=({cx:.2f}, {cy:.2f}) in, dia={dia:.2f} in")
        if sample.metadata:
            info.append(f"Metadata: {json.dumps(sample.metadata)}")
        self.sample_info.delete("1.0", tk.END)
        self.sample_info.insert("1.0", "\n".join(info))

    def show_multi_sample_details(self, sample_ids: Set[str]) -> None:
        if not self.layout_result:
            return
        samples = [sample for sample in self.layout_result.samples if sample.sample_id in sample_ids]
        if not samples:
            return
        min_x = min(sample.x_in for sample in samples)
        min_y = min(sample.y_in for sample in samples)
        max_x = max(sample.x_in + sample.width_in for sample in samples)
        max_y = max(sample.y_in + sample.height_in for sample in samples)
        self.selected_x_var.set("")
        self.selected_y_var.set("")
        self.selected_w_var.set("")
        self.selected_h_var.set("")
        info = [
            f"Selected samples: {len(samples)}",
            f"Group bounds: X {min_x:.2f} to {max_x:.2f}, Y {min_y:.2f} to {max_y:.2f}",
            f"Group size: {(max_x - min_x):.2f} x {(max_y - min_y):.2f} in",
            "",
            "Selected IDs:",
        ]
        info.extend(f"- {sample.sample_id}" for sample in samples[:20])
        if len(samples) > 20:
            info.append(f"...and {len(samples) - 20} more")
        self.sample_info.delete("1.0", tk.END)
        self.sample_info.insert("1.0", "\n".join(info))

    def _selected_samples(self) -> List[SampleRect]:
        if not self.layout_result:
            return []
        return [sample for sample in self.layout_result.samples if sample.sample_id in self.selected_sample_ids]

    def set_selected_distance_from_edge(self) -> None:
        if not self.layout_result:
            return
        samples = self._selected_samples()
        if not samples:
            messagebox.showinfo("No Selection", "Drag a selection box or click a sample first.")
            return
        try:
            distance = float(self.edge_distance_var.get())
            if distance < 0:
                raise ValueError("Distance must be zero or greater.")
            side = self.edge_side_var.get()
            self.push_undo("Set selected edge distance")
            moved = []
            for sample in samples:
                if sample.metadata.get("locked") == "true":
                    continue
                new_x = sample.x_in
                new_y = sample.y_in
                if side == "Left":
                    new_x = distance
                elif side == "Right":
                    new_x = self.layout_result.board_width_in - distance - sample.width_in
                elif side == "Top":
                    new_y = distance
                elif side == "Bottom":
                    new_y = self.layout_result.board_height_in - distance - sample.height_in
                if self.snap_enabled_var.get():
                    new_x = self._snap_value(new_x)
                    new_y = self._snap_value(new_y)
                bounds_error = self._sample_bounds_error(sample, new_x, new_y)
                if bounds_error:
                    raise ValueError(f"{sample.sample_id}: {bounds_error}")
                self._move_sample(sample, new_x, new_y)
                moved.append(sample.sample_id)
            self._rebuild_score_lines()
            self.draw_layout()
            self.select_samples(set(moved) or self.selected_sample_ids)
            self.show_warnings()
            if moved:
                self.sample_info.insert(tk.END, f"\n\nMoved {len(moved)} sample(s) to {distance:g} in from {side} edge.")
        except Exception as exc:
            self.undo()
            messagebox.showerror("Set Edge Distance Error", str(exc))

    def _sample_bounds_error(self, sample: SampleRect, new_x: float, new_y: float) -> Optional[str]:
        if not self.layout_result:
            return "No active layout."
        eps = 1e-6
        if new_x < -eps or new_y < -eps:
            return "sample cannot leave the board."
        if new_x + sample.width_in > self.layout_result.board_width_in + eps:
            return "sample cannot leave the board width."
        if new_y + sample.height_in > self.layout_result.board_height_in + eps:
            return "sample cannot leave the board height."
        return None

    def apply_selected_part_edits(self) -> None:
        if not self.layout_result or not self.selected_sample_id:
            return
        sample = self._get_sample(self.selected_sample_id)
        if not sample:
            return
        try:
            new_x = float(self.selected_x_var.get())
            new_y = float(self.selected_y_var.get())
            new_w = float(self.selected_w_var.get())
            new_h = float(self.selected_h_var.get())
            if new_w <= 0 or new_h <= 0:
                raise ValueError("Width and height must be greater than zero.")
            if self.snap_enabled_var.get():
                new_x = self._snap_value(new_x)
                new_y = self._snap_value(new_y)
            old_w, old_h = sample.width_in, sample.height_in
            sample.width_in = new_w
            sample.height_in = new_h
            blocked = self._validate_sample_position(sample, new_x, new_y, allow_locked=True)
            if blocked:
                sample.width_in, sample.height_in = old_w, old_h
                raise ValueError(blocked)
            self.push_undo("Edit selected part")
            self._move_sample(sample, new_x, new_y)
            sample.width_in = new_w
            sample.height_in = new_h
            self.draw_layout()
            self.show_sample_details(sample.sample_id)
            self.show_warnings()
        except Exception as exc:
            messagebox.showerror("Part Edit Error", str(exc))

    def duplicate_selected_sample(self) -> None:
        if not self.layout_result or not self.selected_sample_id:
            return
        source = self._get_sample(self.selected_sample_id)
        if not source:
            return
        self.push_undo("Duplicate sample")
        duplicate = copy.deepcopy(source)
        base = re.sub(r"-copy\d+$", "", source.sample_id)
        duplicate.sample_id = f"{base}-copy{self._next_part_number(base + '-copy')}"
        duplicate.metadata.pop("locked", None)
        duplicate.drill_centers = [(cx - source.x_in, cy - source.y_in, dia) for cx, cy, dia in source.drill_centers]
        duplicate.x_in = 0
        duplicate.y_in = 0
        self._place_sample_first_fit(self.layout_result, duplicate)
        self.layout_result.samples.append(duplicate)
        self.draw_layout()
        self.select_sample(duplicate.sample_id)
        self.show_warnings()

    def delete_selected_sample(self) -> None:
        if not self.layout_result or not self.selected_sample_ids:
            return
        self.push_undo("Delete sample")
        delete_ids = set(self.selected_sample_ids)
        self.layout_result.samples = [s for s in self.layout_result.samples if s.sample_id not in delete_ids]
        self.selected_sample_id = None
        self.selected_sample_ids.clear()
        self.selection_count_var.set("Selected: 0")
        self.draw_layout()
        self.show_warnings()

    def rotate_selected_sample(self) -> None:
        if not self.layout_result or not self.selected_sample_id:
            return
        sample = self._get_sample(self.selected_sample_id)
        if not sample:
            return
        self.push_undo("Rotate sample")
        sample.width_in, sample.height_in = sample.height_in, sample.width_in
        sample.rotation_deg = (sample.rotation_deg + 90) % 360
        blocked = self._validate_sample_position(sample, sample.x_in, sample.y_in, allow_locked=True)
        if blocked:
            self.undo()
            messagebox.showerror("Rotate Error", blocked)
            return
        self.draw_layout()
        self.select_sample(sample.sample_id)

    def toggle_selected_lock(self) -> None:
        if not self.layout_result or not self.selected_sample_id:
            return
        sample = self._get_sample(self.selected_sample_id)
        if not sample:
            return
        self.push_undo("Toggle lock")
        if sample.metadata.get("locked") == "true":
            sample.metadata.pop("locked", None)
        else:
            sample.metadata["locked"] = "true"
        self.draw_layout()
        self.show_sample_details(sample.sample_id)

    def on_canvas_drag(self, event) -> None:
        if self.measure_mode_var.get():
            return
        if self.marquee_dragging:
            self.update_marquee_selection(event)
            return
        if not self.layout_result or not self.drag_sample_id or not self.drag_last_xy:
            return
        sample = self._get_sample(self.drag_sample_id)
        if not sample:
            return

        last_x, last_y = self.drag_last_xy
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        dx_px = current_x - last_x
        dy_px = current_y - last_y
        if abs(dx_px) < 1 and abs(dy_px) < 1:
            return

        dx_in = dx_px / self.canvas_scale
        dy_in = dy_px / self.canvas_scale
        new_x = sample.x_in + dx_in
        new_y = sample.y_in + dy_in
        if self.snap_enabled_var.get():
            new_x = self._snap_value(new_x)
            new_y = self._snap_value(new_y)
        blocked_reason = self._validate_sample_position(sample, new_x, new_y)
        if blocked_reason:
            if blocked_reason != self.last_drag_error:
                self.last_drag_error = blocked_reason
                self.show_sample_details(sample.sample_id)
                self.sample_info.insert(tk.END, f"\n\n{blocked_reason}")
            return

        if not self.drag_undo_recorded:
            self.push_undo("Move sample")
            self.drag_undo_recorded = True
        self._move_sample(sample, new_x, new_y)
        self.draw_layout()
        self.drag_last_xy = (current_x, current_y)
        self.last_drag_error = None
        self.show_sample_details(sample.sample_id)

    def on_canvas_release(self, event) -> None:
        if self.marquee_dragging:
            self.finish_marquee_selection(event)
            return
        if self.drag_sample_id:
            self.show_sample_details(self.drag_sample_id)
        self.drag_sample_id = None
        self.drag_last_xy = None
        self.last_drag_error = None
        self.drag_undo_recorded = False

    def _move_sample(self, sample: SampleRect, new_x: float, new_y: float) -> None:
        dx = new_x - sample.x_in
        dy = new_y - sample.y_in
        old_x = sample.x_in
        old_y = sample.y_in
        old_w = sample.width_in
        old_h = sample.height_in
        sample.x_in = new_x
        sample.y_in = new_y
        sample.drill_centers = [(cx + dx, cy + dy, dia) for cx, cy, dia in sample.drill_centers]
        if self.layout_result:
            for line in self.layout_result.line_entities:
                line_inside_sample = (
                    old_x - 1e-6 <= line.x1_in <= old_x + old_w + 1e-6
                    and old_x - 1e-6 <= line.x2_in <= old_x + old_w + 1e-6
                    and old_y - 1e-6 <= line.y1_in <= old_y + old_h + 1e-6
                    and old_y - 1e-6 <= line.y2_in <= old_y + old_h + 1e-6
                )
                if line_inside_sample and line.layer == "STP308_SCORE":
                    line.x1_in += dx
                    line.x2_in += dx
                    line.y1_in += dy
                    line.y2_in += dy

    def _snap_value(self, value: float) -> float:
        try:
            increment = float(self.snap_increment_var.get())
        except ValueError:
            increment = 0.25
        if increment <= 0:
            return value
        return round(value / increment) * increment

    def _validate_sample_position(self, sample: SampleRect, new_x: float, new_y: float, allow_locked: bool = False) -> Optional[str]:
        if not self.layout_result:
            return "Move blocked: no active layout."

        eps = 1e-6
        if sample.metadata.get("locked") == "true" and not allow_locked:
            return f"Move blocked: {sample.sample_id} is locked."
        if new_x < -eps or new_y < -eps:
            return "Move blocked: sample cannot leave the board."
        if new_x + sample.width_in > self.layout_result.board_width_in + eps:
            return "Move blocked: sample cannot leave the board."
        if new_y + sample.height_in > self.layout_result.board_height_in + eps:
            return "Move blocked: sample cannot leave the board."

        for zone_x, zone_y, zone_w, zone_h in self.layout_result.scrap_zones:
            if self._rects_overlap(new_x, new_y, sample.width_in, sample.height_in, zone_x, zone_y, zone_w, zone_h):
                return f"Move blocked: {sample.sample_id} cannot be placed in a discard zone."

        min_margin = sample.metadata.get("min_margin")
        if min_margin:
            margin = float(min_margin)
            if (
                new_x < margin - eps
                or new_y < margin - eps
                or new_x + sample.width_in > self.layout_result.board_width_in - margin + eps
                or new_y + sample.height_in > self.layout_result.board_height_in - margin + eps
            ):
                return f"Move blocked: {sample.sample_id} must stay at least {margin:g} in from board edges/ends."
        elif sample_uses_shop_edge_inset(sample):
            margin = SHOP_EDGE_INSET_IN
            if (
                new_x < margin - eps
                or new_y < margin - eps
                or new_x + sample.width_in > self.layout_result.board_width_in - margin + eps
                or new_y + sample.height_in > self.layout_result.board_height_in - margin + eps
            ):
                return f"Move blocked: {sample.sample_id} must stay at least {margin:g} in off board edges unless it is edge shear."

        min_end = sample.metadata.get("min_end")
        if min_end:
            margin = float(min_end)
            if self.layout_result.board_height_in >= self.layout_result.board_width_in:
                if new_y < margin - eps or new_y + sample.height_in > self.layout_result.board_height_in - margin + eps:
                    return f"Move blocked: {sample.sample_id} must stay at least {margin:g} in from board ends."
            elif new_x < margin - eps or new_x + sample.width_in > self.layout_result.board_width_in - margin + eps:
                return f"Move blocked: {sample.sample_id} must stay at least {margin:g} in from board ends."

        if sample.metadata.get("centered_width") == "true":
            centered_x = (self.layout_result.board_width_in - sample.width_in) / 2
            if abs(new_x - centered_x) > 0.05:
                return f"Move blocked: {sample.sample_id} must stay centered across the board width."

        if sample_requires_formed_edge(sample):
            if self.layout_result.board_height_in >= self.layout_result.board_width_in:
                on_edge = abs(new_x) <= 0.05 or abs(new_x + sample.width_in - self.layout_result.board_width_in) <= 0.05
            else:
                on_edge = abs(new_y) <= 0.05 or abs(new_y + sample.height_in - self.layout_result.board_height_in) <= 0.05
            if not on_edge:
                return f"Move blocked: {sample.sample_id} must stay on a formed/paper edge."

        if sample.sample_id.startswith("DTP-11"):
            rule = RULES_JSON["DTP-11"]
            side_offset = rule["side_offset_in"]
            end_offset = rule["end_offset_in"]
            if new_x < side_offset - eps or new_x + sample.width_in > self.layout_result.board_width_in - side_offset + eps:
                return "Move blocked: DTP-11 samples must stay out of the side discard zones."
            if new_y < end_offset - eps or new_y + sample.height_in > self.layout_result.board_height_in - end_offset + eps:
                return "Move blocked: DTP-11 samples must stay out of the end discard zones."

        for other in self.layout_result.samples:
            if other.sample_id == sample.sample_id:
                continue
            if self._rects_overlap(
                new_x,
                new_y,
                sample.width_in,
                sample.height_in,
                other.x_in,
                other.y_in,
                other.width_in,
                other.height_in,
            ):
                return f"Move blocked: {sample.sample_id} overlaps {other.sample_id}."

        return None

    @staticmethod
    def _rects_overlap(
        ax: float,
        ay: float,
        aw: float,
        ah: float,
        bx: float,
        by: float,
        bw: float,
        bh: float,
    ) -> bool:
        eps = 1e-6
        return ax < bx + bw - eps and ax + aw > bx + eps and ay < by + bh - eps and ay + ah > by + eps

    def draw_animation_placeholder(self, sample_id: Optional[str] = None) -> None:
        self.animation_canvas.delete("all")
        dtp_id = self.setup_preview_dtp_id or self.current_dtp_id
        if sample_id:
            dtp_id = next((rule_id for rule_id in RULES_JSON if sample_id.startswith(rule_id)), dtp_id)
            if sample_id.startswith("DTP16"):
                dtp_id = "DTP-16"
        if dtp_id not in RULES_JSON:
            dtp_id = self.current_dtp_id
        self.setup_preview_dtp_id = dtp_id

        self.animation_canvas.create_text(130, 20, text=f"{dtp_id} Setup", font=("Segoe UI", 11, "bold"))
        if dtp_id == "DTP-11":
            self.animation_canvas.create_rectangle(40, 70, 220, 180, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_rectangle(95, 95, 165, 165, fill="#f59e0b", outline="#92400e", width=2)
            self.animation_canvas.create_text(130, 195, text="Center wooden block on specimen")
        elif dtp_id == "DTP-13":
            self.animation_canvas.create_rectangle(45, 55, 215, 185, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_oval(125, 115, 135, 125, outline="#111827", width=2)
            self.animation_canvas.create_oval(102, 92, 158, 148, outline="#6b7280", width=2)
            self.animation_canvas.create_text(130, 200, text="Disc + center hole + fastener setup")
        elif dtp_id == "DTP-15":
            self.animation_canvas.create_rectangle(50, 100, 210, 140, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_line(30, 80, 230, 80, arrow=tk.LAST, width=3)
            self.animation_canvas.create_text(130, 65, text="Machine Direction")
            self.animation_canvas.create_text(130, 180, text="Long axis must follow MD")
        elif dtp_id == "DTP-16":
            self.animation_canvas.create_rectangle(82, 78, 178, 174, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_rectangle(70, 174, 190, 190, fill="#cbd5e1", outline="#64748b")
            self.animation_canvas.create_oval(113, 109, 147, 143, outline="#dc2626", width=3)
            self.animation_canvas.create_line(130, 52, 130, 103, arrow=tk.LAST, fill="#dc2626", width=3)
            self.animation_canvas.create_text(130, 62, text="Drop weight")
            self.animation_canvas.create_text(130, 206, text="Center impact on 6 x 6 sample")
        elif dtp_id == "STP308":
            self.animation_canvas.create_rectangle(50, 82, 210, 156, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_rectangle(50, 82, 72, 156, fill="#bfdbfe", outline="#1d4ed8")
            self.animation_canvas.create_line(86, 82, 86, 156, fill="#dc2626", width=3, dash=(5, 3))
            self.animation_canvas.create_line(50, 68, 210, 68, arrow=tk.LAST, width=2)
            self.animation_canvas.create_text(130, 55, text="6.0 in length")
            self.animation_canvas.create_text(61, 170, text="edge", font=("Segoe UI", 8))
            self.animation_canvas.create_text(148, 184, text="Score line 1.25 in from leading end")
        elif dtp_id == "STP312":
            self.animation_canvas.create_rectangle(30, 58, 230, 178, fill="#f8fafc", outline="#94a3b8", width=2)
            self.animation_canvas.create_rectangle(52, 82, 120, 150, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_rectangle(144, 76, 202, 158, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_line(52, 162, 120, 162, arrow=tk.LAST, fill="#16a34a", width=3)
            self.animation_canvas.create_line(214, 158, 214, 76, arrow=tk.LAST, fill="#16a34a", width=3)
            self.animation_canvas.create_text(86, 174, text="CD")
            self.animation_canvas.create_text(178, 174, text="MD")
            self.animation_canvas.create_text(130, 204, text="Keep flexural samples 4 in from edges")
        elif dtp_id == "STP311":
            self.animation_canvas.create_rectangle(78, 76, 182, 180, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_oval(111, 109, 149, 147, outline="#64748b", width=3)
            self.animation_canvas.create_oval(125, 123, 135, 133, fill="#111827", outline="#111827")
            self.animation_canvas.create_line(130, 122, 130, 55, arrow=tk.LAST, fill="#dc2626", width=3)
            self.animation_canvas.create_text(130, 62, text="Pull")
            self.animation_canvas.create_text(130, 202, text="Nail pull point centered on 6 x 6")
        elif dtp_id == "STP315":
            self.animation_canvas.create_rectangle(30, 56, 230, 184, fill="#f8fafc", outline="#94a3b8", width=2)
            self.animation_canvas.create_rectangle(68, 88, 192, 152, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_line(82, 120, 178, 120, arrow=tk.LAST, fill="#16a34a", width=3)
            self.animation_canvas.create_text(130, 106, text="12 in MD")
            self.animation_canvas.create_line(68, 72, 192, 72, arrow=tk.BOTH, width=2)
            self.animation_canvas.create_text(130, 62, text="24 in")
            self.animation_canvas.create_text(130, 204, text="Center across width, 12 in from ends")
        elif dtp_id == "STP318":
            self.animation_canvas.create_rectangle(42, 54, 218, 184, fill="#f8fafc", outline="#94a3b8", width=2)
            self.animation_canvas.create_rectangle(42, 54, 64, 184, fill="#c7d2fe", outline="#4f46e5", width=2)
            self.animation_canvas.create_text(53, 194, text="formed edge", font=("Segoe UI", 8), angle=0)
            self.animation_canvas.create_rectangle(64, 82, 118, 122, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_rectangle(64, 132, 118, 172, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_line(64, 76, 118, 76, arrow=tk.BOTH, width=2)
            self.animation_canvas.create_text(91, 66, text='4" side')
            self.animation_canvas.create_text(158, 124, text="6 in from board ends", width=90)
        else:
            self.animation_canvas.create_rectangle(60, 76, 200, 170, fill="#dbeafe", outline="#1d4ed8", width=2)
            self.animation_canvas.create_text(130, 200, text=RULES_JSON[dtp_id].get("test_name", "Setup"), width=230)

        if sample_id:
            self.animation_canvas.create_text(130, 215, text=f"Selected: {sample_id}", font=("Segoe UI", 8))

    def export_svg(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Nothing to export", "Generate a layout first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".svg", filetypes=[("SVG files", "*.svg")])
        if not path:
            return
        try:
            SvgExporter.export(path, self.layout_result)
            self.last_export_folder = os.path.dirname(path)
            messagebox.showinfo("Export Complete", f"SVG saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def export_csv(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Nothing to export", "Generate a layout first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            CsvExporter.export(path, self.layout_result)
            self.last_export_folder = os.path.dirname(path)
            messagebox.showinfo("Export Complete", f"CSV saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def export_dxf(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Nothing to export", "Generate a layout first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".dxf", filetypes=[("DXF files", "*.dxf")])
        if not path:
            return
        try:
            DxfExporter.export(path, self.layout_result)
            self.last_export_folder = os.path.dirname(path)
            messagebox.showinfo("Export Complete", f"DXF saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def export_pdf_report(self) -> None:
        if not self.layout_result:
            messagebox.showinfo("Nothing to export", "Generate a layout first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        try:
            PdfReportExporter.export(path, self.layout_result, self._board_validation_issues(self.layout_result))
            self.last_export_folder = os.path.dirname(path)
            messagebox.showinfo("Export Complete", f"PDF report saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def export_job_package(self) -> None:
        if not self.layout_result and not self.saved_sheets:
            messagebox.showinfo("Nothing to export", "Generate or save at least one sheet first.")
            return
        folder = filedialog.askdirectory(title="Choose folder for job package")
        if not folder:
            return
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem = self._safe_file_stem(self.project_var.get() or self.sheet_name_var.get() or "gp_cnc_job")
            package_dir = os.path.join(folder, f"{stem}_{timestamp}")
            os.makedirs(package_dir, exist_ok=True)
            sheets = self.saved_sheets[:] or ([self.layout_result] if self.layout_result else [])
            for index, sheet in enumerate(sheets, start=1):
                name = self._safe_file_stem(sheet.metadata.get("sheet_name", f"sheet{index}")) or f"sheet{index}"
                DxfExporter.export(os.path.join(package_dir, f"{name}.dxf"), sheet)
                CsvExporter.export(os.path.join(package_dir, f"{name}.csv"), sheet)
                SvgExporter.export(os.path.join(package_dir, f"{name}.svg"), sheet)
                PdfReportExporter.export(os.path.join(package_dir, f"{name}_report.pdf"), sheet, self._board_validation_issues(sheet))
            project_path = os.path.join(package_dir, "project.gpcnc.json")
            self._write_project_file(project_path)
            zip_path = f"{package_dir}.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for root_dir, _dirs, files in os.walk(package_dir):
                    for file_name in files:
                        full_path = os.path.join(root_dir, file_name)
                        zf.write(full_path, os.path.relpath(full_path, package_dir))
            self.last_export_folder = package_dir
            messagebox.showinfo("Export Complete", f"Job package saved to:\n{package_dir}\n\nZip:\n{zip_path}")
        except Exception as exc:
            messagebox.showerror("Package Export Error", str(exc))

    def open_last_export_folder(self) -> None:
        folder = self.last_export_folder
        if not folder or not os.path.isdir(folder):
            folder = os.path.dirname(os.path.abspath(__file__))
        os.startfile(folder)

    def save_project_file(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".gpcnc.json", filetypes=[("GP CNC project", "*.gpcnc.json"), ("JSON files", "*.json")])
        if not path:
            return
        try:
            self._write_project_file(path)
            messagebox.showinfo("Project Saved", f"Project saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Project Save Error", str(exc))

    def _write_project_file(self, path: str) -> None:
        data = {
            "version": 1,
            "project": {
                "operator": self.operator_var.get(),
                "project_id": self.project_var.get(),
                "product": self.product_var.get(),
                "sheet_name": self.sheet_name_var.get(),
                "board_width": self.board_w_var.get(),
                "board_height": self.board_h_var.get(),
                "thickness": self.thickness_var.get(),
                "machine_direction": self.md_var.get(),
            },
            "current_layout": self.layout_to_dict(self.layout_result) if self.layout_result else None,
            "saved_sheets": [self.layout_to_dict(sheet) for sheet in self.saved_sheets],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def load_project_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("GP CNC project", "*.gpcnc.json"), ("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            project = data.get("project", {})
            self.operator_var.set(project.get("operator", ""))
            self.project_var.set(project.get("project_id", ""))
            self.product_var.set(project.get("product", ""))
            self.sheet_name_var.set(project.get("sheet_name", "Sheet 1"))
            self.board_w_var.set(project.get("board_width", "48"))
            self.board_h_var.set(project.get("board_height", "96"))
            self.thickness_var.set(project.get("thickness", "0.625"))
            self.md_var.set(project.get("machine_direction", "Horizontal"))
            self.saved_sheets = [self.layout_from_dict(sheet) for sheet in data.get("saved_sheets", [])]
            self.layout_result = self.layout_from_dict(data["current_layout"]) if data.get("current_layout") else None
            self.refresh_saved_sheet_list()
            self.draw_layout()
            self.show_warnings()
            messagebox.showinfo("Project Loaded", f"Project loaded from:\n{path}")
        except Exception as exc:
            messagebox.showerror("Project Load Error", str(exc))


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
