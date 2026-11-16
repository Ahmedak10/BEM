"""
Wind Turbine Blade Design using Blade Element Momentum (BEM) Theory
Reference: "Wind Energy Explained: Theory, Design and Application"
           Manwell, McGowan & Rogers - Chapter 3

Steps implemented (directly following the PDF):
  Step 4   - airfoil loaded from file
  Step 5   - design aerodynamic conditions (Cl_design, alpha_design)
  Steps 6-7 - blade shape: Glauert-optimum chord & twist
  Step 8   - BEM iteration, Method 2 (Eqs. 3.124-3.133)
  Step 9   - power-coefficient integral (Eqs. 3.134-3.136)
  Step 10  - CP-lambda sweep (Section 3.10.2)

Outputs (written to OUTPUT_DIR):
  airfoil_shape.png
  blade_design.png      chord / twist / AoA / induction plots
  cp_lambda.png         CP vs lambda curve
  blade_elements.csv    per-element BEM results
  cp_lambda.csv         CP vs lambda data
"""

import numpy as np
import matplotlib.pyplot as plt
import os
import openpyxl

# ===============================================================================
#  USER INPUTS  - edit these to match your turbine
# ===============================================================================

# -- Airfoil -------------------------------------------------------------------
# Set AIRFOIL_NAME to the airfoil name (without extension).
# The code will automatically look for a file named "<AIRFOIL_NAME>.txt"
# in the same folder, and use its coordinates for NeuralFoil.
# Example: AIRFOIL_NAME = "NACA2412"  ->  reads "NACA2412.txt"

AIRFOIL_NAME = "airfoil"              # airfoil name (file = AIRFOIL_NAME + ".txt")
AIRFOIL_FILE = AIRFOIL_NAME + ".txt"   # resolved automatically -- do not edit

# -- Mode ----------------------------------------------------------------------
# "Design"   -- full design procedure (Steps 6-7): chord and twist are
#               computed from the Glauert optimum / linear fit equations.
#               NUM_ELEMENTS, BLADE_SHAPE_MODE, and all design parameters apply.
#
# "Evaluate" -- evaluate an existing blade geometry read from BLADE_FILE.
#               Eqs. 3.110-3.116 (Steps 6-7) are SKIPPED entirely.
#               Chord and twist come directly from the file.
#               NUM_ELEMENTS is set automatically to match the file row count.
#               BLADE_SHAPE_MODE is ignored.
#               All BEM equations (Steps 8-10) run unchanged.

MODE       = "Design"      # "Design" | "Evaluate"
BLADE_FILE = "blade.xlsx"    # used only when MODE = "Evaluate"
                             # Sheet1 columns: element number | radial distance | chord | pitch angle

# -- Polar source --------------------------------------------------------------
# POLAR_SOURCE controls where aerodynamic polars come from:
#
#   "neuralfoil" -- NeuralFoil generates polars automatically from the airfoil
#                   coordinates in AIRFOIL_FILE.  No polar files needed.
#                   Re range is computed automatically from turbine geometry.
#                   Full 360-degree alpha range is used.
#                   Requires:  pip install neuralfoil
#
#   "files"      -- Load polars from QBlade .plr files (original behaviour).
#                   Set POLAR_DIR and/or POLAR_EXTRA below.

POLAR_SOURCE = "neuralfoil"   # "neuralfoil" | "files"

# NeuralFoil settings (only used when POLAR_SOURCE = "neuralfoil") ------------
#
# NF_MODEL_SIZE controls the accuracy/speed tradeoff:
#   "small"    -- fastest,  least accurate
#   "medium"   -- balanced
#   "large"    -- recommended for most cases
#   "xlarge"   -- most accurate, slowest  (~5-10x slower than "large")
#
# Re range and number of Re points are computed automatically from the turbine
# geometry (hub/tip radius, wind speed, tip-speed ratio).  The range is
# bracketed with a safety margin and sampled on a log scale.

NF_MODEL_SIZE  = "large"   # "small" | "medium" | "large" | "xlarge"
NF_NUM_RE      = 100       # number of Re points (log-spaced, auto range)
NF_RE_MIN      = 50_000      # override auto minimum Re  (e.g. 50_000)  -- None = auto
NF_RE_MAX      = 700_000      # override auto maximum Re  (e.g. 5_000_000) -- None = auto

# Files-based polar settings (only used when POLAR_SOURCE = "files") ----------
#
# Path to a folder containing QBlade .plr polar files, one file per Re.
# ALL .plr files found in the folder are loaded automatically.
# Re is parsed from each file's "REYNOLDS" header line -- no manual entry needed.
# Excel .xlsx files in the folder are ignored (add them to POLAR_EXTRA if needed).
#
# You can also add individual extra files (any format) via POLAR_EXTRA:
#   POLAR_EXTRA = [("my_polar.xlsx", 600000)]  -- Re must be given for xlsx
#   POLAR_EXTRA = []                            -- nothing extra
#
# Files are sorted by Re automatically. At least 1 file is required.
# The more Re values covered, the more accurate the BEM interpolation.

POLAR_DIR   = "Polar"    # folder with .plr files  (set to "" to skip)
POLAR_EXTRA = []         # optional extra files:  ["file.plr"] or [("file.xlsx", Re)]

# -- Design parameters ----------------------------------------------------------
TIP_SPEED_RATIO  = 7               # design tip-speed ratio  lambda
WIND_SPEED       = 8            # m/s  free-stream wind speed  Uinf
HUB_RADIUS       = 0.312           # m    hub radius  r_hub
ROTOR_RADIUS     = 3.0             # m    rotor (tip) radius  R
NUM_BLADES       = 3               # B    number of blades
NUM_ELEMENTS     = 30              # N    number of blade elements
ALPHA_DESIGN     = 4.0             # deg  design angle of attack  alpha_design
AIR_DENSITY      = 1.225           # kg/m3
DYN_VISCOSITY    = 1.49e-5         # Pa*s  dynamic viscosity of air  mu
CL_DESIGN        = 1.2730          # lift coefficient at ALPHA_DESIGN

RE_ASSUMED       =  400_000        # Re used for initial chord sizing (Eq. 3.125 initial guess)
                                   # Set to match your lowest/design polar Re.
                                   # Also used to select the polar for auto design point.

# -- Design point mode ---------------------------------------------------------
# "manual"   -- use ALPHA_DESIGN and CL_DESIGN as set above (original behaviour).
#
# "auto"     -- automatically find the best Cl/Cd point from the polar at
#               RE_ASSUMED and use that alpha and Cl as the design point.
#               ALPHA_DESIGN and CL_DESIGN above are ignored in this mode.
#               Only affects Design mode; in Evaluate mode the design point
#               has no effect on BEM results regardless.

DESIGN_POINT_MODE = "auto"   # "manual" | "auto"

# -- Step 7: Linear blade shape  (Eqs. 3.115 - 3.116) --------------------------
#
# The textbook says to use the optimum blade as a guide and then choose a
# LINEAR approximation for ease of fabrication:
#
#   chord(r)  = a1 * r  +  b1          (Eq. 3.115)
#   twist(r)  = a2 * (R - r)           (Eq. 3.116)
#
# Three modes are available -- set BLADE_SHAPE_MODE to one of:
#
#   "optimum"   -- use the full Glauert-optimum chord and twist at every
#                  element (Steps 6 only, no linearisation).  Default.
#
#   "auto_fit"  -- automatically least-squares fit a1, b1, a2 to the
#                  optimum distributions computed at the design lambda,
#                  then use those linear functions (Step 7 as written).
#
#   "manual"    -- use the a1, b1, a2 values you supply below.
#                  Set CHORD_A1, CHORD_B1 for chord; TWIST_A2 for twist.
#                  Alternatively set C_ROOT_M + C_TIP_M and the code
#                  converts them to a1/b1 automatically.

BLADE_SHAPE_MODE = "auto_fit"   # "optimum" | "auto_fit" | "manual"

# Manual coefficients (only used when BLADE_SHAPE_MODE = "manual")
# chord(r) = CHORD_A1 * r + CHORD_B1   [r and chord both in metres]
CHORD_A1 = None   # m/m  slope  (None -> derived from C_ROOT_M / C_TIP_M)
CHORD_B1 = None   # m    intercept
# Shorthand: set root and tip chord instead; code computes a1, b1 for you.
C_ROOT_M = None   # m  chord at r = HUB_RADIUS  (None -> use CHORD_A1/B1)
C_TIP_M  = None   # m  chord at r = ROTOR_RADIUS

# twist(r) = TWIST_A2 * (R - r)          [result in degrees]
TWIST_A2 = None   # deg/m  (None -> derived from root twist / tip = 0)

# -- Collective blade pitch -----------------------------------------------------
PITCH_0_DEG = 0.0   # deg  positive = nose-up

# -- BEM solver controls --------------------------------------------------------
MAX_ITER  = 1000
TOLERANCE = 1e-8
RELAX     = 0.35    # under-relaxation (0 = no update, 1 = full Newton step)

# Heavy-loading correction for axial induction (applied when CTr > 0.96)
# "glauert" -- Original Glauert empirical correction (Eq. 3.132 in textbook).
#              Derived at rotor level; simple and widely used in textbooks.
# "buhl"    -- Buhl (2005) modified Glauert correction, used in NREL AeroDyn
#              and OpenFAST. Derived at element level with F embedded directly;
#              more accurate near the tip and for highly loaded elements.
INDUCTION_CORRECTION = "buhl"   # "glauert" | "buhl"

# Prandtl hub loss correction
# True  -- combine tip loss AND hub loss:  F = F_tip * F_hub
#          f_hub = (B/2) * (r - r_hub) / (r_hub * sin(phi))
#          Reduces induction near the root; more accurate for large hub ratios.
# False -- tip loss only (standard textbook BEM, Eq. 3.120).
HUB_LOSS = True   # True | False

# Snel (1994) 3D rotational augmentation correction
# True  -- apply stall-delay correction to Cl (and optionally Cd) at each
#          element before the BEM iteration uses them.
#          Blends the 2D polar Cl toward the attached-flow (linear) value
#          using the local chord-to-radius ratio:
#            f       = min(1,  3 * (c/r)^2)          rotational factor
#            Cl_att  = 2*pi * alpha_rad               attached-flow Cl
#            Cl_3d   = Cl_2d + f * (Cl_att - Cl_2d)  corrected Cl
#            Cd_3d   = Cd_2d * (1 - f)               corrected Cd
#          Strong near root (large c/r), negligible near tip (small c/r).
# False -- use the 2D polar directly (no rotational correction).
SNEL_3D_CORRECTION = True   # True | False

# -- CP-lambda sweep -----------------------------------------------------------------
LAMBDA_MIN   = 2
LAMBDA_MAX   = 12
LAMBDA_STEPS = 11

OUTPUT_DIR = "."    # folder for saved plots and CSVs

# -- Power curve (variable speed, constant lambda = TIP_SPEED_RATIO) -----------
POWER_CURVE_UMIN  =  3.0   # m/s  cut-in  wind speed
POWER_CURVE_UMAX  = 15   # m/s  cut-out wind speed
POWER_CURVE_STEPS =  13    # number of wind speed points
# ===============================================================================


# --- Blade file loader (Evaluate mode) ----------------------------------------
#
# _BLADE_R, _BLADE_CHORD, _BLADE_TWIST are populated by load_blade_file()
# when MODE = "Evaluate".  They store the exact stations from Sheet1.

_BLADE_R     = np.array([])   # radial positions  [m]
_BLADE_CHORD = np.array([])   # chord lengths      [m]
_BLADE_TWIST = np.array([])   # pitch / twist      [deg]


def load_blade_file(filepath: str) -> int:
    """
    Read blade geometry from Sheet1 of an Excel file.

    Expected columns (row 1 = header, ignored):
        Col 1: element number   (ignored)
        Col 2: radial distance  [m]
        Col 3: chord length     [m]
        Col 4: Pitch Angle      [deg]

    Returns the number of stations loaded.
    Rows with any missing value in columns 2-4 are skipped.
    Stations are sorted by radius.
    """
    global _BLADE_R, _BLADE_CHORD, _BLADE_TWIST

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)

    if "Sheet1" not in wb.sheetnames:
        raise ValueError(f"'{filepath}' has no 'Sheet1'. "
                         f"Available sheets: {wb.sheetnames}")

    ws = wb["Sheet1"]
    rs, chords, twists = [], [], []
    first_row = True

    for row in ws.iter_rows(values_only=True):
        if first_row:
            first_row = False
            continue                         # skip header
        if len(row) < 4:
            continue
        _, r, c, t = row[0], row[1], row[2], row[3]
        if any(v is None for v in (r, c, t)):
            continue
        try:
            rs.append(float(r))
            chords.append(float(c))
            twists.append(float(t))
        except (TypeError, ValueError):
            continue

    if not rs:
        raise ValueError(f"No valid data rows found in Sheet1 of '{filepath}'.")

    order        = np.argsort(rs)
    _BLADE_R     = np.array(rs)[order]
    _BLADE_CHORD = np.array(chords)[order]
    _BLADE_TWIST = np.array(twists)[order]

    return len(_BLADE_R)


# --- Multi-Re aerodynamic polar -----------------------------------------------
#
# Storage: a list of polars sorted by Re, each a dict:
#   {"re": float, "alpha": np.array, "cl": np.array, "cd": np.array}
#
# At each BEM element the code calls cl_cd_at_re(alpha_deg, Re_local) which:
#   1. Looks up Cl and Cd in each polar by alpha interpolation
#   2. Interpolates those values across Re
#   If only one polar is loaded, Re interpolation is skipped (single-Re mode).

_POLARS = []    # list of dicts, sorted ascending by "re"


# -- Parsers -------------------------------------------------------------------

def _parse_plr(filepath: str) -> dict:
    """
    Parse a QBlade Multi-RE Polar (.plr) file.

    Extracts:
      - Re from the line starting with "REYNOLDS"
      - AOA, CL, CD columns from the data block (after the "AOA" header line)
      - Also reads CL_ATT if present (used for Snel correction's attached Cl)

    Returns dict with keys: re, alpha, cl, cd, cl_att (optional)
    """
    re_val    = None
    alphas    = []
    cls       = []
    cds       = []
    cl_atts   = []
    in_data   = False
    has_cl_att = False

    with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue

            # parse Re from header
            if line.upper().startswith("REYNOLDS"):
                parts = line.split()
                for p in parts:
                    try:
                        v = float(p)
                        if v > 1000:          # the Re value itself
                            re_val = v
                            break
                        elif v > 0:           # scientific form like 1.0E+05
                            re_val = v
                    except ValueError:
                        continue

            # detect column header line
            if line.upper().startswith("AOA"):
                cols = line.split()
                has_cl_att = "CL_ATT" in [c.upper() for c in cols]
                in_data = True
                continue

            if not in_data:
                continue

            # parse data row
            parts = line.split()
            if len(parts) < 3:
                continue
            try:
                aoa = float(parts[0])
                cl  = float(parts[1])
                cd  = float(parts[2])
                alphas.append(aoa)
                cls.append(cl)
                cds.append(cd)
                if has_cl_att and len(parts) >= 5:
                    cl_atts.append(float(parts[4]))
                else:
                    cl_atts.append(2.0 * np.pi * np.radians(aoa))
            except (ValueError, IndexError):
                continue

    if not alphas:
        raise ValueError(f"No data rows found in '{filepath}'.")
    if re_val is None:
        raise ValueError(f"Could not parse REYNOLDS value from '{filepath}'.")

    order = np.argsort(alphas)
    return dict(
        re     = re_val,
        alpha  = np.array(alphas)[order],
        cl     = np.array(cls)[order],
        cd     = np.array(cds)[order],
        cl_att = np.array(cl_atts)[order],
    )


def _parse_xlsx(filepath: str, re_val: float) -> dict:
    """
    Parse an Excel polar file (same format as the original alpha_data.xlsx):
        Column A: alpha [deg],  B: Cl,  C: Cd
    Re must be supplied as re_val (cannot be inferred from the file).
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    alphas, cls, cds = [], [], []
    first_row = True
    for row in ws.iter_rows(values_only=True):
        if first_row:
            first_row = False
            continue
        if None in (row[0], row[1], row[2]):
            continue
        try:
            alphas.append(float(row[0]))
            cls.append(float(row[1]))
            cds.append(float(row[2]))
        except (TypeError, ValueError):
            continue

    if not alphas:
        raise ValueError(f"No valid data found in '{filepath}'.")

    order = np.argsort(alphas)
    alpha_arr = np.array(alphas)[order]
    return dict(
        re     = float(re_val),
        alpha  = alpha_arr,
        cl     = np.array(cls)[order],
        cd     = np.array(cds)[order],
        cl_att = 2.0 * np.pi * np.radians(alpha_arr),
    )


# -- Loader --------------------------------------------------------------------

def _load_polars_neuralfoil() -> None:
    """
    Generate multi-Re polars using NeuralFoil from the airfoil coordinates
    in AIRFOIL_FILE.

    Re range is computed automatically from the turbine geometry:
      - minimum Re  : hub radius, low chord estimate, hub relative velocity
      - maximum Re  : tip radius, high chord estimate, tip relative velocity
    NF_NUM_RE log-spaced Re points are evaluated across this range.
    Full 360-degree alpha sweep is used.
    """
    global _POLARS
    import neuralfoil as nf

    if not os.path.exists(AIRFOIL_FILE):
        raise FileNotFoundError(
            f"Airfoil file '{AIRFOIL_FILE}' not found. "
            f"Make sure '{AIRFOIL_NAME}.txt' exists in the working directory.")

    pts = load_airfoil(AIRFOIL_FILE)
    print(f"  Airfoil coordinates loaded from '{AIRFOIL_FILE}'  "
          f"({len(pts)} points)")

    # --- auto Re range from turbine geometry ----------------------------------
    nu    = DYN_VISCOSITY / AIR_DENSITY
    omega = TIP_SPEED_RATIO * WIND_SPEED / ROTOR_RADIUS

    # conservative chord estimates before the blade is solved
    c_min = 0.05 * HUB_RADIUS       # ~5 % of hub radius
    c_max = 0.20 * ROTOR_RADIUS     # ~20 % of rotor radius

    V_hub = np.sqrt(WIND_SPEED**2 + (omega * HUB_RADIUS)**2)
    V_tip = np.sqrt(WIND_SPEED**2 + (omega * ROTOR_RADIUS)**2)

    Re_min = NF_RE_MIN if NF_RE_MIN is not None else max(V_hub * c_min / nu, 1e4)
    Re_max = NF_RE_MAX if NF_RE_MAX is not None else min(V_tip * c_max / nu, 2e7)

    re_list = np.geomspace(Re_min, Re_max, NF_NUM_RE).tolist()
    print(f"  Auto Re range : {Re_min:,.0f} -> {Re_max:,.0f}  "
          f"({NF_NUM_RE} points, log-spaced)")
    print(f"  Model size    : {NF_MODEL_SIZE}")
    print(f"  Alpha range   : -180 to +180 deg  (361 points, full 360°)")

    # --- generate polars ------------------------------------------------------
    alpha_range = np.linspace(-180.0, 180.0, 361)

    for re in re_list:
        out = nf.get_aero_from_coordinates(
            coordinates  = pts,
            alpha        = alpha_range,
            Re           = re,
            model_size   = NF_MODEL_SIZE,
        )
        cl_arr = np.array(out["CL"], dtype=float)
        cd_arr = np.array(out["CD"], dtype=float)

        # Clamp Cd to physically sensible lower bound
        cd_arr = np.maximum(cd_arr, 0.0)

        _POLARS.append(dict(
            re     = float(re),
            alpha  = alpha_range.copy(),
            cl     = cl_arr,
            cd     = cd_arr,
            cl_att = 2.0 * np.pi * np.radians(alpha_range),  # thin-airfoil Cl_att
        ))
        print(f"    Re={re:>10,.0f}  "
              f"Cl_max={cl_arr.max():.3f}  "
              f"Cd_min={cd_arr.min():.5f}")

    _POLARS.sort(key=lambda p: p["re"])
    re_vals = [p["re"] for p in _POLARS]
    print(f"  NeuralFoil: {len(_POLARS)} polar(s) ready  |  "
          f"Re range: {re_vals[0]:,.0f} -- {re_vals[-1]:,.0f}")


def _load_polars_files() -> None:
    """
    Scan POLAR_DIR for all .plr files, load them, then append any entries
    in POLAR_EXTRA.  Results are stored in _POLARS sorted ascending by Re.

    .plr files : Re is parsed automatically from the REYNOLDS header line.
    .xlsx files: Re must be supplied explicitly in POLAR_EXTRA as (path, Re).
    """
    global _POLARS

    # --- collect file list ----------------------------------------------------
    file_list = []   # list of (filepath, re_override_or_None)

    if POLAR_DIR:
        if not os.path.isdir(POLAR_DIR):
            raise FileNotFoundError(
                f"POLAR_DIR '{POLAR_DIR}' does not exist. "
                f"Create the folder and place your .plr files inside it.")
        plr_files = sorted(
            f for f in os.listdir(POLAR_DIR)
            if f.lower().endswith(".plr"))
        if not plr_files:
            raise FileNotFoundError(
                f"No .plr files found in POLAR_DIR '{POLAR_DIR}'.")
        for fname in plr_files:
            file_list.append((os.path.join(POLAR_DIR, fname), None))

    for entry in POLAR_EXTRA:
        if isinstance(entry, (list, tuple)):
            file_list.append((entry[0], entry[1]))
        else:
            file_list.append((entry, None))

    if not file_list:
        raise ValueError("No polar files specified. "
                         "Set POLAR_DIR or add entries to POLAR_EXTRA.")

    # --- load each file -------------------------------------------------------
    print(f"  Scanning '{POLAR_DIR}' ... found {len(file_list)} file(s)")
    for filepath, re_override in file_list:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".plr":
            polar = _parse_plr(filepath)
        elif ext in (".xlsx", ".xls"):
            if re_override is None:
                raise ValueError(
                    f"Excel polar '{filepath}' needs an explicit Re value. "
                    f"Use ('{filepath}', re_value) in POLAR_EXTRA.")
            polar = _parse_xlsx(filepath, re_override)
        else:
            raise ValueError(
                f"Unsupported format '{ext}' for '{filepath}'. "
                f"Use .plr or .xlsx.")

        _POLARS.append(polar)
        print(f"    {os.path.basename(filepath):<60}  "
              f"Re={polar['re']:>10,.0f}  "
              f"({len(polar['alpha'])} pts)")

    # --- sort by Re -----------------------------------------------------------
    _POLARS.sort(key=lambda p: p["re"])
    re_list = [p["re"] for p in _POLARS]
    print(f"  Loaded {len(_POLARS)} polar(s)  |  "
          f"Re range: {re_list[0]:,.0f} -- {re_list[-1]:,.0f}")


def load_polars() -> None:
    """
    Dispatch to NeuralFoil or file-based polar loading depending on POLAR_SOURCE.
    """
    global _POLARS
    _POLARS = []

    if POLAR_SOURCE == "neuralfoil":
        _load_polars_neuralfoil()
    elif POLAR_SOURCE == "files":
        _load_polars_files()
    else:
        raise ValueError(
            f"Unknown POLAR_SOURCE '{POLAR_SOURCE}'. "
            f"Use 'neuralfoil' or 'files'.")


# -- Lookup --------------------------------------------------------------------

def cl_cd_at_re(alpha_deg: float, re: float):
    """
    Return (Cl, Cd, Cl_att) interpolated at the given alpha and Re.

    With one polar  -> simple alpha interpolation (single-Re mode).
    With N polars   -> alpha interpolation in each polar, then linear Re
                       interpolation between the two bracketing polars.
                       Re outside the range is clamped to the nearest polar.
    """
    if len(_POLARS) == 1:
        p = _POLARS[0]
        cl     = float(np.interp(alpha_deg, p["alpha"], p["cl"]))
        cd     = float(np.interp(alpha_deg, p["alpha"], p["cd"]))
        cl_att = float(np.interp(alpha_deg, p["alpha"], p["cl_att"]))
        return cl, cd, cl_att

    # find bracketing polars
    re_vals = [p["re"] for p in _POLARS]

    if re <= re_vals[0]:
        p = _POLARS[0]
        return (float(np.interp(alpha_deg, p["alpha"], p["cl"])),
                float(np.interp(alpha_deg, p["alpha"], p["cd"])),
                float(np.interp(alpha_deg, p["alpha"], p["cl_att"])))

    if re >= re_vals[-1]:
        p = _POLARS[-1]
        return (float(np.interp(alpha_deg, p["alpha"], p["cl"])),
                float(np.interp(alpha_deg, p["alpha"], p["cd"])),
                float(np.interp(alpha_deg, p["alpha"], p["cl_att"])))

    # interpolate between two bracketing polars
    idx_hi = next(i for i, rv in enumerate(re_vals) if rv >= re)
    idx_lo = idx_hi - 1
    p_lo, p_hi = _POLARS[idx_lo], _POLARS[idx_hi]
    t = (re - p_lo["re"]) / (p_hi["re"] - p_lo["re"])   # 0..1

    def _interp_pair(key):
        lo = float(np.interp(alpha_deg, p_lo["alpha"], p_lo[key]))
        hi = float(np.interp(alpha_deg, p_hi["alpha"], p_hi[key]))
        return lo + t * (hi - lo)

    return _interp_pair("cl"), _interp_pair("cd"), _interp_pair("cl_att")


# -- Backwards-compatible wrappers (used outside BEM where Re is unknown) ------

def cl_from_alpha(alpha_deg: float) -> float:
    """Cl at the design Re (first/only polar).  Used for initial chord sizing."""
    p = _POLARS[0] if _POLARS else None
    if p is None:
        return 2.0 * np.pi * np.radians(alpha_deg)
    return float(np.interp(alpha_deg, p["alpha"], p["cl"]))


def cd_from_alpha(alpha_deg: float) -> float:
    """Cd at the design Re (first/only polar).  Used for initial chord sizing."""
    p = _POLARS[0] if _POLARS else None
    if p is None:
        return 0.01
    return float(np.interp(alpha_deg, p["alpha"], p["cd"]))



# --- Snel (1994) 3D rotational augmentation correction -----------------------

def snel_3d_cl_cd(alpha_deg: float, chord: float, r: float, re: float):
    """
    Apply Snel et al. (1994) 3D stall-delay correction to Cl and Cd.

    Parameters
    ----------
    alpha_deg : local angle of attack  [deg]
    chord     : local chord            [m]
    r         : element radius         [m]
    re        : local Reynolds number  (used for multi-Re polar lookup)

    Returns
    -------
    Cl_3d, Cd_3d : corrected coefficients
    """
    Cl_2d, Cd_2d, Cl_att_table = cl_cd_at_re(alpha_deg, re)

    if not SNEL_3D_CORRECTION or r < 1e-6:
        return Cl_2d, Cd_2d

    # rotational factor  (capped at 1 to avoid over-correction)
    f = min(1.0, 3.0 * (chord / r) ** 2)

    # attached-flow Cl: use CL_ATT from polar if available, else thin-aerofoil
    Cl_att = Cl_att_table if abs(Cl_att_table) > 1e-9 else \
             2.0 * np.pi * np.radians(alpha_deg)

    Cl_3d = Cl_2d + f * (Cl_att - Cl_2d)
    Cd_3d = Cd_2d * (1.0 - f)
    Cd_3d = max(Cd_3d, 0.0)

    return Cl_3d, Cd_3d




def _prandtl_factor(f: float) -> float:
    """F = (2/pi) * arccos(exp(-f)),  clamped to [1e-6, 1]."""
    return max((2.0 / np.pi) * np.arccos(np.clip(np.exp(-f), -1.0, 1.0)), 1e-6)


def tip_loss(B: int, r: float, R: float, phi_rad: float) -> float:
    """
    Combined Prandtl loss factor  F = F_tip * F_hub  (if HUB_LOSS=True)
    or tip loss only              F = F_tip           (if HUB_LOSS=False).

    Tip loss  (Eq. 3.120 / 3.128):
        f_tip = (B/2) * (R - r)   / (r      * |sin(phi)|)

    Hub loss (same form, hub end):
        f_hub = (B/2) * (r - r_hub) / (r_hub * |sin(phi)|)
    """
    sin_phi = abs(np.sin(phi_rad))
    if sin_phi < 1e-9:
        return 1.0

    f_tip = (B / 2.0) * (R - r)             / (r               * sin_phi)
    F_tip = _prandtl_factor(f_tip)

    if HUB_LOSS:
        f_hub = (B / 2.0) * (r - HUB_RADIUS) / (HUB_RADIUS * sin_phi)
        F_hub = _prandtl_factor(f_hub)
        return F_tip * F_hub

    return F_tip


# --- Optimum blade geometry ---------------------------------------------------

def opt_chord(r: float, R: float, lam: float, B: int, Cl_d: float) -> float:
    """
    Glauert optimum chord (Eq. 3.112 rearranged):
        c = (8pi r * sin phi) / (3 B Cl_d * lambda_r)
    where lambda_r = lambda * r/R  and  phi = (2/3) arctan(1/lambda_r)  [Eq. 3.111]
    This form stays bounded at all radii.
    """
    lr  = lam * r / R
    phi = (2.0 / 3.0) * np.arctan(1.0 / max(lr, 1e-9))
    return (8.0 * np.pi * r * np.sin(phi)) / (3.0 * B * Cl_d * lr)


def opt_twist(r: float, R: float, lam: float, alpha_d: float) -> float:
    """
    Optimum twist  theta_T = phi - alpha_design  (Eqs. 3.111, 3.113-3.114).
    """
    lr  = lam * r / R
    phi_deg = np.degrees((2.0 / 3.0) * np.arctan(1.0 / max(lr, 1e-9)))
    return phi_deg - alpha_d


# --- Step 7: Linear blade shape coefficients (Eqs. 3.115 - 3.116) ------------
#
# These are module-level variables set once by resolve_linear_coefficients().
# They are used inside analyse_blade() when BLADE_SHAPE_MODE != "optimum".

_A1 = None   # chord slope    [m/m]
_B1 = None   # chord intercept [m]
_A2 = None   # twist slope     [deg/m]


def resolve_linear_coefficients(lam: float) -> None:
    """
    Determine a1, b1, a2 and store them in _A1, _B1, _A2.

    "auto_fit" mode
        Compute the optimum chord and twist at NUM_ELEMENTS radii for the
        given design lambda, then fit:
            chord  = a1*r + b1   by least squares (np.polyfit degree 1)
            twist  = a2*(R - r)  by least squares (force through twist=0 at r=R)

    "manual" mode
        Use CHORD_A1/CHORD_B1/TWIST_A2 if provided.
        If C_ROOT_M and C_TIP_M are given instead, derive a1/b1 from them.
        If TWIST_A2 is None, derive a2 from opt_twist at hub (tip twist = 0).
    """
    global _A1, _B1, _A2

    r_bounds = np.linspace(HUB_RADIUS, ROTOR_RADIUS, NUM_ELEMENTS + 1)
    r_mids   = 0.5 * (r_bounds[:-1] + r_bounds[1:])

    if BLADE_SHAPE_MODE == "auto_fit":
        # --- chord: Eq. 3.115  c = a1*r + b1 ---
        c_opt = np.array([opt_chord(r, ROTOR_RADIUS, lam,
                                    NUM_BLADES, CL_DESIGN) for r in r_mids])
        a1, b1 = np.polyfit(r_mids, c_opt, 1)   # least-squares linear fit

        # --- twist: Eq. 3.116  theta = a2*(R - r) ---
        # Rearrange: a2 = theta / (R - r); fit through origin in (R-r) space.
        t_opt  = np.array([opt_twist(r, ROTOR_RADIUS, lam,
                                     ALPHA_DESIGN) for r in r_mids])
        Rr     = ROTOR_RADIUS - r_mids           # (R - r) values
        # force intercept = 0: a2 = sum(Rr * t_opt) / sum(Rr^2)
        a2     = np.dot(Rr, t_opt) / np.dot(Rr, Rr)

        _A1, _B1, _A2 = a1, b1, a2

    else:  # "manual"
        # chord coefficients
        if CHORD_A1 is not None and CHORD_B1 is not None:
            _A1, _B1 = CHORD_A1, CHORD_B1
        elif C_ROOT_M is not None and C_TIP_M is not None:
            # two-point solve:  c_root = a1*r_hub + b1,  c_tip = a1*R + b1
            _A1 = (C_TIP_M - C_ROOT_M) / (ROTOR_RADIUS - HUB_RADIUS)
            _B1 = C_ROOT_M - _A1 * HUB_RADIUS
        else:
            raise ValueError(
                "BLADE_SHAPE_MODE='manual': supply either CHORD_A1+CHORD_B1 "
                "or C_ROOT_M+C_TIP_M.")

        # twist coefficient
        if TWIST_A2 is not None:
            _A2 = TWIST_A2
        else:
            # derive from opt twist at hub (tip twist forced to 0 per Eq. 3.116)
            twist_hub = opt_twist(HUB_RADIUS, ROTOR_RADIUS, lam, ALPHA_DESIGN)
            _A2 = twist_hub / (ROTOR_RADIUS - HUB_RADIUS)


def linear_chord(r: float) -> float:
    """Eq. 3.115:  c_i = a1 * r_i + b1"""
    return _A1 * r + _B1


def linear_twist(r: float) -> float:
    """Eq. 3.116:  theta_T,i = a2 * (R - r_i)"""
    return _A2 * (ROTOR_RADIUS - r)



def bem_element(r, R, B, chord, twist_deg, pitch0_deg, lam, U):
    """
    Solve BEM equations for one blade element.

    Parameters
    ----------
    r          radius of element mid-point  [m]
    R          rotor tip radius             [m]
    B          number of blades
    chord      local chord                  [m]
    twist_deg  geometric twist (nose-up +)  [deg]
    pitch0_deg collective blade pitch       [deg]
    lam        tip-speed ratio
    U          wind speed                   [m/s]

    Returns dict: a, a_prime, phi_deg, alpha_deg, Cl, Cd, F, CTr, converged
    """
    lr      = lam * r / R
    yp      = twist_deg + pitch0_deg          # total pitch at this section
    sig_p   = B * chord / (2.0 * np.pi * r)  # local solidity  (Eq. 3.118)
    nu_air  = DYN_VISCOSITY / AIR_DENSITY

    # estimate initial Re from rotor speed (refined each iteration)
    V_tip   = lam * U
    Re_est  = max(V_tip * chord / nu_air, 1.0)

    # -- Initial guesses (Eqs. 3.124-3.126) ----------------------------------
    phi0   = (2.0 / 3.0) * np.arctan(1.0 / max(lr, 1e-9))   # Eq. 3.124
    Cl_g, _ = snel_3d_cl_cd(np.degrees(phi0) - yp, chord, r, Re_est)
    a      = 1.0 / (1.0 + 4.0 * np.sin(phi0)**2 /           # Eq. 3.125
                    max(sig_p * Cl_g * np.cos(phi0), 1e-9))
    a      = np.clip(a, 0.0, 0.5)
    a_p_d  = 4.0 * a - 1.0
    a_p    = max((1.0 - 3.0 * a) / a_p_d, 0.0) \
             if abs(a_p_d) > 1e-9 else 0.01               # Eq. 3.126

    phi = phi0
    Cl  = cl_from_alpha(ALPHA_DESIGN)
    Cd  = cd_from_alpha(ALPHA_DESIGN)
    F   = 1.0
    CTr = 0.5
    converged = False

    for _ in range(MAX_ITER):
        den = (1.0 + a_p) * lr
        phi = np.arctan((1.0 - a) / max(den, 1e-9))        # Eq. 3.127
        F   = tip_loss(B, r, R, phi)                        # Eq. 3.128

        alpha_deg = np.degrees(phi) - yp                    # Eq. 3.129

        # update local Re from current induction estimates
        V_rel  = np.sqrt((U * (1.0 - a))**2 +
                         (lam * U / R * r * (1.0 + a_p))**2)
        Re_est = max(V_rel * chord / nu_air, 1.0)

        Cl, Cd = snel_3d_cl_cd(alpha_deg, chord, r, Re_est)  # polar + Snel 3D

        # local thrust coefficient  (Eq. 3.130)
        CTr = (sig_p * (1.0 - a)**2 *
               (Cl * np.cos(phi) + Cd * np.sin(phi)) /
               max(np.sin(phi)**2, 1e-12))

        # update axial induction  (Eqs. 3.131-3.132 / Buhl 2005)
        if CTr < 0.96:
            # Standard momentum equation  (Eq. 3.131) -- both corrections use this
            a_new = 1.0 / (1.0 + 4.0 * F * np.sin(phi)**2 /
                            max(sig_p * Cl * np.cos(phi), 1e-9))
        else:
            if INDUCTION_CORRECTION == "buhl":
                # ----------------------------------------------------------
                # Buhl (2005) modified Glauert correction
                # Reference: Buhl M.L. Jr., "A New Empirical Relationship
                #   between Thrust Coefficient and Induction Factor for the
                #   Turbulent Windmill State", NREL/TP-500-36834, 2005.
                # Used in NREL AeroDyn / OpenFAST.
                #
                #   a = [18F - 20 - 3*sqrt(CTr*(50 - 36F) + 12F*(3F-4))]
                #       / (36F - 50)
                # ----------------------------------------------------------
                F_s   = max(F, 1e-6)
                inner = CTr * (50.0 - 36.0 * F_s) + 12.0 * F_s * (3.0 * F_s - 4.0)
                disc  = max(0.0, inner)
                a_new = (18.0 * F_s - 20.0 - 3.0 * np.sqrt(disc)) / (36.0 * F_s - 50.0)
            else:
                # ----------------------------------------------------------
                # Original Glauert empirical correction  (Eq. 3.132)
                # ----------------------------------------------------------
                disc  = max(0.0, 0.0203 - 0.6427 * (0.889 - CTr))
                a_new = (1.0 / max(F, 1e-9)) * (0.143 + np.sqrt(disc))

        a_new = np.clip(a_new, 0.0, 0.95)

        # update angular induction  (Eq. 3.133)
        denom2 = 4.0 * F * np.cos(phi) / max(sig_p * Cl, 1e-9) - 1.0
        a_p_new = max(1.0 / denom2, 0.0) if denom2 > 1e-9 else 0.0

        if abs(a_new - a) < TOLERANCE and abs(a_p_new - a_p) < TOLERANCE:
            a, a_p = a_new, a_p_new
            converged = True
            break

        a   += RELAX * (a_new   - a)
        a_p += RELAX * (a_p_new - a_p)

    # compute final Snel factor for reporting
    f_snel = min(1.0, 3.0 * (chord / max(r, 1e-6)) ** 2) if SNEL_3D_CORRECTION else 0.0

    # local relative velocity and Reynolds number
    nu_air = DYN_VISCOSITY / AIR_DENSITY
    V_rel  = np.sqrt((U * (1.0 - a))**2 +
                     (lam * U / R * r * (1.0 + a_p))**2)
    Re_local = V_rel * chord / nu_air

    return dict(a=a, a_prime=a_p, phi_deg=np.degrees(phi),
                alpha_deg=np.degrees(phi) - yp,
                Cl=Cl, Cd=Cd, F=F, CTr=CTr,
                f_snel=f_snel, V_rel=V_rel, Re_local=Re_local,
                converged=converged)


# --- Rotor CP (Eqs. 3.134-3.136) ---------------------------------------------

def rotor_cp(elements: list, lam: float) -> float:
    """
    Power coefficient integral (Eqs. 3.134-3.136).

    Original Eq. 3.135 assumes uniform element spacing:
        delta_lr = lam / N   (same for every element)

    This implementation uses the ACTUAL width of each element so that
    non-uniform station distributions (e.g. from a blade file) give
    the correct result.  For evenly spaced elements both methods are
    identical.

    Each element's delta_lr is computed as:
        delta_lr_i = lam * delta_r_i / R
    where delta_r_i is the radial width of element i, derived from the
    midpoint radii:
        delta_r_i = (r_{i+1} - r_{i-1}) / 2   (central difference)
        delta_r_0 = r_1 - r_0                  (forward  at root)
        delta_r_N = r_N - r_{N-1}              (backward at tip)
    """
    N     = len(elements)
    r_arr = np.array([e["r"] for e in elements])

    # actual radial width of each element (central difference, edge cases fwd/bwd)
    delta_r       = np.empty(N)
    delta_r[0]    = r_arr[1]    - r_arr[0]              # forward  difference at root
    delta_r[-1]   = r_arr[-1]   - r_arr[-2]             # backward difference at tip
    delta_r[1:-1] = (r_arr[2:]  - r_arr[:-2]) / 2.0    # central  difference mid-span

    total = 0.0
    for i, e in enumerate(elements):
        phi = np.radians(e["phi_deg"])
        lr  = e["lr"]
        F   = e["F"]
        Cl  = e["Cl"]
        Cd  = e["Cd"]
        if abs(Cl) < 1e-9:
            continue

        # per-element delta_lr using actual width (Eq. 3.135 generalised)
        delta_lr_i = lam * delta_r[i] / ROTOR_RADIUS

        # Eq. 3.136 integrand × per-element width
        term = (delta_lr_i
                * F * np.sin(phi)**2
                * (np.cos(phi) - lr * np.sin(phi))
                * (np.sin(phi) + lr * np.cos(phi))
                * (1.0 - (Cd / Cl) * (np.cos(phi) / max(np.sin(phi), 1e-9)))
                * lr**2)
        total += term

    return 8.0 / lam**2 * total


# --- Full blade analysis at a given lambda ----------------------------------------

def analyse_blade(lam: float = TIP_SPEED_RATIO,
                  lam_geometry: float = None,
                  wind_speed_override: float = None):
    """
    Run BEM for all elements at tip-speed-ratio lam.

    Parameters
    ----------
    lam                  : tip-speed ratio used for the BEM aerodynamic
                           calculation (induction factors, inflow angle, CP).
    lam_geometry         : tip-speed ratio used to compute chord and twist in
                           Design mode.  If None, defaults to lam.
                           For the CP-lambda sweep set to TIP_SPEED_RATIO so
                           blade geometry is fixed at the design point.
                           Ignored in Evaluate mode.
    wind_speed_override  : if supplied, overrides WIND_SPEED for this call
                           (used by the power curve to vary U while keeping lam
                           and blade geometry fixed).

    MODE = "Design"
        Blade shape from Glauert optimum or linear fit (Steps 6-7).
        Eqs. 3.110-3.116 applied at lam_geometry.
        Element radii evenly spaced between HUB_RADIUS and ROTOR_RADIUS.

    MODE = "Evaluate"
        Blade shape read directly from BLADE_FILE.
        Eqs. 3.110-3.116 SKIPPED.
        Element radii = exact file stations.
    """
    U        = wind_speed_override if wind_speed_override is not None else WIND_SPEED
    lam_geom = lam_geometry if lam_geometry is not None else lam

    if MODE == "Evaluate":
        r_mids = _BLADE_R.copy()
    else:
        r_bounds = np.linspace(HUB_RADIUS, ROTOR_RADIUS, NUM_ELEMENTS + 1)
        r_mids   = 0.5 * (r_bounds[:-1] + r_bounds[1:])

    # actual element widths (central difference)
    N            = len(r_mids)
    r_arr        = np.array(r_mids)
    delta_r      = np.empty(N)
    delta_r[0]   = r_arr[1]  - r_arr[0]
    delta_r[-1]  = r_arr[-1] - r_arr[-2]
    delta_r[1:-1]= (r_arr[2:] - r_arr[:-2]) / 2.0

    elements = []

    for i, r in enumerate(r_mids):
        # -- chord and twist ---------------------------------------------------
        if MODE == "Evaluate":
            chord = _BLADE_CHORD[i]
            twist = _BLADE_TWIST[i]

        elif BLADE_SHAPE_MODE == "optimum":
            chord = opt_chord(r, ROTOR_RADIUS, lam_geom, NUM_BLADES, CL_DESIGN)
            twist = opt_twist(r, ROTOR_RADIUS, lam_geom, ALPHA_DESIGN)

        else:
            chord = linear_chord(r)
            chord = max(chord, 0.005)
            twist = linear_twist(r)

        lr = lam * r / ROTOR_RADIUS
        el = bem_element(r, ROTOR_RADIUS, NUM_BLADES, chord,
                         twist, PITCH_0_DEG, lam, U)
        el.update(r=r, chord=chord, twist_deg=twist, lr=lr, dr=delta_r[i])
        elements.append(el)

    return elements, rotor_cp(elements, lam)


# --- CP-lambda sweep ---------------------------------------------------------------

def cp_sweep():
    """
    Sweep tip-speed-ratio from LAMBDA_MIN to LAMBDA_MAX and compute CP at each.

    The blade geometry (chord and twist) is FIXED at TIP_SPEED_RATIO for the
    entire sweep -- the same blade shape is evaluated at different wind speeds.
    This is physically correct: a real blade does not change shape with lambda.

    In Design mode with BLADE_SHAPE_MODE="optimum", lam_geometry=TIP_SPEED_RATIO
    ensures opt_chord() and opt_twist() are called once at the design point, not
    re-computed at every lambda.  For "auto_fit" and "manual" modes the linear
    coefficients are already fixed, so lam_geometry has no additional effect.

    In Evaluate mode geometry always comes from the file -- lam_geometry is
    irrelevant but passed for consistency.
    """
    lams = np.linspace(LAMBDA_MIN, LAMBDA_MAX, LAMBDA_STEPS)
    cps  = []
    for l in lams:
        _, cp = analyse_blade(l, lam_geometry=TIP_SPEED_RATIO)
        cps.append(cp)
    return lams, np.array(cps)


# --- Load airfoil -------------------------------------------------------------

def load_airfoil(path):
    pts = []
    with open(path) as fh:
        for line in fh:
            p = line.split()
            if len(p) == 2:
                pts.append((float(p[0]), float(p[1])))
    return np.array(pts)


# --- Output helpers -----------------------------------------------------------

def _savefig(name):
    p = os.path.join(OUTPUT_DIR, name)
    plt.savefig(p, dpi=150, bbox_inches="tight")
    print(f"  Plot saved -> {p}")


def save_xlsx(elements, lams, cps, path, pc=None):
    """
    Write all BEM results into a single formatted Excel workbook with
    three sheets (plus optional Power Curve sheet):

      Sheet 1 - Blade Elements   : per-element BEM results
      Sheet 2 - CP-Lambda Curve  : power coefficient vs tip-speed-ratio
      Sheet 3 - Design Summary   : all input parameters and key results
      Sheet 4 - Power Curve      : P, Q, RPM, T, CT vs wind speed (if pc given)
    """
    wb = openpyxl.Workbook()

    # ── shared styles ──────────────────────────────────────────────────────────
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                                 numbers as xl_numbers)
    from openpyxl.utils import get_column_letter

    HDR_FILL   = PatternFill("solid", fgColor="1F4E79")   # dark blue
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
    SUB_FILL   = PatternFill("solid", fgColor="2E75B6")   # medium blue
    SUB_FONT   = Font(bold=True, color="FFFFFF", size=10)
    ALT_FILL   = PatternFill("solid", fgColor="D6E4F0")   # light blue
    OK_FILL    = PatternFill("solid", fgColor="C6EFCE")   # green
    FAIL_FILL  = PatternFill("solid", fgColor="FFC7CE")   # red
    WARN_FILL  = PatternFill("solid", fgColor="FFEB9C")   # yellow
    CENTER     = Alignment(horizontal="center", vertical="center")
    LEFT       = Alignment(horizontal="left",   vertical="center")
    THIN       = Side(style="thin", color="AAAAAA")
    BORDER     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def _hdr_cell(ws, row, col, value, fill=None, font=None):
        c = ws.cell(row=row, column=col, value=value)
        c.fill      = fill or HDR_FILL
        c.font      = font or HDR_FONT
        c.alignment = CENTER
        c.border    = BORDER
        return c

    def _data_cell(ws, row, col, value, num_fmt=None, fill=None, align=None):
        c = ws.cell(row=row, column=col, value=value)
        if num_fmt: c.number_format = num_fmt
        if fill:    c.fill          = fill
        c.alignment = align or CENTER
        c.border    = BORDER
        return c

    def _autofit(ws, extra=4):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + extra, 30)

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 1 — Blade Elements
    # ══════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Blade Elements"
    ws1.freeze_panes = "A3"

    # title row
    ws1.merge_cells("A1:O1")
    t = ws1["A1"]
    t.value     = (f"Blade Element BEM Results   "
                   f"lambda={TIP_SPEED_RATIO}  B={NUM_BLADES}  "
                   f"R={ROTOR_RADIUS}m  U={WIND_SPEED}m/s  "
                   f"Mode={BLADE_SHAPE_MODE}")
    t.font      = Font(bold=True, color="FFFFFF", size=11)
    t.fill      = PatternFill("solid", fgColor="0D2B4E")
    t.alignment = CENTER

    headers = [
        ("r (m)",        "Radius",               "0.000"),
        ("c (m)",        "Chord",                "0.0000"),
        ("Twist (deg)",  "Geometric Twist",      "0.00"),
        ("AoA (deg)",    "Angle of Attack",      "0.00"),
        ("phi (deg)",    "Inflow Angle",         "0.00"),
        ("Cl",           "Lift Coefficient",     "0.0000"),
        ("Cd",           "Drag Coefficient",     "0.00000"),
        ("a",            "Axial Induction",      "0.0000"),
        ("a'",           "Angular Induction",    "0.0000"),
        ("F",            "Prandtl Loss",         "0.0000"),
        ("f_snel",       "Snel 3D Factor",       "0.0000"),
        ("V_rel (m/s)",  "Relative Velocity",    "0.00"),
        ("Re_local",     "Local Reynolds No.",   "#,##0"),
        ("CTr",          "Local Thrust Coeff",   "0.0000"),
        ("Converged",    "BEM Converged",        None),
    ]

    keys = ["r","chord","twist_deg","alpha_deg","phi_deg",
            "Cl","Cd","a","a_prime","F","f_snel",
            "V_rel","Re_local","CTr","converged"]

    # header row (row 2)
    for col, (short, long_, _) in enumerate(headers, 1):
        c = _hdr_cell(ws1, 2, col, short)
        c.comment = None
        # add long name as second line via wrap
        ws1.cell(row=2, column=col).alignment = CENTER

    # data rows
    for i, el in enumerate(elements, 3):
        fill = ALT_FILL if i % 2 == 0 else None
        for col, (_, _, fmt) in enumerate(headers, 1):
            key = keys[col - 1]
            val = el[key]
            if key == "converged":
                c = _data_cell(ws1, i, col, "YES" if val else "NO",
                               fill=OK_FILL if val else FAIL_FILL)
            else:
                _data_cell(ws1, i, col, val, num_fmt=fmt, fill=fill)

    # re_polar range column colour hint: green=OK, orange=LOW, purple=HIGH
    re_pmin = _POLARS[0]["re"]
    re_pmax = _POLARS[-1]["re"]
    re_col  = keys.index("Re_local") + 1
    for i, el in enumerate(elements, 3):
        re = el["Re_local"]
        if   re < re_pmin: c_fill = PatternFill("solid", fgColor="FFD700")
        elif re > re_pmax: c_fill = PatternFill("solid", fgColor="C0C0FF")
        else:              c_fill = PatternFill("solid", fgColor="C6EFCE")
        ws1.cell(row=i, column=re_col).fill = c_fill

    ws1.row_dimensions[1].height = 22
    ws1.row_dimensions[2].height = 30
    _autofit(ws1)

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 2 — CP-Lambda Curve
    # ══════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("CP-Lambda Curve")
    ws2.freeze_panes = "A3"

    ws2.merge_cells("A1:C1")
    t2 = ws2["A1"]
    t2.value     = "Power Coefficient vs Tip-Speed-Ratio"
    t2.font      = Font(bold=True, color="FFFFFF", size=11)
    t2.fill      = PatternFill("solid", fgColor="0D2B4E")
    t2.alignment = CENTER

    for col, label in enumerate(["lambda", "CP", "Note"], 1):
        _hdr_cell(ws2, 2, col, label)

    idx_max = int(np.argmax(cps))
    design_lam = TIP_SPEED_RATIO
    betz = 16.0 / 27.0

    for i, (lam, cp) in enumerate(zip(lams, cps), 3):
        fill = None
        note = ""
        if i - 3 == idx_max:
            fill = PatternFill("solid", fgColor="FFD700")
            note = "CP max"
        elif abs(lam - design_lam) < (lams[1] - lams[0]) * 0.51:
            fill = PatternFill("solid", fgColor="C6EFCE")
            note = "Design point"
        row_fill = ALT_FILL if i % 2 == 0 else None
        _data_cell(ws2, i, 1, round(lam, 4), "0.000",  fill=fill or row_fill)
        _data_cell(ws2, i, 2, round(cp, 6),  "0.0000", fill=fill or row_fill)
        _data_cell(ws2, i, 3, note, fill=fill or row_fill, align=LEFT)

    # summary block to the right
    summary = [
        ("CP max",           f"{cps[idx_max]:.5f}"),
        ("lambda at CP max", f"{lams[idx_max]:.3f}"),
        ("CP at design lambda", f"{float(np.interp(design_lam, lams, cps)):.5f}"),
        ("Betz limit",       f"{betz:.5f}"),
        ("CP / Betz (%)",    f"{float(np.interp(design_lam, lams, cps))/betz*100:.1f}"),
    ]
    for i, (label, val) in enumerate(summary, 3):
        _hdr_cell(ws2, i, 5, label, fill=SUB_FILL, font=SUB_FONT)
        _data_cell(ws2, i, 6, val, align=LEFT)

    ws2.row_dimensions[1].height = 22
    ws2.row_dimensions[2].height = 20
    _autofit(ws2)

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 3 — Design Summary
    # ══════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Design Summary")

    ws3.merge_cells("A1:C1")
    t3 = ws3["A1"]
    t3.value     = "Wind Turbine Blade Design - Input Parameters & Key Results"
    t3.font      = Font(bold=True, color="FFFFFF", size=11)
    t3.fill      = PatternFill("solid", fgColor="0D2B4E")
    t3.alignment = CENTER

    sections = [
        ("Rotor Geometry", [
            ("Rotor radius R",          ROTOR_RADIUS,        "m"),
            ("Hub radius r_hub",        HUB_RADIUS,          "m"),
            ("Hub/Rotor ratio",         round(HUB_RADIUS/ROTOR_RADIUS, 4), ""),
            ("Number of blades B",      NUM_BLADES,          ""),
            ("Number of elements N",    NUM_ELEMENTS,        ""),
        ]),
        ("Flow Conditions", [
            ("Wind speed U_inf",        WIND_SPEED,          "m/s"),
            ("Tip speed ratio lambda",  TIP_SPEED_RATIO,     ""),
            ("Angular velocity Omega",  round(TIP_SPEED_RATIO*WIND_SPEED/ROTOR_RADIUS, 3), "rad/s"),
            ("Air density rho",         AIR_DENSITY,         "kg/m3"),
            ("Dynamic viscosity mu",    DYN_VISCOSITY,       "Pa*s"),
        ]),
        ("Aerodynamic Design", [
            ("Design AoA alpha",        ALPHA_DESIGN,        "deg"),
            ("Design Cl",               CL_DESIGN,           ""),
            ("Blade shape mode",        BLADE_SHAPE_MODE,    ""),
            ("Induction correction",    INDUCTION_CORRECTION,""),
            ("Hub loss correction",     HUB_LOSS,            ""),
            ("Snel 3D correction",      SNEL_3D_CORRECTION,  ""),
        ]),
        ("Polar Data", [
            ("Polar folder",            POLAR_DIR,           ""),
            ("Polars loaded",           len(_POLARS),        ""),
            ("Re range covered",        f"{_POLARS[0]['re']:,.0f} - {_POLARS[-1]['re']:,.0f}", ""),
        ]),
        ("Key Results", [
            ("Power coefficient CP",    round(float(np.interp(TIP_SPEED_RATIO, lams, cps)), 5), ""),
            ("Betz limit",              round(16/27, 5),     ""),
            ("CP / Betz limit",         f"{float(np.interp(TIP_SPEED_RATIO,lams,cps))/(16/27)*100:.1f}", "%"),
            ("CP max",                  round(float(np.max(cps)), 5), ""),
            ("lambda at CP max",        round(float(lams[np.argmax(cps)]), 3), ""),
            ("BEM convergence",         f"{sum(e['converged'] for e in elements)}/{len(elements)}", ""),
            ("Chord at hub",            round(elements[0]["chord"],  4), "m"),
            ("Chord at tip",            round(elements[-1]["chord"], 4), "m"),
            ("Twist at hub",            round(elements[0]["twist_deg"],  2), "deg"),
            ("Twist at tip",            round(elements[-1]["twist_deg"], 2), "deg"),
        ]),
    ]

    row = 3
    for section_name, params in sections:
        # section header
        ws3.merge_cells(f"A{row}:C{row}")
        c = ws3.cell(row=row, column=1, value=section_name)
        c.fill      = SUB_FILL
        c.font      = SUB_FONT
        c.alignment = LEFT
        c.border    = BORDER
        row += 1
        for label, value, unit in params:
            fill = ALT_FILL if row % 2 == 0 else None
            _data_cell(ws3, row, 1, label, fill=fill, align=LEFT)
            _data_cell(ws3, row, 2, value, fill=fill)
            _data_cell(ws3, row, 3, unit,  fill=fill, align=LEFT)
            row += 1
        row += 1   # blank row between sections

    ws3.row_dimensions[1].height = 22
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 20
    ws3.column_dimensions["C"].width = 10

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 4 — Power Curve
    # ══════════════════════════════════════════════════════════════════════════
    if pc is not None:
        ws4 = wb.create_sheet("Power Curve")
        ws4.freeze_panes = "A3"

        ws4.merge_cells("A1:I1")
        t4 = ws4["A1"]
        t4.value     = (f"Power Curve  —  lambda = {TIP_SPEED_RATIO}  "
                        f"(variable speed, fixed blade geometry)  "
                        f"U = {POWER_CURVE_UMIN}-{POWER_CURVE_UMAX} m/s")
        t4.font      = Font(bold=True, color="FFFFFF", size=11)
        t4.fill      = PatternFill("solid", fgColor="0D2B4E")
        t4.alignment = CENTER

        pc_headers = [
            ("U (m/s)",     "0.00"),
            ("Ω (rad/s)",   "0.000"),
            ("RPM",         "0.00"),
            ("CP",          "0.0000"),
            ("P1 - CP (W)", "0.0"),
            ("Q (N·m)",     "0.00"),
            ("P2 - Torque (W)", "0.0"),
            ("T (N)",       "0.00"),
            ("CT",          "0.0000"),
        ]
        pc_keys = ["U", "Omega", "RPM", "CP", "P1", "Q", "P2", "T", "CT"]

        for col, (label, _) in enumerate(pc_headers, 1):
            _hdr_cell(ws4, 2, col, label)

        for i, idx in enumerate(range(len(pc["U"])), 3):
            fill = ALT_FILL if i % 2 == 0 else None
            for col, (_, fmt) in enumerate(pc_headers, 1):
                _data_cell(ws4, i, col,
                           round(float(pc[pc_keys[col-1]][idx]), 4),
                           num_fmt=fmt, fill=fill)

        ws4.row_dimensions[1].height = 22
        ws4.row_dimensions[2].height = 20
        _autofit(ws4)

    # ── save ──────────────────────────────────────────────────────────────────
    wb.save(path)
    n_sheets = 4 if pc is not None else 3
    print(f"  Excel saved -> {path}  "
          f"({len(elements)} elements, {len(lams)} CP points, {n_sheets} sheets)")


# --- Plots --------------------------------------------------------------------

def plot_airfoil(pts):
    fig, ax = plt.subplots(figsize=(9, 3))
    ax.plot(pts[:, 0], pts[:, 1], "k-", lw=1.5)
    ax.set_aspect("equal")
    ax.set(xlabel="x/c", ylabel="y/c", title="Airfoil profile")
    ax.grid(True, ls="--", alpha=0.4)
    _savefig("airfoil_shape.png")
    plt.show()


def plot_blade(elements):
    r      = [e["r"]          for e in elements]
    chord  = [e["chord"]      for e in elements]
    twist  = [e["twist_deg"]  for e in elements]
    alpha  = [e["alpha_deg"]  for e in elements]
    Cl     = [e["Cl"]         for e in elements]
    a_ax   = [e["a"]          for e in elements]
    a_ang  = [e["a_prime"]    for e in elements]
    phi    = [e["phi_deg"]    for e in elements]

    fig, axes = plt.subplots(2, 3, figsize=(15, 8))
    fig.suptitle(
        f"Blade Design  lambda={TIP_SPEED_RATIO}, B={NUM_BLADES}, "
        f"R={ROTOR_RADIUS} m, Uinf={WIND_SPEED} m/s",
        fontsize=13)

    def _ax(i, j, xs, ys, lbl, ttl, col="steelblue",
            hline=None, hline_lbl=None):
        axes[i, j].plot(xs, ys, color=col, marker="o", ms=3, lw=1.5)
        if hline is not None:
            axes[i, j].axhline(hline, color="k", ls="--",
                                label=hline_lbl, alpha=0.7)
            axes[i, j].legend(fontsize=8)
        axes[i, j].set(xlabel="Radius (m)", ylabel=lbl, title=ttl)
        axes[i, j].grid(True, ls="--", alpha=0.4)

    _ax(0, 0, r, chord,  "Chord (m)",          "Chord distribution")
    _ax(0, 1, r, twist,  "Twist (deg)",           "Twist distribution", col="firebrick")
    _ax(0, 2, r, phi,    "Inflow angle phi (deg)",  "Inflow angle",       col="darkorange")
    _ax(1, 0, r, alpha,  "AoA (deg)",             "Angle of attack",
        col="seagreen", hline=ALPHA_DESIGN, hline_lbl=f"Design alpha={ALPHA_DESIGN}deg")
    _ax(1, 1, r, Cl,     "Cl",                  "Lift coefficient",
        col="purple", hline=CL_DESIGN, hline_lbl=f"Design Cl={CL_DESIGN}")
    _ax(1, 2, r, a_ax,   "Axial induction a",   "Induction factors",  col="navy")
    axes[1, 2].plot(r, a_ang, color="tomato", marker="s", ms=3, lw=1.5,
                    label="a' (angular)")
    axes[1, 2].axhline(1 / 3, color="k", ls="--", alpha=0.5,
                       label="Ideal a = 1/3")
    axes[1, 2].legend(fontsize=8)

    plt.tight_layout()
    _savefig("blade_design.png")
    plt.show()


def plot_cp(lams, cps, cp_design):
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.plot(lams, cps, "steelblue", lw=2, label="CP-lambda curve")
    idx = int(np.argmax(cps))
    ax.plot(lams[idx], cps[idx], "ro", ms=9,
            label=f"CP,max = {cps[idx]:.4f}  @  lambda = {lams[idx]:.2f}")
    ax.plot(TIP_SPEED_RATIO, cp_design, "g^", ms=9,
            label=f"Design point  CP = {cp_design:.4f}  @  lambda = {TIP_SPEED_RATIO}")
    ax.axhline(16 / 27, color="gray", ls="--", alpha=0.7,
               label="Betz limit = 0.593")
    ax.set(xlabel="Tip-speed ratio lambda", ylabel="Power coefficient CP",
           title="CP - lambda Curve")
    ax.legend()
    ax.grid(True, ls="--", alpha=0.4)
    _savefig("cp_lambda.png")
    plt.show()


# --- Step 7 comparison: optimum vs linear blade shape ------------------------

def plot_shape_comparison(r_mids):
    """
    Overlay the Glauert-optimum chord/twist against the fitted linear
    approximation (Eqs. 3.115-3.116) so the user can judge the fit quality.
    """
    c_opt  = [opt_chord(r, ROTOR_RADIUS, TIP_SPEED_RATIO,
                        NUM_BLADES, CL_DESIGN) for r in r_mids]
    t_opt  = [opt_twist(r, ROTOR_RADIUS, TIP_SPEED_RATIO,
                        ALPHA_DESIGN) for r in r_mids]
    c_lin  = [linear_chord(r) for r in r_mids]
    t_lin  = [linear_twist(r) for r in r_mids]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4))
    fig.suptitle("Step 7 - Blade Shape: Optimum vs Linear Approximation",
                 fontsize=12)

    ax1.plot(r_mids, c_opt, "b--", lw=1.5, label="Glauert optimum")
    ax1.plot(r_mids, c_lin, "r-",  lw=2,   label=f"Linear: c = {_A1:+.4f}*r + {_B1:+.4f}")
    ax1.set(xlabel="Radius (m)", ylabel="Chord (m)",
            title="Chord distribution  (Eq. 3.115)")
    ax1.legend(fontsize=8); ax1.grid(True, ls="--", alpha=0.4)

    ax2.plot(r_mids, t_opt, "b--", lw=1.5, label="Glauert optimum")
    ax2.plot(r_mids, t_lin, "r-",  lw=2,
             label=f"Linear: theta = {_A2:+.4f}*(R - r)")
    ax2.set(xlabel="Radius (m)", ylabel="Twist (deg)",
            title="Twist distribution  (Eq. 3.116)")
    ax2.legend(fontsize=8); ax2.grid(True, ls="--", alpha=0.4)

    plt.tight_layout()
    _savefig("blade_shape_comparison.png")
    plt.show()


# --- Polar plot ---------------------------------------------------------------

def plot_polar():
    """Plot Cl, Cd, and Cl/Cd vs alpha for all loaded polars, overlaid."""
    colours = plt.cm.tab10.colors

    fig, axes = plt.subplots(1, 3, figsize=(15, 4))
    n_polars  = len(_POLARS)
    fig.suptitle(f"Airfoil Polars  ({n_polars} Re{'s' if n_polars > 1 else ''})", fontsize=12)

    for i, p in enumerate(_POLARS):
        col   = colours[i % len(colours)]
        label = f"Re={p['re']/1e3:.0f}k"
        with np.errstate(divide='ignore', invalid='ignore'):
            ld = np.where(p["cd"] > 0, p["cl"] / p["cd"], 0.0)

        axes[0].plot(p["alpha"], p["cl"], color=col, lw=1.5, label=label)
        axes[1].plot(p["alpha"], p["cd"], color=col, lw=1.5, label=label)
        axes[2].plot(p["alpha"], ld,      color=col, lw=1.5, label=label)

    for ax in axes:
        ax.axvline(ALPHA_DESIGN, color="k", ls="--", lw=1,
                   label=f"Design alpha={ALPHA_DESIGN}deg")
        ax.legend(fontsize=7)
        ax.grid(True, ls="--", alpha=0.4)

    axes[0].set(xlabel="Alpha (deg)", ylabel="Cl", title="Lift curve")
    axes[1].set(xlabel="Alpha (deg)", ylabel="Cd", title="Drag polar")
    axes[2].set(xlabel="Alpha (deg)", ylabel="Cl/Cd", title="Lift-to-drag ratio")

    plt.tight_layout()
    _savefig("polar.png")
    plt.show()


# --- Reynolds number comparison and recommendation ----------------------------

def reynolds_analysis(elements):
    """
    Compare the local Re at every blade element against the Re range
    covered by the loaded polars.

    Status flags:
      [OK]    element Re is inside the polar Re range  (interpolated)
      [LOW]   element Re is below the lowest polar Re  (clamped to lowest)
      [HIGH]  element Re is above the highest polar Re (clamped to highest)

    Also reports a chord-weighted average Re and recommends adding a polar
    at that Re if it falls outside the covered range.
    """
    re_vals  = np.array([e["Re_local"] for e in elements])
    r_vals   = np.array([e["r"]        for e in elements])
    c_vals   = np.array([e["chord"]    for e in elements])
    v_vals   = np.array([e["V_rel"]    for e in elements])

    Re_avg      = np.average(re_vals, weights=c_vals)
    Re_min      = re_vals.min()
    Re_max      = re_vals.max()
    Re_mid_span = float(np.interp(0.5 * (HUB_RADIUS + ROTOR_RADIUS),
                                  r_vals, re_vals))

    re_polar_min = _POLARS[0]["re"]
    re_polar_max = _POLARS[-1]["re"]
    n_polars     = len(_POLARS)

    # recommended Re: chord-weighted avg rounded to nearest 50k
    Re_recommended = int(round(Re_avg / 50_000.0) * 50_000)

    print(f"\n[Reynolds Number Analysis]")
    print(f"  Polars loaded                 : {n_polars}  "
          f"(Re = {re_polar_min:,.0f} -- {re_polar_max:,.0f})")
    print(f"  Local Re range (BEM result)   : "
          f"{Re_min:>12,.0f}  to  {Re_max:>12,.0f}")
    print(f"  Chord-weighted average Re     : {Re_avg:>12,.0f}")
    print(f"  Re at mid-span                : {Re_mid_span:>12,.0f}")
    print()

    hdr = (f"{'r(m)':>7} {'c(m)':>7} {'Vrel(m/s)':>10} "
           f"{'Re_local':>12} {'Re_polar_range':>20} {'status':>7}")
    print(hdr)
    print("-" * len(hdr))

    n_ok = n_low = n_high = 0
    for e in elements:
        re  = e["Re_local"]
        rng = f"{re_polar_min:,.0f}-{re_polar_max:,.0f}"
        if   re < re_polar_min: status = "[LOW] "; n_low  += 1
        elif re > re_polar_max: status = "[HIGH]"; n_high += 1
        else:                   status = "[OK]  "; n_ok   += 1
        print(f"{e['r']:7.3f} {e['chord']:7.4f} {e['V_rel']:10.2f} "
              f"{re:12,.0f} {rng:>20} {status}")

    print()
    print(f"  Summary: {n_ok} interpolated (OK),  "
          f"{n_low} below range (LOW),  "
          f"{n_high} above range (HIGH)  "
          f"(out of {len(elements)} elements)")

    # Recommendation
    print(f"\n[Recommendation]")
    if n_low == 0 and n_high == 0:
        print(f"  All elements are within the polar Re range "
              f"({re_polar_min:,.0f} -- {re_polar_max:,.0f}).")
        print(f"  Interpolation is fully valid. No additional polars needed.")
    else:
        if n_low > 0:
            print(f"  {n_low} element(s) below Re={re_polar_min:,.0f}  "
                  f"--> add a polar at Re ~ {int(Re_min/50000)*50000:,.0f} or lower.")
        if n_high > 0:
            print(f"  {n_high} element(s) above Re={re_polar_max:,.0f}  "
                  f"--> add a polar at Re ~ {int(np.ceil(Re_max/50000))*50000:,.0f} or higher.")
        print(f"  Chord-weighted average Re = {Re_avg:,.0f}  "
              f"(recommended representative Re = {Re_recommended:,.0f})")

    return re_vals, re_polar_min, re_polar_max, Re_recommended


def plot_reynolds(elements, re_polar_min, re_polar_max, Re_recommended):
    """
    Plot local Re vs span.

    Shaded band = Re range covered by loaded polars (fully interpolated).
    Orange dots = elements below polar range (clamped to lowest polar).
    Purple dots = elements above polar range (clamped to highest polar).
    Blue dots   = elements inside polar range (interpolated correctly).
    Green dashed line = chord-weighted recommended Re.
    """
    r_vals  = np.array([e["r"]        for e in elements])
    re_vals = np.array([e["Re_local"] for e in elements])

    fig, ax = plt.subplots(figsize=(10, 5))

    # shaded band = polar coverage
    ax.axhspan(re_polar_min, re_polar_max,
               color="steelblue", alpha=0.12,
               label=f"Polar Re coverage ({re_polar_min/1e3:.0f}k -- {re_polar_max/1e3:.0f}k)")

    # boundary lines
    ax.axhline(re_polar_min, color="steelblue", ls="--", lw=1.2,
               label=f"Min polar Re = {re_polar_min:,.0f}")
    ax.axhline(re_polar_max, color="steelblue", ls="-",  lw=1.2,
               label=f"Max polar Re = {re_polar_max:,.0f}")

    # recommended Re line
    ax.axhline(Re_recommended, color="green", ls="--", lw=1.5,
               label=f"Chord-weighted avg Re = {Re_recommended:,.0f}")

    # local Re curve
    ax.plot(r_vals, re_vals, "k-", lw=1.2, zorder=3)

    # colour-coded dots
    colours = []
    for re in re_vals:
        if   re < re_polar_min: colours.append("orange")
        elif re > re_polar_max: colours.append("purple")
        else:                   colours.append("steelblue")
    ax.scatter(r_vals, re_vals, c=colours, zorder=5, s=35)

    # mark loaded polar Re values as horizontal tick marks
    for p in _POLARS:
        ax.axhline(p["re"], color="gray", ls=":", lw=0.7, alpha=0.6)

    ax.set(xlabel="Span radius (m)",
           ylabel="Reynolds number",
           title=f"Local Re vs Polar Coverage  ({len(_POLARS)} polars loaded)")
    ax.legend(fontsize=8)
    ax.grid(True, ls="--", alpha=0.4)
    ax.yaxis.set_major_formatter(
        plt.FuncFormatter(lambda x, _: f"{x/1e6:.2f}M" if x >= 1e6
                          else f"{x/1e3:.0f}k"))

    plt.tight_layout()
    _savefig("reynolds_comparison.png")
    plt.show()



# =============================================================================
#  HAWC2 c2_def OUTPUT  (independent section — does not modify main code)
# =============================================================================
#
#  Generates a formatted Excel file containing the HAWC2 c2_def block:
#
#    begin c2_def;
#        nsec N;
#        ;  nr    x         y         z(radius)   twist
#        sec  1   x1        y1        z1          t1 ;
#        ...
#    end c2_def;
#
#  Coordinate system (from HAWC2 manual / DTU slides):
#    Z  = spanwise radius along blade axis        = e["r"]
#    twist = theta_z about Z (right-hand rule,
#            positive nose-up, zero at tip)       = e["twist_deg"]
#
#  Half-chord point offset from the quarter-chord (pitch axis):
#    In the untwisted cross-section frame:
#      dx_local = -0.25 * chord   (toward trailing edge, negative X_c2)
#      dy_local =  0.0            (no camber-line offset used here)
#    Rotated into main-body frame by twist angle:
#      x = dx_local * cos(twist) - dy_local * sin(twist)
#      y = dx_local * sin(twist) + dy_local * cos(twist)
#
#  Usage:
#    1. Run main() first to get `elements` and `lams`, `cps`
#    2. Call:  generate_hawc2_c2def(elements, "hawc2_c2def.xlsx")
#    3. The file is written to OUTPUT_DIR
#
#  To run standalone (after main has already been run in the same session):
#    generate_hawc2_c2def(elements, "hawc2_c2def.xlsx")
# =============================================================================

def _hawc2_half_chord_coords(chord, twist_deg):
    """
    Compute the (x, y) position of the half-chord point in the HAWC2
    main-body coordinate system.

    The pitch axis in this code is the quarter-chord point.
    The half-chord point is offset by c/4 toward the trailing edge
    from the quarter-chord in the local cross-section frame, then
    rotated by the local twist angle into the main-body frame.

    Parameters
    ----------
    chord     : local chord length  [m]
    twist_deg : local twist angle   [deg]  (positive = nose-up)

    Returns
    -------
    x, y : half-chord coordinates in main-body frame  [m]
    """
    # offset from quarter-chord to half-chord in untwisted local frame
    # X_c2 points toward leading edge, so half-chord is at -c/4 in X_c2
    dx_local = -0.25 * chord   # toward trailing edge
    dy_local =  0.0            # no offset in suction-side direction

    # rotate into main-body frame by twist angle (right-hand rule about Z)
    theta = np.radians(twist_deg)
    x = dx_local * np.cos(theta) - dy_local * np.sin(theta)
    y = dx_local * np.sin(theta) + dy_local * np.cos(theta)
    return x, y


def generate_hawc2_c2def(elements, filename="hawc2_c2def.xlsx"):
    """
    Generate a HAWC2 c2_def block and write it to a formatted Excel file.

    The file contains two sheets:
      Sheet 1 - c2_def Table  : the numeric data ready to copy into HAWC2
      Sheet 2 - HAWC2 Text    : the complete formatted c2_def text block
                                 (can be copy-pasted directly into the .htc file)

    Parameters
    ----------
    elements : list of element dicts from analyse_blade()
    filename : output filename (saved to OUTPUT_DIR)
    """
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side)
    from openpyxl.utils import get_column_letter

    # -- build c2_def data ---------------------------------------------------
    sections = []
    for nr, el in enumerate(elements, 1):
        z     = el["r"]
        chord = el["chord"]
        twist = el["twist_deg"]
        x, y  = _hawc2_half_chord_coords(chord, twist)
        sections.append({
            "nr"   : nr,
            "x"    : x,
            "y"    : y,
            "z"    : z,
            "twist": twist,
            "chord": chord,   # informational
        })

    # -- styles --------------------------------------------------------------
    HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
    HDR_FONT  = Font(bold=True, color="FFFFFF", size=10)
    SUB_FILL  = PatternFill("solid", fgColor="2E75B6")
    SUB_FONT  = Font(bold=True, color="FFFFFF", size=10)
    ALT_FILL  = PatternFill("solid", fgColor="D6E4F0")
    TIP_FILL  = PatternFill("solid", fgColor="C6EFCE")   # green for tip (ref)
    ROOT_FILL = PatternFill("solid", fgColor="FFD700")   # gold for root
    CENTER    = Alignment(horizontal="center", vertical="center")
    LEFT      = Alignment(horizontal="left",   vertical="center")
    MONO      = Font(name="Courier New", size=9)
    THIN      = Side(style="thin", color="AAAAAA")
    BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def _hcell(ws, row, col, val, fill=None, font=None, align=None):
        c = ws.cell(row=row, column=col, value=val)
        c.fill      = fill  or HDR_FILL
        c.font      = font  or HDR_FONT
        c.alignment = align or CENTER
        c.border    = BORDER
        return c

    def _dcell(ws, row, col, val, fmt=None, fill=None, align=None):
        c = ws.cell(row=row, column=col, value=val)
        if fmt:  c.number_format = fmt
        if fill: c.fill          = fill
        c.alignment = align or CENTER
        c.border    = BORDER
        return c

    wb = openpyxl.Workbook()

    # =========================================================================
    # SHEET 1 — c2_def numeric table
    # =========================================================================
    ws1 = wb.active
    ws1.title        = "c2_def Table"
    ws1.freeze_panes = "A3"

    # title
    ws1.merge_cells("A1:G1")
    t = ws1["A1"]
    t.value     = (f"HAWC2 c2_def  |  "
                   f"R={ROTOR_RADIUS}m  B={NUM_BLADES}  "
                   f"lambda={TIP_SPEED_RATIO}  mode={BLADE_SHAPE_MODE}  "
                   f"Pitch axis=quarter-chord  Half-chord offsets applied")
    t.font      = Font(bold=True, color="FFFFFF", size=11)
    t.fill      = PatternFill("solid", fgColor="0D2B4E")
    t.alignment = CENTER
    ws1.row_dimensions[1].height = 22

    # column headers
    cols = [
        ("nr",         "Section No.",       None),
        ("x (m)",      "Half-chord X\n(main body)",  "0.00000"),
        ("y (m)",      "Half-chord Y\n(main body)",  "0.00000"),
        ("z / r (m)",  "Radius\n(spanwise)",          "0.000"),
        ("twist (deg)","Twist theta_z\n(right-hand)", "0.000"),
        ("chord (m)",  "Chord\n(informational)",      "0.0000"),
        ("note",       "Note",              None),
    ]
    for col, (short, long_, _) in enumerate(cols, 1):
        c = _hcell(ws1, 2, col, short)
        c.alignment = CENTER

    ws1.row_dimensions[2].height = 28

    # data
    for i, s in enumerate(sections, 3):
        if   i == 3:           fill = ROOT_FILL    # root
        elif i == len(sections) + 2: fill = TIP_FILL   # tip
        else:                  fill = ALT_FILL if i % 2 == 0 else None

        note = ""
        if i == 3:                    note = "Root section"
        elif i == len(sections) + 2:  note = "Tip section (twist ref = 0)"

        _dcell(ws1, i, 1, s["nr"],    fmt="0",        fill=fill)
        _dcell(ws1, i, 2, s["x"],     fmt="0.000000", fill=fill)
        _dcell(ws1, i, 3, s["y"],     fmt="0.000000", fill=fill)
        _dcell(ws1, i, 4, s["z"],     fmt="0.000",    fill=fill)
        _dcell(ws1, i, 5, s["twist"], fmt="0.000",    fill=fill)
        _dcell(ws1, i, 6, s["chord"], fmt="0.0000",   fill=fill)
        _dcell(ws1, i, 7, note,       fill=fill, align=LEFT)

    # autofit
    for col in ws1.columns:
        ltr = get_column_letter(col[0].column)
        mx  = max((len(str(c.value or "")) for c in col), default=8)
        ws1.column_dimensions[ltr].width = min(mx + 4, 28)

    # -- summary block --------------------------------------------------------
    sr = len(sections) + 4
    ws1.merge_cells(f"A{sr}:G{sr}")
    ws1.cell(row=sr, column=1).value = "Coordinate System Notes"
    ws1.cell(row=sr, column=1).fill  = SUB_FILL
    ws1.cell(row=sr, column=1).font  = SUB_FONT
    ws1.cell(row=sr, column=1).alignment = LEFT

    notes = [
        ("Z axis",       "Spanwise, hub to tip"),
        ("X axis (X_c2)","Points toward leading edge in cross-section frame"),
        ("Y axis (Y_c2)","Points toward suction side (upward)"),
        ("Pitch axis",   "Quarter-chord point (c/4 from leading edge)"),
        ("Half-chord",   "Offset -c/4 from pitch axis toward trailing edge, then rotated by twist"),
        ("Twist sign",   "Positive = leading edge rotated into wind (right-hand rule about Z)"),
        ("Tip twist",    f"{sections[-1]['twist']:.3f} deg  (should be 0 or near-zero)"),
        ("Root twist",   f"{sections[0]['twist']:.3f} deg"),
    ]
    for j, (lbl, val) in enumerate(notes, sr + 1):
        fill = ALT_FILL if j % 2 == 0 else None
        ws1.merge_cells(f"A{j}:C{j}")
        c = ws1.cell(row=j, column=1, value=lbl)
        c.font = Font(bold=True, size=9)
        c.fill = fill or PatternFill(); c.alignment = LEFT; c.border = BORDER
        ws1.merge_cells(f"D{j}:G{j}")
        c2 = ws1.cell(row=j, column=4, value=val)
        c2.fill = fill or PatternFill(); c2.alignment = LEFT; c2.border = BORDER

    # =========================================================================
    # SHEET 2 — HAWC2 text block (copy-paste into .htc file)
    # =========================================================================
    ws2 = wb.create_sheet("HAWC2 Text Block")

    ws2.merge_cells("A1:A1")
    h = ws2["A1"]
    h.value     = "HAWC2 c2_def text block  --  copy and paste directly into your .htc file"
    h.font      = Font(bold=True, color="FFFFFF", size=10)
    h.fill      = PatternFill("solid", fgColor="0D2B4E")
    h.alignment = LEFT
    ws2.row_dimensions[1].height = 20

    lines = []
    lines.append("begin c2_def;  Definition of centerline (main_body coordinates)")
    lines.append(f"    nsec {len(sections)} ;")
    lines.append(f"    ;{'nr':>5}  {'x':>10}  {'y':>10}  {'z (radius)':>12}  {'twist':>10}")
    for s in sections:
        line = (f"    sec {s['nr']:>3}  "
                f"{s['x']:>10.4f}  "
                f"{s['y']:>10.4f}  "
                f"{s['z']:>12.3f}  "
                f"{s['twist']:>10.3f} ;")
        lines.append(line)
    lines.append("end c2_def ;")

    for row, line in enumerate(lines, 2):
        c = ws2.cell(row=row, column=1, value=line)
        c.font      = MONO
        c.alignment = LEFT
        fill = ALT_FILL if row % 2 == 0 else None
        if row == 2 or row == len(lines) + 1:  # begin/end lines
            c.font = Font(name="Courier New", size=9, bold=True, color="1F4E79")
        if fill:
            c.fill = fill

    ws2.column_dimensions["A"].width = 72
    ws2.row_dimensions[1].height = 20

    # =========================================================================
    # save
    # =========================================================================
    out_path = os.path.join(OUTPUT_DIR, filename)
    wb.save(out_path)
    print(f"\n[HAWC2] c2_def written -> {out_path}")
    print(f"  Sections : {len(sections)}")
    print(f"  z range  : {sections[0]['z']:.3f} m  to  {sections[-1]['z']:.3f} m")
    print(f"  twist    : {sections[0]['twist']:.3f} deg (root)  ->  "
          f"{sections[-1]['twist']:.3f} deg (tip)")
    print(f"  x range  : {min(s['x'] for s in sections):.4f} m  to  "
          f"{max(s['x'] for s in sections):.4f} m")
    print(f"  Sheets   : 'c2_def Table'  |  'HAWC2 Text Block'")
    return out_path


# --- Power curve (variable speed, constant lambda) ----------------------------

def power_curve():
    """
    Compute power, torque, RPM, thrust and CT vs wind speed at constant
    tip-speed ratio (variable-speed operation).

    For each wind speed U:
      - Omega  = TIP_SPEED_RATIO * U / R          (variable speed)
      - Run BEM at that U with fixed blade geometry
      - Method 1:  P1 = CP * 0.5 * rho * pi * R^2 * U^3
      - Method 2:  dF_theta = 0.5*rho*Vrel^2*c*(Cl*sin(phi)-Cd*cos(phi))*dr
                   Q  = B * sum(dF_theta * r)
                   P2 = Q * Omega
      - Thrust:    dF_x = 0.5*rho*Vrel^2*c*(Cl*cos(phi)+Cd*sin(phi))*dr
                   T  = B * sum(dF_x)
                   CT = T / (0.5 * rho * pi * R^2 * U^2)

    Returns dict with arrays: U, Omega, RPM, CP, P1, Q, P2, T, CT
    """
    U_arr   = np.linspace(POWER_CURVE_UMIN, POWER_CURVE_UMAX, POWER_CURVE_STEPS)
    A_rotor = np.pi * ROTOR_RADIUS**2

    res = dict(U=[], Omega=[], RPM=[], CP=[], P1=[], Q=[], P2=[], T=[], CT=[])

    for U in U_arr:
        Omega = TIP_SPEED_RATIO * U / ROTOR_RADIUS
        RPM   = Omega * 60.0 / (2.0 * np.pi)

        # run BEM at this wind speed with fixed geometry
        elements, CP = analyse_blade(TIP_SPEED_RATIO,
                                     lam_geometry=TIP_SPEED_RATIO,
                                     wind_speed_override=U)

        # Method 1 — CP based
        P1 = CP * 0.5 * AIR_DENSITY * A_rotor * U**3

        # Method 2 — torque integration
        Q  = 0.0
        T  = 0.0
        for e in elements:
            phi    = np.radians(e["phi_deg"])
            Vrel   = e["V_rel"]
            c      = e["chord"]
            Cl     = e["Cl"]
            Cd     = e["Cd"]
            r      = e["r"]
            dr     = e["dr"]
            q_dyn  = 0.5 * AIR_DENSITY * Vrel**2 * c * dr
            dF_t   = q_dyn * (Cl * np.sin(phi) - Cd * np.cos(phi))
            dF_x   = q_dyn * (Cl * np.cos(phi) + Cd * np.sin(phi))
            Q     += dF_t * r
            T     += dF_x

        Q  *= NUM_BLADES
        T  *= NUM_BLADES
        P2  = Q * Omega
        CT  = T / (0.5 * AIR_DENSITY * A_rotor * U**2)

        res["U"].append(U);     res["Omega"].append(Omega)
        res["RPM"].append(RPM); res["CP"].append(CP)
        res["P1"].append(P1);   res["Q"].append(Q)
        res["P2"].append(P2);   res["T"].append(T)
        res["CT"].append(CT)

    for k in res:
        res[k] = np.array(res[k])
    return res


def plot_power_curve_power(pc):
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.plot(pc["U"], pc["P1"] / 1000, "steelblue", lw=2,
            label="Method 1 — CP based")
    ax.plot(pc["U"], pc["P2"] / 1000, "firebrick",  lw=2, ls="--",
            label="Method 2 — Torque based")
    ax.set(xlabel="Wind speed (m/s)", ylabel="Power (kW)",
           title=f"Power Curve  (lambda = {TIP_SPEED_RATIO}, variable speed)")
    ax.legend()
    ax.grid(True, ls="--", alpha=0.4)
    _savefig("power_curve_power.png")
    plt.show()


def plot_power_curve_torque_rpm(pc):
    fig, ax1 = plt.subplots(figsize=(8, 5))
    ax1.plot(pc["U"], pc["Q"], "darkorange", lw=2, label="Torque (N·m)")
    ax1.set_xlabel("Wind speed (m/s)")
    ax1.set_ylabel("Torque (N·m)", color="darkorange")
    ax1.tick_params(axis="y", labelcolor="darkorange")
    ax1.grid(True, ls="--", alpha=0.4)

    ax2 = ax1.twinx()
    ax2.plot(pc["U"], pc["RPM"], "steelblue", lw=2, ls="--", label="RPM")
    ax2.set_ylabel("Rotor speed (RPM)", color="steelblue")
    ax2.tick_params(axis="y", labelcolor="steelblue")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2)
    fig.suptitle(f"Torque & RPM vs Wind Speed  (lambda = {TIP_SPEED_RATIO})")
    _savefig("power_curve_torque_rpm.png")
    plt.show()


def plot_power_curve_thrust_ct(pc):
    fig, ax1 = plt.subplots(figsize=(8, 5))
    ax1.plot(pc["U"], pc["T"], "seagreen", lw=2, label="Thrust T (N)")
    ax1.set_xlabel("Wind speed (m/s)")
    ax1.set_ylabel("Thrust (N)", color="seagreen")
    ax1.tick_params(axis="y", labelcolor="seagreen")
    ax1.grid(True, ls="--", alpha=0.4)

    ax2 = ax1.twinx()
    ax2.plot(pc["U"], pc["CT"], "purple", lw=2, ls="--", label="CT")
    ax2.set_ylabel("Thrust coefficient CT", color="purple")
    ax2.tick_params(axis="y", labelcolor="purple")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2)
    fig.suptitle(f"Thrust & CT vs Wind Speed  (lambda = {TIP_SPEED_RATIO})")
    _savefig("power_curve_thrust_ct.png")
    plt.show()


# --- Main --------------------------------------------------------------------

def main():
    sep = "=" * 66
    print(sep)
    print("  Wind Turbine Blade Design - BEM Theory")
    print(sep)

    Omega = TIP_SPEED_RATIO * WIND_SPEED / ROTOR_RADIUS
    nu    = DYN_VISCOSITY / AIR_DENSITY

    # -- Load polars (must be done before any BEM call) ------------------------
    if POLAR_SOURCE == "neuralfoil":
        print(f"\n[Polar] Generating polars via NeuralFoil  "
              f"(airfoil: '{AIRFOIL_NAME}', model: {NF_MODEL_SIZE}) ...")
    else:
        print(f"\n[Polar] Loading polars from folder '{POLAR_DIR}' ...")
    load_polars()

    # -- Auto design point (best Cl/Cd at RE_ASSUMED) -------------------------
    global ALPHA_DESIGN, CL_DESIGN
    if DESIGN_POINT_MODE == "auto" and MODE != "Evaluate":
        # find the polar closest to RE_ASSUMED
        re_vals = [p["re"] for p in _POLARS]
        idx_re  = int(np.argmin(np.abs(np.array(re_vals) - RE_ASSUMED)))
        p_design = _POLARS[idx_re]

        # best Cl/Cd index (restrict to positive Cl to avoid flat-plate regimes)
        with np.errstate(divide='ignore', invalid='ignore'):
            ld = np.where((p_design["cd"] > 0) & (p_design["cl"] > 0),
                          p_design["cl"] / p_design["cd"], 0.0)
        idx_best   = int(np.argmax(ld))
        ALPHA_DESIGN = float(p_design["alpha"][idx_best])
        CL_DESIGN    = float(p_design["cl"][idx_best])
        print(f"\n[Design Point - AUTO]  Using polar at Re={p_design['re']:,.0f}  "
              f"(closest to RE_ASSUMED={RE_ASSUMED:,})")
        print(f"  Best Cl/Cd = {ld[idx_best]:.2f}  "
              f"at alpha = {ALPHA_DESIGN:.2f} deg  ->  "
              f"ALPHA_DESIGN={ALPHA_DESIGN:.2f}  CL_DESIGN={CL_DESIGN:.4f}")
    elif DESIGN_POINT_MODE == "auto" and MODE == "Evaluate":
        print(f"\n[Design Point - AUTO]  Evaluate mode: design point not used, skipping auto.")

    # summary for each loaded polar
    print()
    for p in _POLARS:
        cl_des = float(np.interp(ALPHA_DESIGN, p["alpha"], p["cl"]))
        cd_des = float(np.interp(ALPHA_DESIGN, p["alpha"], p["cd"]))
        with np.errstate(divide='ignore', invalid='ignore'):
            ld = np.where(p["cd"] > 0, p["cl"] / p["cd"], 0.0)
        idx_best = int(np.argmax(ld))
        print(f"  Re = {p['re']:>10,.0f}  |  "
              f"Cl({ALPHA_DESIGN}deg)={cl_des:.4f}  "
              f"Cd({ALPHA_DESIGN}deg)={cd_des:.5f}  "
              f"Cl/Cd={cl_des/cd_des:.1f}  |  "
              f"Best Cl/Cd={ld[idx_best]:.1f} at alpha={p['alpha'][idx_best]:.1f}deg")

    # warn if design Cl from first polar differs from CL_DESIGN (manual mode only)
    if DESIGN_POINT_MODE == "manual":
        cl_des_first = cl_from_alpha(ALPHA_DESIGN)
        if abs(cl_des_first - CL_DESIGN) > 0.05:
            print(f"\n  ** NOTE: CL_DESIGN={CL_DESIGN} but first polar gives "
                  f"Cl={cl_des_first:.4f} at {ALPHA_DESIGN}deg.")
            print(f"     BEM uses the polar value. CL_DESIGN is only for chord sizing.")

    # polar plot
    plot_polar()

    # -- Load blade file (Evaluate mode only) ---------------------------------
    if MODE == "Evaluate":
        if not os.path.exists(BLADE_FILE):
            raise FileNotFoundError(
                f"BLADE_FILE '{BLADE_FILE}' not found. "
                f"Set the correct path or switch MODE to 'Design'.")
        n_stations = load_blade_file(BLADE_FILE)

    inputs = [
        ("Mode",                      MODE,                    ""),
        ("Tip-speed ratio lambda",    TIP_SPEED_RATIO,         ""),
        ("Wind speed Uinf",           WIND_SPEED,              "m/s"),
        ("Rotor radius R",            ROTOR_RADIUS,            "m"),
        ("Hub radius r_hub",          HUB_RADIUS,              "m"),
        ("Number of blades B",        NUM_BLADES,              ""),
        ("Number of elements N",      len(_BLADE_R) if MODE == "Evaluate" else NUM_ELEMENTS, ""),
        ("Design AoA alpha",          ALPHA_DESIGN,            "deg"),
        ("Design Cl",                 CL_DESIGN,               ""),
        ("Air density rho",           AIR_DENSITY,             "kg/m3"),
        ("Dynamic viscosity mu",      DYN_VISCOSITY,           "Pa*s"),
        ("Kinematic viscosity nu",    round(nu, 8),            "m2/s"),
        ("Angular velocity Omega",    round(Omega, 3),         "rad/s"),
        ("Assumed Reynolds number",   RE_ASSUMED,              ""),
        ("Induction correction",      INDUCTION_CORRECTION,    ""),
        ("Hub loss correction",       HUB_LOSS,                ""),
        ("Snel 3D correction",        SNEL_3D_CORRECTION,      ""),
        ("Initial induction factor",  round(1/3, 6),           ""),
    ]
    print("\n[Inputs]")
    for label, val, unit in inputs:
        print(f"  {label:<30}: {val}  {unit}")

    # -- Airfoil --------------------------------------------------------------
    if os.path.exists(AIRFOIL_FILE):
        pts = load_airfoil(AIRFOIL_FILE)
        print(f"\n[Airfoil] {len(pts)} coordinate points loaded from "
              f"'{AIRFOIL_FILE}'")
        # Basic geometry metrics
        xu = pts[pts[:, 1] >= 0, 0]; yu = pts[pts[:, 1] >= 0, 1]
        xl = pts[pts[:, 1] <  0, 0]; yl = pts[pts[:, 1] <  0, 1]
        thickness_max = np.max(yu) - np.min(yl) if len(yu) and len(yl) else 0
        print(f"  Max thickness/chord ~ {thickness_max:.4f}  "
              f"({thickness_max*100:.1f} %)")
        plot_airfoil(pts)
    else:
        print(f"\n[Airfoil] '{AIRFOIL_FILE}' not found - airfoil plot skipped.")

    # -- Step 7: Resolve linear blade shape coefficients ----------------------
    if MODE == "Evaluate":
        print(f"\n[Evaluate Mode] Blade loaded from '{BLADE_FILE}'")
        print(f"  Stations loaded : {len(_BLADE_R)}")
        print(f"  r range         : {_BLADE_R[0]:.4f} m  to  {_BLADE_R[-1]:.4f} m")
        print(f"  chord range     : {_BLADE_CHORD.min():.4f} m  to  {_BLADE_CHORD.max():.4f} m")
        print(f"  twist range     : {_BLADE_TWIST.min():.3f} deg  to  {_BLADE_TWIST.max():.3f} deg")
        print(f"  Eqs. 3.110-3.116 (Steps 6-7) SKIPPED  -- geometry from file")
        print()
        print(f"  {'Station':>8} {'r (m)':>8} {'chord (m)':>10} {'twist (deg)':>12}")
        print(f"  {'-'*42}")
        for i, (r, c, t) in enumerate(zip(_BLADE_R, _BLADE_CHORD, _BLADE_TWIST), 1):
            print(f"  {i:>8} {r:>8.4f} {c:>10.4f} {t:>12.3f}")

    elif BLADE_SHAPE_MODE != "optimum":
        resolve_linear_coefficients(TIP_SPEED_RATIO)
        r_bounds = np.linspace(HUB_RADIUS, ROTOR_RADIUS, NUM_ELEMENTS + 1)
        r_mids   = 0.5 * (r_bounds[:-1] + r_bounds[1:])
        c_root   = linear_chord(HUB_RADIUS)
        c_tip    = linear_chord(ROTOR_RADIUS)
        t_root   = linear_twist(HUB_RADIUS)
        print(f"\n[Step 7 - Linear Blade Shape  (Eqs. 3.115 - 3.116)]")
        print(f"  Mode                    : {BLADE_SHAPE_MODE}")
        print(f"  Chord:  c = a1*r + b1")
        print(f"    a1 = {_A1:+.6f}  m/m")
        print(f"    b1 = {_B1:+.6f}  m")
        print(f"    c at hub  (r={HUB_RADIUS:.3f} m) = {c_root:.4f} m")
        print(f"    c at tip  (r={ROTOR_RADIUS:.3f} m) = {c_tip:.4f} m")
        print(f"  Twist:  theta = a2*(R - r)")
        print(f"    a2 = {_A2:+.6f}  deg/m")
        print(f"    twist at hub = {t_root:.2f} deg")
        print(f"    twist at tip = 0.00 deg  (by definition)")
        # comparison plot: optimum vs linear
        plot_shape_comparison(r_mids)

    # -- Design-point BEM -----------------------------------------------------
    print(f"\n[BEM] Design-point analysis  lambda = {TIP_SPEED_RATIO} ...")
    elements, CP = analyse_blade(TIP_SPEED_RATIO)

    print(f"\n[Results - Design Point  lambda = {TIP_SPEED_RATIO}]")
    print(f"  Power coefficient CP        : {CP:.4f}")
    print(f"  Betz limit                  : {16/27:.4f}")
    print(f"  CP / Betz limit             : {CP/(16/27)*100:.1f} %")
    n_conv = sum(e["converged"] for e in elements)
    print(f"  BEM convergence             : {n_conv}/{len(elements)} elements")

    hdr = (f"{'r(m)':>7} {'c(m)':>7} {'twistdeg':>8} {'AoAdeg':>7} "
           f"{'phideg':>7} {'Cl':>7} {'Cd':>8} {'a':>7} {'aprime':>8} "
           f"{'F':>6} {'f_snel':>7} {'OK':>4}")
    print(f"\n{hdr}")
    print("-" * len(hdr))
    for e in elements:
        ok = "[OK]" if e["converged"] else "[X]"
        print(f"{e['r']:7.3f} {e['chord']:7.4f} {e['twist_deg']:8.2f} "
              f"{e['alpha_deg']:7.2f} {e['phi_deg']:7.2f} "
              f"{e['Cl']:7.4f} {e['Cd']:8.5f} "
              f"{e['a']:7.4f} {e['a_prime']:8.4f} "
              f"{e['F']:6.4f} {e['f_snel']:7.4f} {ok:>4}")

    # -- Reynolds number comparison --------------------------------------------
    re_vals, re_pmin, re_pmax, Re_rec = reynolds_analysis(elements)
    plot_reynolds(elements, re_pmin, re_pmax, Re_rec)

    # -- CP-lambda sweep ------------------------------------------------------------
    print(f"\n[CP-lambda Sweep]  {LAMBDA_STEPS} points, "
          f"lambda = {LAMBDA_MIN} -> {LAMBDA_MAX} ...")
    lams, cps = cp_sweep()
    idx = int(np.argmax(cps))
    print(f"  CP,max = {cps[idx]:.4f}  at  lambda = {lams[idx]:.2f}")
    print(f"  CP at design lambda = {TIP_SPEED_RATIO} :  {CP:.4f}")

    # -- Save -----------------------------------------------------------------
    print("\n[Saving outputs]")
    pc = None
    print(f"\n[Power Curve]  U = {POWER_CURVE_UMIN} -> {POWER_CURVE_UMAX} m/s, "
          f"{POWER_CURVE_STEPS} points, lambda fixed = {TIP_SPEED_RATIO} ...")
    pc = power_curve()
    idx_p = int(np.argmax(pc["P1"]))
    print(f"  P1 max = {pc['P1'][idx_p]/1000:.2f} kW  at  U = {pc['U'][idx_p]:.1f} m/s")
    print(f"  P2 max = {pc['P2'][idx_p]/1000:.2f} kW  at  U = {pc['U'][idx_p]:.1f} m/s")
    print(f"  T  max = {pc['T'][idx_p]:.1f} N        at  U = {pc['U'][idx_p]:.1f} m/s")

    save_xlsx(elements, lams, cps,
              os.path.join(OUTPUT_DIR, "blade_results.xlsx"), pc=pc)

    # -- Plots -----------------------------------------------------------------
    plot_blade(elements)
    plot_cp(lams, cps, CP)
    plot_power_curve_power(pc)
    plot_power_curve_torque_rpm(pc)
    plot_power_curve_thrust_ct(pc)

    # -- HAWC2 c2_def output --------------------------------------------------
    generate_hawc2_c2def(elements, "hawc2_c2def.xlsx")

    print("\nDone.")


if __name__ == "__main__":
    main()