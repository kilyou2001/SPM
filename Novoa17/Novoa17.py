import argparse
import os
import numpy as np
import pandas as pd

# ------------------------------------------------------------
# Novoa17 implementation
#
# Workflow :
#   1) Convert Rrs -> rho_w using: rho_w(λ) = π * Rrs(λ)
#   2) Use rho_w(655) to select a spectral regime:
#        Green       : rho655 < 0.007
#        Green-Red   : 0.007 ≤ rho655 < 0.016
#        Red         : 0.016 ≤ rho655 < 0.08
#        Red-NIR     : 0.08  ≤ rho655 < 0.12 
#        NIR         : rho655 ≥ 0.12
#   3) Apply the regime-specific SPM equation.
# ------------------------------------------------------------

# Switching points
S_G  = 0.007
S_GR = 0.016
S_R  = 0.08
S_RN = 0.12


def _ln_weights(x, s_minus, s_plus):
    """Return (alpha, beta) for ln-weighted transition between (s_minus, s_plus)."""
    denom = np.log(s_plus / s_minus)
    alpha = np.log(s_plus / x) / denom
    beta  = np.log(x / s_minus) / denom
    return alpha, beta


def spm_novoa17(rrs_561, rrs_655, rrs_865):
  
    rrs_561 = np.asarray(rrs_561, dtype=float)
    rrs_655 = np.asarray(rrs_655, dtype=float)
    rrs_865 = np.asarray(rrs_865, dtype=float)

    rho561 = np.pi * rrs_561
    rho655 = np.pi * rrs_655
    rho865 = np.pi * rrs_865

    # Validity: finite inputs and non-negative rho655 
    valid = np.isfinite(rho561) & np.isfinite(rho655) & np.isfinite(rho865) & (rho655 >= 0)

    estimate = np.full(rho655.shape, np.nan, dtype=float)
    types = np.full(rho655.shape, np.nan, dtype=object)

    # Equations
    spm_green = 130.1 * rho561
    spm_red   = 531.5 * rho655
    spm_nir   = 37150.0 * (rho865 ** 2) + 1751.0 * rho865

    m_green = valid & (rho655 < S_G)
    m_gr    = valid & (rho655 >= S_G)  & (rho655 < S_GR)
    m_red   = valid & (rho655 >= S_GR) & (rho655 < S_R)
    m_rn    = valid & (rho655 >= S_R)  & (rho655 < S_RN)
    m_nir   = valid & (rho655 >= S_RN)

    # Green
    estimate[m_green] = spm_green[m_green]
    types[m_green] = "Green"

    # Green-Red (blend Green + Red)
    if np.any(m_gr):
        x = rho655[m_gr]
        with np.errstate(divide="ignore", invalid="ignore"):
            a, b = _ln_weights(x, S_G, S_GR)
        estimate[m_gr] = a * spm_green[m_gr] + b * spm_red[m_gr]
        types[m_gr] = "Green-Red"

    # Red
    estimate[m_red] = spm_red[m_red]
    types[m_red] = "Red"

    # Red-NIR (blend Red + NIR)
    if np.any(m_rn):
        x = rho655[m_rn]
        with np.errstate(divide="ignore", invalid="ignore"):
            a, b = _ln_weights(x, S_R, S_RN)
        estimate[m_rn] = a * spm_red[m_rn] + b * spm_nir[m_rn]
        types[m_rn] = "Red-NIR"

    # NIR
    estimate[m_nir] = spm_nir[m_nir]
    types[m_nir] = "NIR"

    return estimate, types


def process_csv(input_csv: str, output_csv: str) -> None:
    """Read CSV, run Novoa17, append (Estimate, Type), and write CSV."""
    df = pd.read_csv(input_csv)

    required = ["Rrs_561", "Rrs_655", "Rrs_865"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing} (required: {required})")

    est, typ = spm_novoa17(df["Rrs_561"], df["Rrs_655"], df["Rrs_865"])
    df["Estimate"] = est
    df["Type"] = typ

    os.makedirs(os.path.dirname(output_csv) or ".", exist_ok=True)
    df.to_csv(output_csv, index=False, na_rep="NaN")
    print(f"Processed {len(df)} records -> saved to {output_csv}")


def main() -> None:
    p = argparse.ArgumentParser(description="Novoa17 SPM retrieval (requires Rrs_561, Rrs_655, Rrs_865).")
    p.add_argument("-i", "--input", required=True, help="Input CSV path.")
    p.add_argument("-o", "--output", required=True, help="Output CSV path.")
    args = p.parse_args()
    process_csv(args.input, args.output)


if __name__ == "__main__":
    main()
