import argparse
import os
import numpy as np
import pandas as pd


def spm_teng25(rrs_555, rrs_670, rrs_810):
    """
    Teng25 SPM retrieval with water-type classification.

    Implements:
      1) POC/SPM ratio proxy from Rrs_555:
         ratio = 10 ** (a * log10(Rrs_555) + b),  a=-0.973, b=-3.213
      2) Water-type rules and class-specific power-law SPM models:
         - org_rich:  ratio > 0.12 OR Rrs_670 < 0.01
         - min_rich:  0.02 < ratio < 0.12
         - min_rich(E): ratio < 0.02
    """
    rrs_555 = np.asarray(rrs_555, dtype=float)
    rrs_670 = np.asarray(rrs_670, dtype=float)
    rrs_810 = np.asarray(rrs_810, dtype=float)

    a, b = -0.973, -3.213
    with np.errstate(divide="ignore", invalid="ignore"):
        poc_spm_ratio = 10 ** (a * np.log10(rrs_555) + b)

    water_types = np.full(rrs_555.shape, "undefined", dtype=object)
    spm_conc = np.full(rrs_555.shape, np.nan, dtype=float)

    org_rich = (poc_spm_ratio > 0.12) | (rrs_670 < 0.01)
    min_rich = (poc_spm_ratio > 0.02) & (poc_spm_ratio < 0.12)
    min_rich_e = (poc_spm_ratio < 0.02)

    water_types[org_rich] = "org_rich"
    spm_conc[org_rich] = 1992.2 * (rrs_670[org_rich] ** 1.027)

    water_types[min_rich] = "min_rich"
    spm_conc[min_rich] = 12662.7 * (rrs_810[min_rich] ** 1.157)

    water_types[min_rich_e] = "min_rich(E)"
    spm_conc[min_rich_e] = 50556.7 * (rrs_810[min_rich_e] ** 1.371)

    return poc_spm_ratio, water_types, spm_conc


def process_csv(input_csv: str, output_csv: str) -> None:
    """Run Teng25 on a CSV and write results to a new CSV."""
    df = pd.read_csv(input_csv)

    required = ["Rrs_555", "Rrs_670", "Rrs_810"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing} (required: {required})")

    ratio, wtype, spm = spm_teng25(df["Rrs_555"], df["Rrs_670"], df["Rrs_810"])

    df["poc_spm_ratio"] = ratio
    df["water_types"] = wtype
    df["spm_conc"] = spm

    os.makedirs(os.path.dirname(output_csv) or ".", exist_ok=True)
    df.to_csv(output_csv, index=False)
    print(f"Processed {len(df)} records -> saved to {output_csv}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Teng25 SPM retrieval (requires Rrs_555, Rrs_670, Rrs_810)."
    )
    parser.add_argument("-i", "--input", required=True, help="Input CSV path.")
    parser.add_argument("-o", "--output", required=True, help="Output CSV path.")
    args = parser.parse_args()

    process_csv(args.input, args.output)


if __name__ == "__main__":
    main()
