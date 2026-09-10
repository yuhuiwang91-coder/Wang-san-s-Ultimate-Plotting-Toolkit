# Kinetic superposition analysis

This directory contains the two Python scripts used for the kinetic superposition analyses reported in the associated manuscript.

## Scripts

### `time_temperature_superposition.py`
Performs time–temperature superposition using a modified decreasing Hill function. The script uses 1373 K as the reference temperature, fits the reference master curve, determines multiplicative temperature shift factors (`alpha_T`) by minimizing the residual sum of squares, and exports the fitted parameters, shift factors, fit statistics, and the master-curve figure.

### `time_pressure_superposition.py`
Performs time–pressure superposition using the same modified decreasing Hill function. The script uses 0.5 GPa as the reference pressure, determines multiplicative pressure shift factors (`alpha_P`), and exports the fitted parameters, shift factors, fit statistics, and the master-curve figure.

## Input data

Both scripts read an Excel workbook named:

`hill.xlsx`

The workbook must contain at least the following columns:

- `time` — reaction time (s); rows with `time <= 0` are excluded
- `temperature` — experimental temperature (K)
- `pressure` — experimental pressure (GPa)
- `c_normalized` — unit-cell height / structural response variable used in the Hill-function fitting

Place `hill.xlsx` in the same working directory as the scripts before running them.

The input dataset used for the published calculations should be supplied with the article's Source Data / Supplementary Data or deposited in the associated data repository so that the analyses can be reproduced.

## Required Python packages

Install the required packages with:

```bash
pip install -r requirements.txt
```

The scripts require `pandas`, `numpy`, `matplotlib`, `scipy`, and `openpyxl` (for Excel input/output).

## Running the analyses

From this directory, run:

```bash
python time_temperature_superposition.py
python time_pressure_superposition.py
```

## Main outputs

The temperature-superposition script produces:

- `temperature_shift_factors_multiply_corrected.xlsx`
- `hill_master_curve_temperature_corrected.png`
- `hill_fit_parameters_temperature_corrected.xlsx`
- `hill_fit_statistics_temperature_corrected.xlsx`

The pressure-superposition script produces:

- `pressure_shift_factors_multiply_corrected.xlsx`
- `hill_master_curve_multiply_corrected.png`
- `hill_fit_parameters_corrected.xlsx`
- `hill_fit_statistics_corrected.xlsx`

## Notes

The scripts implement the analysis exactly as used for the associated manuscript, including a reference temperature of 1373 K and a reference pressure of 0.5 GPa. The repository version should be archived at publication in a DOI-minting repository (for example, Zenodo or Code Ocean) to provide a permanent, citable record of the code version used in the paper.
