# Kinetic superposition analysis

This directory contains the experimental input data and Python scripts used for the time–temperature and time–pressure kinetic superposition analyses in the associated manuscript.

## Directory structure

```text
kinetic-superposition/
├── README.md
├── requirements.txt
├── time_temperature_superposition.py
├── time_pressure_superposition.py
└── data/
    ├── hill_temperature.xlsx
    └── hill_pressure.xlsx
```

## Experimental input data

### `data/hill_temperature.xlsx`
Experimental time–temperature dataset used by `time_temperature_superposition.py`.

Columns:
- `temperature` — experimental temperature (K)
- `time` — reaction time (s)
- `c_normalized` — structural response variable used in the modified Hill-function fitting

The dataset contains measurements at 1173, 1373, 1573 and 1723 K, with reaction times of 600, 3600, 21600, 43200 and 86400 s at each temperature.

### `data/hill_pressure.xlsx`
Experimental time–pressure dataset used by `time_pressure_superposition.py`.

Columns:
- `pressure` — experimental pressure (GPa)
- `time` — reaction time (s)
- `c_normalized` — structural response variable used in the modified Hill-function fitting

The dataset contains measurements at 0.5, 1, 2 and 3 GPa, with reaction times of 600, 3600, 21600, 43200 and 86400 s at each pressure.

## Scripts

### `time_temperature_superposition.py`
Performs time–temperature superposition using a modified decreasing Hill function. The script uses 1373 K as the reference temperature, fits the reference master curve, determines multiplicative temperature shift factors (`alpha_T`) by minimizing the residual sum of squares, and exports the fitted parameters, shift factors, fit statistics and master-curve figure.

### `time_pressure_superposition.py`
Performs time–pressure superposition using the same modified decreasing Hill function. The script uses 0.5 GPa as the reference pressure, determines multiplicative pressure shift factors (`alpha_P`), and exports the fitted parameters, shift factors, fit statistics and master-curve figure.

## Software requirements

Install the required Python packages with:

```bash
pip install -r requirements.txt
```

Required packages are `pandas`, `numpy`, `matplotlib`, `scipy` and `openpyxl`.

## Running the analyses

Run both scripts from the `kinetic-superposition` directory:

```bash
python time_temperature_superposition.py
python time_pressure_superposition.py
```

The scripts read their corresponding Excel files automatically from the `data/` subdirectory.

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

## Reproducibility

The two scripts and the accompanying experimental input datasets were tested together after the data files were standardized. Both scripts run successfully and generate the expected output tables and figures.
