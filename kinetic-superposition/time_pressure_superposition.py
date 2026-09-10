import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit, minimize
# === Read raw data ===
df = pd.read_excel("data/hill_pressure.xlsx")
df.columns = df.columns.str.strip()
df = df[df['time'] > 0]
# === Set reference pressure ===
P_ref = 0.5
df_ref = df[df['pressure'] == P_ref].sort_values('time')
t_ref = df_ref['time'].values
c_ref = df_ref['c_normalized'].values  # actually the original unit cell height
# === Modified Hill equation (decreasing form) ===
def hill_func(t, Cmax, Cmin, A, h):
    return Cmax + (Cmin - Cmax) / (1 + (A / t)**h)
# === Fit the master curve ===
p0 = [max(c_ref), min(c_ref), np.median(t_ref), 1.0]
bounds = ([max(c_ref)-0.2, min(c_ref)-0.2, 1, 0.1], [max(c_ref)+0.2, min(c_ref)+0.2, 1e8, 5])
popt_master, _ = curve_fit(hill_func, t_ref, c_ref, p0, bounds=bounds, maxfev=20000)
# === Compute ln(alpha_P) and alpha_P (multiplicative shift factors) ===
ln_alpha_P = {}
pressures = sorted(df['pressure'].unique())
pressures = [p for p in pressures if p != P_ref]
for P in pressures:
    df_p = df[df['pressure'] == P].sort_values('time')
    t = df_p['time'].values
    c = df_p['c_normalized'].values
    def shift_loss(ln_alpha):
        t_shifted = t * np.exp(ln_alpha)
        c_pred = hill_func(t_shifted, *popt_master)
        return np.sum((c - c_pred)**2)
    res = minimize(shift_loss, x0=[0.0])
    ln_alpha_P[P] = res.x[0]
# === Output alpha_P table ===
df_alpha = pd.DataFrame({
    "Pressure (GPa)": list(ln_alpha_P.keys()),
    "ln(alpha_P)": list(ln_alpha_P.values()),
    "alpha_P": [np.exp(v) for v in ln_alpha_P.values()]
}).sort_values("Pressure (GPa)")
df_alpha.to_excel("pressure_shift_factors_multiply_corrected.xlsx", index=False)
print("✓ alpha_P table saved as: pressure_shift_factors_multiply_corrected.xlsx")
# === Plot master curve and all shifted data ===
t_plot = np.logspace(np.log10(t_ref.min()), np.log10(t_ref.max()*10), 500)
c_plot = hill_func(t_plot, *popt_master)
plt.figure(figsize=(10, 6))
plt.semilogx(t_plot, c_plot, 'k-', label="Master Curve (0.5 GPa)")
plt.semilogx(t_ref, c_ref, 'ko', label="0.5 GPa Data")
for P in ln_alpha_P:
    df_p = df[df['pressure'] == P]
    t_shifted = df_p['time'].values * np.exp(ln_alpha_P[P])
    c_vals = df_p['c_normalized'].values
    plt.semilogx(t_shifted, c_vals, 'o', label=f"{P} GPa (shifted)")
plt.xlabel("Time × alpha_P (s, log scale)")
plt.ylabel("Unit Cell Height c (Å)")
plt.title("Master Curve via Hill Fit (Corrected: Decreasing Form)")
plt.grid(True, which='both', linestyle='--', alpha=0.5)
plt.legend()
plt.tight_layout()
plt.savefig("hill_master_curve_multiply_corrected.png", dpi=300)
plt.show()
# === Export Hill fit parameters ===
hill_params = {
    "C_max (initial)": [popt_master[0]],
    "C_min (final)": [popt_master[1]],
    "A":     [popt_master[2]],
    "Hill_coefficient_h": [popt_master[3]]
}
df_hill_params = pd.DataFrame(hill_params)
df_hill_params.to_excel("hill_fit_parameters_corrected.xlsx", index=False)
print("✓ Hill fit parameters saved as: hill_fit_parameters_corrected.xlsx")
# === Fit error analysis ===
c_pred = hill_func(t_ref, *popt_master)
residuals = c_ref - c_pred
ss_res = np.sum(residuals**2)
ss_tot = np.sum((c_ref - np.mean(c_ref))**2)
r_squared = 1 - (ss_res / ss_tot)
rmse = np.sqrt(np.mean(residuals**2))
fit_stats = pd.DataFrame({
    "R_squared": [r_squared],
    "RMSE": [rmse],
    "RSS": [ss_res]
})
fit_stats.to_excel("hill_fit_statistics_corrected.xlsx", index=False)
print("✓ Hill fit statistics saved as: hill_fit_statistics_corrected.xlsx")
