import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit, minimize
# === Read raw data ===
df = pd.read_excel("hill.xlsx")
df.columns = df.columns.str.strip()
df = df[df['time'] > 0]
# === Set reference temperature ===
T_ref = 1373  # K
df_ref = df[df['temperature'] == T_ref].sort_values('time')
t_ref = df_ref['time'].values
c_ref = df_ref['c_normalized'].values  # original unit cell height
# === Modified Hill equation (decreasing form) ===
def hill_func(t, Cmax, Cmin, A, h):
    return Cmax + (Cmin - Cmax) / (1 + (A / t)**h)
# === Fit the master curve ===
p0 = [max(c_ref), min(c_ref), np.median(t_ref), 1.0]
bounds = ([max(c_ref)-0.2, min(c_ref)-0.2, 1, 0.1], 
          [max(c_ref)+0.2, min(c_ref)+0.2, 1e8, 5])
popt_master, _ = curve_fit(hill_func, t_ref, c_ref, p0, bounds=bounds, maxfev=20000)
# === Compute ln(alpha_T) and alpha_T (multiplicative shift factors) ===
ln_alpha_T = {}
temps = sorted(df['temperature'].unique())
temps = [T for T in temps if T != T_ref]
for T in temps:
    df_T = df[df['temperature'] == T].sort_values('time')
    t = df_T['time'].values
    c = df_T['c_normalized'].values
    def shift_loss(ln_alpha):
        t_shifted = t * np.exp(ln_alpha)
        c_pred = hill_func(t_shifted, *popt_master)
        return np.sum((c - c_pred)**2)
    res = minimize(shift_loss, x0=[0.0])
    ln_alpha_T[T] = res.x[0]
# === Output alpha_T table ===
df_alpha = pd.DataFrame({
    "Temperature (K)": list(ln_alpha_T.keys()),
    "ln(alpha_T)": list(ln_alpha_T.values()),
    "alpha_T": [np.exp(v) for v in ln_alpha_T.values()]
}).sort_values("Temperature (K)")
df_alpha.to_excel("temperature_shift_factors_multiply_corrected.xlsx", index=False)
print("✓ alpha_T table saved as: temperature_shift_factors_multiply_corrected.xlsx")
# === Plot master curve and all shifted data ===
t_plot = np.logspace(np.log10(t_ref.min()), np.log10(t_ref.max()*10), 500)
c_plot = hill_func(t_plot, *popt_master)

plt.figure(figsize=(10, 6))
plt.semilogx(t_plot, c_plot, 'k-', label="Master Curve (1373 K)")
plt.semilogx(t_ref, c_ref, 'ko', label="1373 K Data")
for T in ln_alpha_T:
    df_T = df[df['temperature'] == T]
    t_shifted = df_T['time'].values * np.exp(ln_alpha_T[T])
    c_vals = df_T['c_normalized'].values
    plt.semilogx(t_shifted, c_vals, 'o', label=f"{T} K (shifted)")
plt.xlabel("Time × alpha_T (s, log scale)")
plt.ylabel("Unit Cell Height c (Å)")
plt.title("Master Curve via Hill Fit (Temperature Superposition)")
plt.grid(True, which='both', linestyle='--', alpha=0.5)
plt.legend()
plt.tight_layout()
plt.savefig("hill_master_curve_temperature_corrected.png", dpi=300)
plt.show()
# === Export Hill fit parameters ===
hill_params = {
    "C_max (initial)": [popt_master[0]],
    "C_min (final)": [popt_master[1]],
    "A":     [popt_master[2]],
    "Hill_coefficient_h": [popt_master[3]]
}
df_hill_params = pd.DataFrame(hill_params)
df_hill_params.to_excel("hill_fit_parameters_temperature_corrected.xlsx", index=False)
print("✓ Hill fit parameters saved as: hill_fit_parameters_temperature_corrected.xlsx")

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
fit_stats.to_excel("hill_fit_statistics_temperature_corrected.xlsx", index=False)
print("✓ Hill fit statistics saved as: hill_fit_statistics_temperature_corrected.xlsx")
