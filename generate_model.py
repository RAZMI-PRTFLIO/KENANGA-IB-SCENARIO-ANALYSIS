import pandas as pd

# Base assumptions
base_revenue_2024 = 1000  # RM million
base_pbt_2024 = 117.2
years = list(range(2025, 2030))

scenarios = {
    "Bull Case": {"revenue_growth": 0.10, "pbt_margin": 0.17, "cost_to_income": [85, 80, 78, 75, 72]},
    "Base Case": {"revenue_growth": 0.05, "pbt_margin": 0.12, "cost_to_income": [85, 85, 84, 83, 82]},
    "Bear Case": {"revenue_growth": -0.02, "pbt_margin": 0.07, "cost_to_income": [88, 90, 92, 94, 95]},
}

# Calculate projections
results = {}
for name, params in scenarios.items():
    revenue = [base_revenue_2024]
    pbt = [base_pbt_2024]
    for i in range(5):
        revenue.append(revenue[-1] * (1 + params["revenue_growth"]))
        pbt.append(revenue[-1] * params["pbt_margin"])
    df = pd.DataFrame({
        "Year": years,
        "Revenue (RM mil)": revenue[1:],
        "PBT (RM mil)": pbt[1:],
        "Cost-to-Income Ratio (%)": params["cost_to_income"]
    })
    results[name] = df

# Export to Excel
with pd.ExcelWriter("Kenanga_Scenario_Analysis_2025_2029.xlsx") as writer:
    for name, df in results.items():
        df.to_excel(writer, sheet_name=name, index=False)

print("Excel file created successfully.")