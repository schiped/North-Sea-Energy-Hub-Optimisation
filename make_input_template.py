"""Write input/input_file_template.xlsx with every sheet and parameter the model reads.

The original input file (weather series and cost assumptions from the 2022 TNO study)
is not public. Fill the template with your own data and save it as input/input_file.xlsx.
"""
from pathlib import Path

import pandas as pd

PARAMS = {
    "General": [("area_hub", "m2"), ("discount_rate", "-"), ("shore_distance", "km"), ("system_lifetime", "years")],
    "Economic": [
        ("CAPEX_wind", "EUR/turbine"), ("OPEX_wind", "EUR/turbine/yr"), ("CAPEX_FPV", "EUR/m2"), ("OPEX_solar_total", "EUR/m2/yr"),
        ("CAPEX_battery", "EUR/kWh"), ("OPEX_battery", "EUR/kWh/yr"), ("CAPEX_Electrolysis", "EUR/kW"), ("OPEX_Electrolysis", "EUR/kW/yr"),
        ("CAPEX_compressor", "EUR/kW"), ("OPEX_compressor", "EUR/kW/yr"), ("CAPEX_H2_storage", "EUR/kg"), ("OPEX_H2_storage", "EUR/kg/yr"),
    ],
    "Solar": [("FPV_kWpm2", "kWp/m2"), ("area_FPV", "m2")],
    "Wind": [("P_turbine", "kW"), ("required_area_turbine", "m2")],
    "Storage": [("charge_rate", "efficiency"), ("discharge_rate", "efficiency"), ("charge_discharge_power", "C-rate"), ("Self_discharge", "-/h")],
    "Hydrogen": [
        ("Electrolysis_Efficiency", "kWh/kg"), ("CAPEX_stack", "EUR/kW"), ("lifetime_stack", "h"), ("output_pressure", "bar"),
        ("export_pressure", "bar"), ("pipe_capacity", "kg/h"), ("CAPEX_pipe", "EUR/km"), ("OPEX_pipe", "EUR/km/yr"),
    ],
}

out = Path(__file__).parent / "input" / "input_file_template.xlsx"
out.parent.mkdir(exist_ok=True)
hours = range(1, 8761)
with pd.ExcelWriter(out) as xw:
    for sheet, rows in PARAMS.items():
        pd.DataFrame([(n, None, u) for n, u in rows], columns=["Parameter", "Input", "Unit"]).to_excel(xw, sheet_name=sheet, index=False)
    pd.DataFrame({"Hour": hours, "Wind_Power_1": None, "Unit: kWh per turbine": None}).to_excel(xw, sheet_name="Wind_Power_Data", index=False)
    pd.DataFrame({"Hour": hours, "Power_Output[kWh/m2]": None, "Unit: kWh per m2 of FPV": None}).to_excel(xw, sheet_name="Solar_Power_Data", index=False)
    pd.DataFrame({"Hour": hours, "Average_2015_2022_[EUR/kWh]": None}).to_excel(xw, sheet_name="Day-ahead Prices_2015-2022", index=False)
print(f"written {out}")
