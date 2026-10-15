# North Sea energy hub optimisation (wind, floating PV, batteries, hydrogen)

A mixed-integer optimisation model (Pyomo) that finds the first layout of an offshore energy hub in the North Sea. It maximises NPV subject to a limited sea area, with the hub selling electricity through an export cable and hydrogen through a pipeline. Developed at **TNO** (2022) during my European Master in Sustainable Energy System Management (Best Research Project Award).

> **ES:** Modelo MILP en Pyomo para dimensionar un hub energético offshore en el Mar del Norte: eólica marina, solar flotante, baterías, electrolizadores, compresión y almacenamiento de hidrógeno, maximizando el VAN con restricción de área. Desarrollado en TNO (2022).

![image](https://github.com/user-attachments/assets/b73dbc09-7d8a-47bd-97e0-645e658aedca)

Background on the hubs studied: [North Sea Energy atlas](https://north-sea-energy.eu/en/energy-atlas/).

## Formulation

**Decision variables (sizing)**

| | Variable | Type |
|---|---|---|
| x1 | Floating PV area (m²) | continuous |
| x2 | Number of wind turbines | integer |
| x3 | Battery capacity (kWh) | continuous |
| x4 | Electrolyser capacity (kW) | continuous |
| x5, x5b | Compressor power: pipeline export, and 200 bar storage (kW) | continuous |
| x6 | Hydrogen buffer storage (kg) | continuous |

**Hourly flows (8,760 h)**
- Wind and PV generation is split between grid export, battery charging and hydrogen production. Curtailment is explicit.
- The battery discharges to the grid, the electrolyser or the compressors.
- Electrolyser output goes to the pipeline (via compression) or to buffer storage. The buffer then feeds the pipeline.

**Constraints**
- Electricity and hydrogen balances.
- Battery and hydrogen storage dynamics.
- Electrolyser and compressor capacity.
- Compressor power as a function of pressure ratio (30 bar → export pressure, and → 200 bar).
- Pipeline capacity.
- **Total area of turbines plus floating PV ≤ hub area.**

**Objective.** Maximise NPV over the system lifetime: electricity revenue at day-ahead prices (2015–2022 average profile) plus hydrogen revenue, minus capex and the present value of opex for every asset, the hydrogen pipeline and electrolyser stack replacements.

## Run it

```bash
pip install pyomo highspy pandas numpy numpy-financial openpyxl matplotlib
python make_input_template.py            # writes input/input_file_template.xlsx
# fill the template with your data, save it as input/input_file.xlsx, then:
python Energy_Hub_Optimization_REV3.py   # set SOLVER=gurobi to use Gurobi
```

The original weather series and cost inputs came from the TNO study and are not public. `make_input_template.py` lists every sheet, parameter and unit the model reads, so you can replace them with your own data. Wind can come from renewables.ninja, PV from PVGIS, and prices from ENTSO-E or Energy-Charts.

## Related work

- [argentina-bess-valuation](https://github.com/schiped/argentina-bess-valuation): battery dispatch and project finance for Argentina's AlmaSADI tender.
- [hybrid-pv-bess-mater](https://github.com/schiped/hybrid-pv-bess-mater): PV + BESS sizing behind a grid export limit.
- [feeder-hosting-capacity-bess](https://github.com/schiped/feeder-hosting-capacity-bess): distribution planning with LinDistFlow.
