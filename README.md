# North-Sea-Energy-Hub-Optimisation
North Sea - Energy Hub Optimisation (Hydrogen, Wind and Folating Solar)

The aim of the research was to study renewable-based energy hubs in the North Sea region and the creation of an optimisation framework that could serve to simulate early studies and obtain first layouts based on costs, electricity and hydrogen prices subject to a limited area. This was done by defining a hub layout based on offshore wind farms, floating solar farms, hydrogen production, batteries and hydrogen storage. For this, a MILP optimisation model was designed to study the energy hub with six decision variables with an objective function that maximises the NPV. To run simulations in the optimisation model, a scenario was defined based on literature representing future hub developments and their roles in the North Sea, selecting a distance to shore and a representative area within a hub.

- MILP Model.
- Decision Variables:
  x1 - Wind turbines
  x2 - Floating solar installed capacity
  x3 - Battery capacity
  x4 - Electrolyser capacity
  x5 - compressor capacity
  x6 - hydrogen storage capacity
  Objective Function: Maximise NPV. 
  The weather data and input file are not avilable for sharing but someone car replicate the work with own data for each asset. 
