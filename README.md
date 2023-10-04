# Effects-of-Pressure-and-Temperature-on-Isentropic-Efficiency-Excel-VBA
Thermodynamics Data Analysis using Excel VBA to Visualize Data

This Excel VBA project calculates the isentropic efficiency of a gas turbine using specified operating conditions and thermodynamic properties.

## Background

The calculations are based on the main article provided as Reference 1, and assume the following conditions:

- The gas behaves as an ideal gas throughout the process.
- The gas composition is assumed to be dry air.
- The process is considered isentropic.
- Specific heat ratio (k) and specific heat capacity (cp) values are based on dry air properties at the specified conditions in Reference 1.
- Mass flow rate (m), inlet temperature (T1r), outlet temperature (T2r), inlet pressure (P1r), outlet pressure (P2r), isentropic efficiency (Ntr), and turbine axial work (W) values are in accordance with Reference 1.
- The calculations assume steady-state conditions.

## Units

- `Ratio`: Expansion Ratio
- `Nt`: Isentropic Efficiency
- `k`: Specific heat ratio for dry air at 1400 K
- `cp`: Specific heat for dry air at 1400 K (MJ/kgK)
- `m`: Mass Flow Rate (kg/s)
- `T1r`: Inlet Temp (K)
- `T2r`: Outlet Temp (K)
- `P1r`: Inlet Pressure (MPa)
- `P2r`: Outlet Pressure (MPa)
- `Ntr`: Isentropic Efficiency
- `W`: Turbine Axial Work (MW)
- `RatioR`: Expansion Ratio
- `Ws`: Isentropic Turbine Work (MW)
- `T2s`: Ideal Temperature Output (K)
- `RatioP2s`: P2/P1

## Usage

1. Open the Excel workbook containing this VBA project.

2. Enter the values for Inlet Temperature and Expansion Ratio

3. Run the `CalculateIsentropicEfficiency` macro to calculate and display the isentropic efficiency.

4. Run the `VisualizeData` macro to create a scatter plot visualization of the data.

## Functions List
(1) CelciusToKelvin
    Converting Celcius Inlet Temperatures to Kelvin by adding 273.15
(2) SolveT2s
    Solving for Outlet Temperatures using Ideal Gas Law 
    Equation used: T2s = (T1)(P2s/P1)^(k-1/k)
(3) SolveWs
    Solving for Isentropic Turbine Work 
    Equation Used: Ws = m*cp*(T1/T2s)
(4) SolveIsentropicEfficiency 
    Solving for Isentopic Efficiency
    Equation Used: Nt = W / Ws

## Error Handling

This project includes error handling for the following scenarios:

- Division by zero when calculating `Ws` or `Nt`.
- Warning if `T2s` is greater than `T1`.
- Warning if `Nt` is outside the expected range of 0 to 1.

Please double-check your data and units if you encounter any warnings.

## References

1. [Main Article - Gas Turbine Efficiency Calculation](https://www.sciencedirect.com/science/article/pii/S0360544221009142)
2. [Thermodynamic Properties of Dry Air](https://www.engineeringtoolbox.com/dry-air-properties-d_973.html)
3. [FUNDAMENTALS OF GAS TURBINE ENGINES](https://www.cast-safety.org/pdf/3_engine_fundamentals.pdf)
4. [First Law of Thermodynamics for an Open System](https://www.isisvarese.edu.it/wp-content/uploads/2016/03/first-law-of-thermodynamics-for-an-open-system-.pdf)

Feel free to refer to these references for more information on the calculations and assumptions used in this project.

**Note:** This code is intended for educational and demonstration purposes - it is mostly to test Excel VBA skills rather than Thermodynamic Analysis skills. Please use it responsibly and verify the results for your specific application.
