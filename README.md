## Overview
This project uses Excel VBA and Excel Solver to solve a transportation optimization problem between multiple plants and cities. The model determines the optimal shipment quantities needed to minimize total transportation cost while satisfying supply and demand constraints. The workbook also supports automated scenario analysis and custom user-generated scenarios through a VBA UserForm.

---

## Features
- Transportation cost minimization using Excel Solver
- Automated scenario analysis with VBA `For` loops
- Dynamic shipment and cost reporting
- VBA UserForm for custom scenario creation
- Input validation to ensure total supply is greater than or equal to total demand
- Named ranges for cleaner Solver integration
- Performance optimization using:
  - `Application.ScreenUpdating`
  - Manual calculation mode

---


### Results
Displays:
- Optimal shipment plans
- Total transportation cost for each scenario
- Custom scenario results generated from the UserForm


---

## How to Run
1. Open the `.xlsm` workbook
2. Enable macros
3. Ensure the Excel Solver Add-in is enabled
4. Run the `Scenarios()` macro to solve all predefined scenarios
5. Use the "Create your own scenario!" button in the results worksheet to generate custom scenarios

---

## Notes
- Excel calculation mode should be set to `Automatic` if Solver does not update correctly.
- VBA `Val()` was used to safely convert UserForm textbox inputs into numeric values before optimization.
