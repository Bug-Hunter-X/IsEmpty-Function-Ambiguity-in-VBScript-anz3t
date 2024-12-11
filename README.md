This repository demonstrates a subtle but common error in VBScript related to the `IsEmpty` function. The `IsEmpty` function in VBScript returns `True` for both an empty string ("") and an uninitialized variable.  This ambiguity can cause unexpected behavior if your code relies on `IsEmpty` to differentiate between these two scenarios. The `bug.vbs` file showcases the problematic code, while `bugSolution.vbs` provides a corrected version using the `TypeName` function for more precise variable state checks.