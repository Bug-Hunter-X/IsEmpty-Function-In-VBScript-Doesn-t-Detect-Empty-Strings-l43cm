# VBScript IsEmpty Function Bug

This repository demonstrates a subtle bug in VBScript's `IsEmpty` function.  The `IsEmpty` function is intended to check if a variable is uninitialized or empty. However, it does not correctly identify empty strings as empty, potentially leading to unexpected errors in your scripts.  The `bug.vbs` file showcases this issue. The solution provided in `bugSolution.vbs` demonstrates how to correctly handle this using the `Len` function to check for empty strings and custom error handling.

## Bug Description

The `IsEmpty` function is designed to check if a variable is uninitialized or empty. However, when used with empty strings (""), it returns `False`, which means it does not consider them empty. This can lead to incorrect conditional logic and unexpected crashes, especially when handling functions with optional parameters that may be passed as empty strings. 

## Solution

The solution uses the `Len` function to reliably detect empty strings. The `Len` function returns the length of a string, so a length of 0 indicates an empty string. The updated code incorporates more robust error handling, using a custom error number for better debugging.