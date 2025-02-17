# Employee Management System in C++

This is a simple Employee Management System built using C++ that stores and manages employee data in an Excel sheet using the **xlnt** library. The system allows adding, deleting, searching, and displaying employee records, as well as calculating the total salary payout.

## Features
- Store employee details in an Excel sheet (`employees.xlsx`)
- Add new employees with ID, name, department, salary, and attendance
- Delete employees by ID
- Search for an employee by ID
- Display all employee records
- Calculate total salary payout

## Installation
To run this project, you need to install the **xlnt** library, which enables interaction with Excel files in C++.

### Installing xlnt
#### Windows (Using vcpkg)
1. Install vcpkg if you haven't already:
   ```sh
   git clone https://github.com/microsoft/vcpkg.git
   cd vcpkg
   ./bootstrap-vcpkg.bat
