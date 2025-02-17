#include <iostream>
#include <xlnt/xlnt.hpp>

using namespace std;
 
const string FILE_NAME = "employees.xlsx";

 
void initializeExcel() {
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();

    // Set headers
    ws.cell("A1").value("ID");
    ws.cell("B1").value("Name");
    ws.cell("C1").value("Department");
    ws.cell("D1").value("Salary");
    ws.cell("E1").value("Attendance");

    wb.save(FILE_NAME);
}

// Function to add an employee to the Excel file
void addEmployee() {
    xlnt::workbook wb;
    wb.load(FILE_NAME);
    xlnt::worksheet ws = wb.active_sheet();

    int id;
    string name, department, attendance;
    double salary;

    cout << "Enter Employee ID: "; cin >> id;
    cout << "Enter Name: "; cin >> name;
    cout << "Enter Department: "; cin >> department;
    cout << "Enter Salary: "; cin >> salary;
    cout << "Enter Attendance (Present/Absent): "; cin >> attendance;

    int row = 2;
    while (ws.cell("A" + to_string(row)).has_value()) {
        row++;
    }

    ws.cell("A" + to_string(row)).value(id);
    ws.cell("B" + to_string(row)).value(name);
    ws.cell("C" + to_string(row)).value(department);
    ws.cell("D" + to_string(row)).value(salary);
    ws.cell("E" + to_string(row)).value(attendance);

    wb.save(FILE_NAME);
    cout << "Employee added successfully!\n";
}


