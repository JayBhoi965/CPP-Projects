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


 
void deleteEmployee() {
    xlnt::workbook wb;
    wb.load(FILE_NAME);
    xlnt::worksheet ws = wb.active_sheet();

    int id;
    cout << "Enter Employee ID to delete: ";
    cin >> id;

    bool found = false;
    int row = 2;
    while (ws.cell("A" + to_string(row)).has_value()) {
        if (ws.cell("A" + to_string(row)).value<int>() == id) {
            found = true;
            break;
        }
        row++;
    }

    if (found) {
        int lastRow = row;
        while (ws.cell("A" + to_string(lastRow + 1)).has_value()) {
            ws.cell("A" + to_string(lastRow)).value(ws.cell("A" + to_string(lastRow + 1)).value<int>());
            ws.cell("B" + to_string(lastRow)).value(ws.cell("B" + to_string(lastRow + 1)).value<string>());
            ws.cell("C" + to_string(lastRow)).value(ws.cell("C" + to_string(lastRow + 1)).value<string>());
            ws.cell("D" + to_string(lastRow)).value(ws.cell("D" + to_string(lastRow + 1)).value<double>());
            ws.cell("E" + to_string(lastRow)).value(ws.cell("E" + to_string(lastRow + 1)).value<string>());
            lastRow++;
        }
        ws.cell("A" + to_string(lastRow)).clear();
        ws.cell("B" + to_string(lastRow)).clear();
        ws.cell("C" + to_string(lastRow)).clear();
        ws.cell("D" + to_string(lastRow)).clear();
        ws.cell("E" + to_string(lastRow)).clear();

        wb.save(FILE_NAME);
        cout << "Employee deleted successfully!\n";
    } else {
        cout << "Employee not found!\n";
    }
}

// display employees from Excel
void displayEmployees() {
    xlnt::workbook wb;
    wb.load(FILE_NAME);
    xlnt::worksheet ws = wb.active_sheet();

    cout << "\n----- Employee List -----\n";
    cout << "ID\tName\tDepartment\tSalary\tAttendance\n";

    int row = 2;
    while (ws.cell("A" + to_string(row)).has_value()) {
        cout << ws.cell("A" + to_string(row)).value<int>() << "\t"
             << ws.cell("B" + to_string(row)).value<string>() << "\t"
             << ws.cell("C" + to_string(row)).value<string>() << "\t"
             << ws.cell("D" + to_string(row)).value<double>() << "\t"
             << ws.cell("E" + to_string(row)).value<string>() << endl;
        row++;
    }
}

//Search an employee by ID
void searchEmployee() {
    xlnt::workbook wb;
    wb.load(FILE_NAME);
    xlnt::worksheet ws = wb.active_sheet();

    int id;
    cout << "Enter Employee ID to search: ";
    cin >> id;

    int row = 2;
    while (ws.cell("A" + to_string(row)).has_value()) {
        if (ws.cell("A" + to_string(row)).value<int>() == id) {
            cout << "ID: " << ws.cell("A" + to_string(row)).value<int>() << endl;
            cout << "Name: " << ws.cell("B" + to_string(row)).value<string>() << endl;
            cout << "Department: " << ws.cell("C" + to_string(row)).value<string>() << endl;
            cout << "Salary: $" << ws.cell("D" + to_string(row)).value<double>() << endl;
            cout << "Attendance: " << ws.cell("E" + to_string(row)).value<string>() << endl;
            return;
        }
        row++;
    }
    cout << "Employee not found!\n";
}

