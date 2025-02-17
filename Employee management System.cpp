#include <iostream>
#include <xlnt/xlnt.hpp>

using namespace std;
 
const string FILE_NAME = "employees.xlsx";
 
void initializeExcel() {
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();
 
    ws.cell("A1").value("ID");
    ws.cell("B1").value("Name");
    ws.cell("C1").value("Department");
    ws.cell("D1").value("Salary");
    ws.cell("E1").value("Attendance");

    wb.save(FILE_NAME);
}
