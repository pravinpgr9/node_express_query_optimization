const express = require("express");
const mysql = require("mysql2/promise");
const moment = require("moment");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
require("dotenv").config();

const app = express();
app.use(express.json());

// Database connection
const db = mysql.createPool({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASS,
  database: process.env.DB_NAME,
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
});

// GET employee details
const getEmployeeDetails = async (employeeIds) => {
  const query = `
    SELECT 
      id, emp_code, first_name, middle_name, last_name, email, mobile_number,
      cluster_id, plant_id, band_id, designation_id, department_id, 
      sub_department_id, reporting_manager, job_status, status
    FROM users WHERE id IN (?)`;

  const [rows] = await db.query(query, [employeeIds]);
  return rows;
};

// GET timesheet data
const getTimesheetData = async (table, employeeId, startDate, endDate) => {
  const query = `
    SELECT id, emp_id, date, duration, remark, bms_parameter, updated_at
    FROM ${table} 
    WHERE emp_id = ? AND date BETWEEN ? AND ? AND status = 1
    ORDER BY date DESC`;

  const [rows] = await db.query(query, [employeeId, startDate, endDate]);
  return rows;
};

// API to fetch timesheet report
app.post("/overall-timesheet-report", async (req, res) => {
  const { start_date, end_date, report_type } = req.body;

  if (!start_date || !end_date) {
    return res.status(400).json({ error: 1, message: "Start and end date required" });
  }

  try {
    const startDate = moment(start_date, "DD-MM-YYYY").format("YYYY-MM-DD");
    const endDate = moment(end_date, "DD-MM-YYYY").format("YYYY-MM-DD");

    // Fetch all active employees
    const [employeeRows] = await db.query("SELECT id FROM users WHERE status = 1");
    const employeeIds = employeeRows.map((emp) => emp.id);

    // Fetch employee details
    const employees = await getEmployeeDetails(employeeIds);

    let employee_timesheet_data = [];

    // Loop through employees and fetch timesheet data
    for (const employee of employees) {
      const employeeData = {
        id: employee.id,
        emp_code: employee.emp_code,
        full_name: `${employee.first_name} ${employee.middle_name ?? ""} ${employee.last_name}`,
        designation: employee.designation_id,
        department: employee.department_id,
        plant: employee.plant_id,
        band: employee.band_id,
        reporting_manager: employee.reporting_manager,
        job_status: employee.status === 1 ? "Active" : "Separated",
      };

      const timesheetData = {
        kra_kpi: await getTimesheetData("kra_kpi", employee.id, startDate, endDate),
        routine: await getTimesheetData("routine", employee.id, startDate, endDate),
        initiative: await getTimesheetData("initiative", employee.id, startDate, endDate),
        project: await getTimesheetData("project", employee.id, startDate, endDate),
        onetime: await getTimesheetData("onetime", employee.id, startDate, endDate),
        leave: await getTimesheetData("leave_timesheet", employee.id, startDate, endDate),
      };

      employee_timesheet_data.push({ employee: employeeData, ...timesheetData });
    }

    if (report_type === "Excel") {
      return generateExcelReport(employee_timesheet_data, res);
    }

    return res.status(200).json({
      error: 0,
      message: "Timesheet data retrieved successfully",
      data: employee_timesheet_data,
    });

  } catch (err) {
    return res.status(500).json({ error: 1, message: err.message });
  }
});

// Function to generate Excel report
const generateExcelReport = async (employeeTimesheetData, res) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Timesheet Report");

  worksheet.columns = [
    { header: "Emp Code", key: "emp_code", width: 15 },
    { header: "Full Name", key: "full_name", width: 25 },
    { header: "Designation", key: "designation", width: 20 },
    { header: "Department", key: "department", width: 20 },
    { header: "Plant", key: "plant", width: 20 },
    { header: "Job Status", key: "job_status", width: 15 },
    { header: "KRA/KPI", key: "kra_kpi", width: 30 },
    { header: "Routine", key: "routine", width: 30 },
    { header: "Initiative", key: "initiative", width: 30 },
    { header: "Project", key: "project", width: 30 },
    { header: "Onetime", key: "onetime", width: 30 },
    { header: "Leave", key: "leave", width: 30 },
  ];

  employeeTimesheetData.forEach((record) => {
    worksheet.addRow({
      emp_code: record.employee.emp_code,
      full_name: record.employee.full_name,
      designation: record.employee.designation,
      department: record.employee.department,
      plant: record.employee.plant,
      job_status: record.employee.job_status,
      kra_kpi: JSON.stringify(record.kra_kpi),
      routine: JSON.stringify(record.routine),
      initiative: JSON.stringify(record.initiative),
      project: JSON.stringify(record.project),
      onetime: JSON.stringify(record.onetime),
      leave: JSON.stringify(record.leave),
    });
  });

  const folderPath = path.join(__dirname, "reports");
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath);
  }

  const filePath = path.join(folderPath, `timesheet_report_${Date.now()}.xlsx`);
  await workbook.xlsx.writeFile(filePath);

  return res.status(200).json({
    error: 0,
    message: "Excel file generated successfully",
    file_url: `http://localhost:3000/reports/${path.basename(filePath)}`,
  });
};

// Serve reports directory
app.use("/reports", express.static(path.join(__dirname, "reports")));

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
