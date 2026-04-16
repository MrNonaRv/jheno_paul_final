const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const dbPath = path.join(__dirname, 'data', 'payroll.db');
const dbDir = path.dirname(dbPath);

if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
}

let db;
try {
    db = new Database(dbPath);
    // Test if the database is readable
    db.prepare('SELECT name FROM sqlite_master LIMIT 1').get();
} catch (err) {
    if (err.code === 'SQLITE_CORRUPT') {
        console.error('CRITICAL: Database file is corrupted (malformed).');
        console.log('Attempting to self-heal by renaming the corrupted file...');
        const timestamp = Date.now();
        const corruptPath = dbPath + '.corrupt.' + timestamp;
        try {
            if (db) db.close();
            fs.renameSync(dbPath, corruptPath);
            console.log(`Corrupted database renamed to: ${path.basename(corruptPath)}`);
            console.log('Starting with a fresh database...');
            db = new Database(dbPath);
        } catch (renameErr) {
            console.error('Failed to rename corrupted database:', renameErr);
            process.exit(1);
        }
    } else {
        throw err;
    }
}

db.pragma('foreign_keys = ON');

// Initialize tables
db.exec(`
  CREATE TABLE IF NOT EXISTS settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    meta_key TEXT UNIQUE,
    meta_value TEXT
  );

  CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    position TEXT,
    department_id INTEGER,
    monthly_salary REAL DEFAULT 0.00,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (department_id) REFERENCES departments(id) ON DELETE SET NULL
  );

  CREATE TABLE IF NOT EXISTS payroll (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER,
    period_start DATE,
    period_end DATE,
    basic_pay REAL DEFAULT 0.00,
    pera REAL DEFAULT 0.00,
    rata REAL DEFAULT 0.00,
    clothing_allowance REAL DEFAULT 0.00,
    disbursement_date DATE,
    disbursing_officer TEXT,
    is_disbursed INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS payroll_deductions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    payroll_id INTEGER NOT NULL,
    deduction_name TEXT NOT NULL,
    amount REAL DEFAULT 0.00,
    FOREIGN KEY (payroll_id) REFERENCES payroll(id) ON DELETE CASCADE
  );

  CREATE INDEX IF NOT EXISTS idx_payroll_deductions_payroll_id ON payroll_deductions(payroll_id);
  CREATE INDEX IF NOT EXISTS idx_payroll_employee_id ON payroll(employee_id);
  CREATE INDEX IF NOT EXISTS idx_payroll_period ON payroll(period_start, period_end);
  CREATE INDEX IF NOT EXISTS idx_employees_name ON employees(name);
  CREATE INDEX IF NOT EXISTS idx_employees_name_lower ON employees(TRIM(LOWER(name)));
  CREATE INDEX IF NOT EXISTS idx_departments_name_lower ON departments(TRIM(LOWER(name)));
`);

// ── LEGACY MIGRATION ──
try {
    // Check if the legacy column 'department' exists on 'employees' table
    const employeeCols = db.prepare("PRAGMA table_info(employees)").all();
    const hasLegacyDept = employeeCols.some(c => c.name === 'department');
    
    if (hasLegacyDept) {
        console.log("Migrating database to normalized V2 schema...");
        db.transaction(() => {
            // 1. Rename old tables
            db.exec('ALTER TABLE employees RENAME TO legacy_employees');
            db.exec('ALTER TABLE payroll RENAME TO legacy_payroll');
            
            // 2. Recreate schema properly
            db.exec(`
              CREATE TABLE employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                position TEXT,
                department_id INTEGER,
                monthly_salary REAL DEFAULT 0.00,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (department_id) REFERENCES departments(id) ON DELETE SET NULL
              );
              CREATE TABLE payroll (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER,
                period_start DATE,
                period_end DATE,
                basic_pay REAL DEFAULT 0.00,
                pera REAL DEFAULT 0.00,
                rata REAL DEFAULT 0.00,
                clothing_allowance REAL DEFAULT 0.00,
                disbursement_date DATE,
                disbursing_officer TEXT,
                is_disbursed INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
              );
            `);
            
            // 3. Migrate Departments
            const oldDepts = db.prepare("SELECT DISTINCT department FROM legacy_employees WHERE department IS NOT NULL AND department != ''").all();
            const insertDept = db.prepare("INSERT OR IGNORE INTO departments (name) VALUES (?)");
            oldDepts.forEach(d => insertDept.run(d.department));
            
            // 4. Migrate Employees
            const oldEmps = db.prepare("SELECT * FROM legacy_employees").all();
            const getDeptId = db.prepare("SELECT id FROM departments WHERE name = ?");
            const insertEmp = db.prepare("INSERT INTO employees (id, employee_id, name, position, department_id, monthly_salary, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)");
            oldEmps.forEach(e => {
                const deptId = e.department ? (getDeptId.get(e.department)?.id || null) : null;
                insertEmp.run(e.id, e.employee_id, e.name, e.position, deptId, e.monthly_salary, e.created_at);
            });
            
            // 5. Migrate Payroll & Deductions
            const oldPayrolls = db.prepare("SELECT * FROM legacy_payroll").all();
            const insertPay = db.prepare("INSERT INTO payroll (id, employee_id, period_start, period_end, basic_pay, pera, rata, clothing_allowance, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)");
            const insertDed = db.prepare("INSERT INTO payroll_deductions (payroll_id, deduction_name, amount) VALUES (?, ?, ?)");
            
            oldPayrolls.forEach(p => {
                // map "yyyy-mm" to simple guess bounds
                let start = null, end = null;
                if (p.month_year && p.month_year.length === 7) {
                    start = p.month_year + '-01';
                    end = p.month_year + '-30';
                }
                
                insertPay.run(p.id, p.employee_id, start, end, p.basic_pay, p.pera, p.rata, p.clothing_allowance, p.created_at);
                
                const dedMap = ['tax', 'gsis', 'philhealth', 'pagibig', 'mpl_lite', 'emergency_loan', 'computer_loan', 'rural_bank_loan', 'fcb_loan', 'mpl_gsis', 'gfal', 'sss_premium', 'mpl_lite_loan', 'policy_loan'];
                
                dedMap.forEach(dKey => {
                    if (p[dKey] && parseFloat(p[dKey]) > 0) {
                        insertDed.run(p.id, dKey, parseFloat(p[dKey]));
                    }
                });
            });
            
            // 6. Cleanup legacy tables
            // Since `payroll_deductions` was created BEFORE the rename, its foreign key 
            // automatically followed `payroll` -> `legacy_payroll`. 
            // We must recreate it pointing to the NEW `payroll` before dropping.
            db.exec(`
              CREATE TABLE IF NOT EXISTS _new_payroll_deductions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                payroll_id INTEGER NOT NULL,
                deduction_name TEXT NOT NULL,
                amount REAL DEFAULT 0.00,
                FOREIGN KEY (payroll_id) REFERENCES payroll(id) ON DELETE CASCADE
              );
              INSERT INTO _new_payroll_deductions SELECT * FROM payroll_deductions;
              DROP TABLE payroll_deductions;
              ALTER TABLE _new_payroll_deductions RENAME TO payroll_deductions;
            `);
            
            db.exec('DROP TABLE legacy_employees');
            db.exec('DROP TABLE legacy_payroll');
        })();
        console.log("Migration successful!");
    }
} catch (e) {
    console.error("Migration Error:", e);
}

// Function to safely run initial inserts
function seed() {
    const defaultSettings = [
        ['system_logo', 'logo.png'],
        ['lgu_name', 'MUNICIPALITY OF MAMBUSAO'],
        ['lgu_province', 'PROVINCE OF CAPIZ'],
        ['lgu_office', 'OFFICE OF THE MUNICIPAL ACCOUNTANT'],
        ['field_tax', 'Tax'],
        ['field_philhealth', 'PhilHealth'],
        ['field_gsis', 'L/R Premium'],
        ['field_pagibig', 'Pag-ibig Premium'],
        ['field_mpl_lite', 'MPL Lite'],
        ['field_emergency_loan', 'Emergency Loan'],
        ['field_mpl_gsis', 'MPL Loan'],
        ['field_computer_loan', 'Computer Loan'],
        ['field_rural_bank_loan', 'Rural Bank Loan'],
        ['field_fcb_loan', 'FCB Loan'],
        ['field_fcb_loan_2', 'FCB Loan 2'],
        ['field_pera', 'PERA'],
        ['field_rata', 'RATA'],
        ['field_clothing', 'Clothing'],
        ['sheet1_name', 'Sheet 1 (1-15)'],
        ['sheet2_name', 'Sheet 2 (16-31)']
    ];

    const insertSetting = db.prepare('INSERT OR IGNORE INTO settings (meta_key, meta_value) VALUES (?, ?)');
    defaultSettings.forEach(s => insertSetting.run(s[0], s[1]));

    const employeesCount = db.prepare('SELECT COUNT(*) as count FROM employees').get().count;
    if (employeesCount === 0) {
        db.prepare("INSERT OR IGNORE INTO departments (name) VALUES ('Administration')").run();
        const d_id = db.prepare("SELECT id FROM departments WHERE name = 'Administration'").get().id;
        
        const insertEmployee = db.prepare('INSERT INTO employees (employee_id, name, position, department_id, monthly_salary) VALUES (?, ?, ?, ?, ?)');
        insertEmployee.run('EMP001', 'Alcantara, Jose R.', 'Admin Officer III', d_id, 12843.00);
        insertEmployee.run('EMP002', 'Bautista, Maria C.', 'Adm Aide II', d_id, 12183.00);
    }
}

seed();

module.exports = {
    db,
    get_setting: (key, defaultValue = '') => {
        const row = db.prepare('SELECT meta_value FROM settings WHERE meta_key = ?').get(key);
        return row ? row.meta_value : defaultValue;
    }
};
