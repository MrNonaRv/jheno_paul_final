const express = require('express');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
const path = require('path');
const multer = require('multer');
const { db, get_setting } = require('./db');
const fs = require('fs');

const app = express();

// Middleware
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Helper: Get all settings for templates
function getGlobalSettings() {
    const settings = {};
    const rows = db.prepare('SELECT meta_key, meta_value FROM settings').all();
    rows.forEach(r => settings[r.meta_key] = r.meta_value);
    return settings;
}

// Routes
app.get('/', (req, res) => {
    const settings = getGlobalSettings();
    const employees = db.prepare(`SELECT e.*, d.name as department FROM employees e LEFT JOIN departments d ON e.department_id = d.id ORDER BY e.name ASC`).all();
    const departments = db.prepare("SELECT name FROM departments ORDER BY name ASC").all().map(r => r.name);
    
    const periodsRaw = db.prepare("SELECT DISTINCT strftime('%Y-%m', period_start) as month_year FROM payroll WHERE period_start IS NOT NULL ORDER BY month_year DESC").all();
    const months = periodsRaw.map(r => r.month_year).filter(Boolean);
    
    // Fetch payroll with employee details dynamically pivoted
    const payroll_rows = db.prepare(`
        SELECT p.*, e.name, e.position, d.name AS department, e.monthly_salary, e.id AS emp_db_id,
        strftime('%d', p.period_start) as start_day,
        strftime('%Y-%m', p.period_start) as month_year_val,
        MAX(CASE WHEN pd.deduction_name = 'tax' THEN pd.amount END) AS tax,
        MAX(CASE WHEN pd.deduction_name = 'gsis' THEN pd.amount END) AS gsis,
        MAX(CASE WHEN pd.deduction_name = 'philhealth' THEN pd.amount END) AS philhealth,
        MAX(CASE WHEN pd.deduction_name = 'pagibig' THEN pd.amount END) AS pagibig,
        MAX(CASE WHEN pd.deduction_name = 'mpl_lite' THEN pd.amount END) AS mpl_lite,
        MAX(CASE WHEN pd.deduction_name = 'emergency_loan' THEN pd.amount END) AS emergency_loan,
        MAX(CASE WHEN pd.deduction_name = 'mpl_gsis' THEN pd.amount END) AS mpl_gsis,
        MAX(CASE WHEN pd.deduction_name = 'computer_loan' THEN pd.amount END) AS computer_loan,
        MAX(CASE WHEN pd.deduction_name = 'rural_bank_loan' THEN pd.amount END) AS rural_bank_loan,
        MAX(CASE WHEN pd.deduction_name = 'fcb_loan' THEN pd.amount END) AS fcb_loan,
        MAX(CASE WHEN pd.deduction_name = 'fcb_loan_2' THEN pd.amount END) AS fcb_loan_2,
        MAX(CASE WHEN pd.deduction_name = 'gfal' THEN pd.amount END) AS gfal,
        MAX(CASE WHEN pd.deduction_name = 'sss_premium' THEN pd.amount END) AS sss_premium,
        MAX(CASE WHEN pd.deduction_name = 'mpl_lite_loan' THEN pd.amount END) AS mpl_lite_loan,
        MAX(CASE WHEN pd.deduction_name = 'policy_loan' THEN pd.amount END) AS policy_loan
        FROM payroll p
        JOIN employees e ON p.employee_id = e.id
        LEFT JOIN departments d ON e.department_id = d.id
        LEFT JOIN payroll_deductions pd ON p.id = pd.payroll_id
        GROUP BY p.id
        ORDER BY e.id ASC, p.id ASC
    `).all();

    res.render('index', { settings, employees, departments, months, payroll_rows });
});

app.get('/employees', (req, res) => {
    const settings = getGlobalSettings();
    const employees = db.prepare('SELECT e.*, d.name AS department FROM employees e LEFT JOIN departments d ON e.department_id = d.id ORDER BY e.name ASC').all();
    const departments = db.prepare("SELECT name FROM departments ORDER BY name ASC").all().map(r => r.name);
    const editId = req.query.edit;
    let editEmployee = null;
    if (editId) {
        editEmployee = db.prepare('SELECT e.*, d.name AS department FROM employees e LEFT JOIN departments d ON e.department_id = d.id WHERE e.id = ?').get(editId);
    }
    res.render('employees', { settings, employees, departments, editEmployee, editMode: !!editEmployee });
});

app.get('/admin', (req, res) => {
    const settings = getGlobalSettings();
    const departments = db.prepare("SELECT * FROM departments ORDER BY name ASC").all();
    res.render('admin', { settings, departments, message: req.query.message || '' });
});

app.post('/api/add-department', (req, res) => {
    try {
        db.prepare("INSERT INTO departments (name) VALUES (?)").run(req.body.name);
        res.redirect('/admin?message=Department added successfully.');
    } catch (e) {
        res.send("Error adding department: " + e.message);
    }
});

app.post('/api/delete-department', (req, res) => {
    try {
        db.prepare("DELETE FROM departments WHERE id = ?").run(req.body.id);
        res.redirect('/admin?message=Department deleted successfully.');
    } catch (e) {
        res.send("Error deleting department: " + e.message);
    }
});

// API Endpoints
app.post('/api/update-cell', (req, res) => {
    const { type, field, id, value } = req.body;
    try {
        if (type === 'employee') {
            if (field === 'department') {
                const dept = db.prepare("SELECT id FROM departments WHERE name = ?").get(value);
                if (dept) {
                    db.prepare(`UPDATE employees SET department_id = ? WHERE id = ?`).run(dept.id, id);
                } else {
                    return res.json({ success: false, message: 'Invalid department' });
                }
            } else {
                db.prepare(`UPDATE employees SET ${field} = ? WHERE id = ?`).run(value, id);
            }
        } else if (type === 'payroll') {
            const deds = ['tax', 'gsis', 'philhealth', 'pagibig', 'mpl_lite', 'emergency_loan', 'computer_loan', 'rural_bank_loan', 'fcb_loan', 'fcb_loan_2', 'mpl_gsis', 'gfal', 'sss_premium', 'mpl_lite_loan', 'policy_loan'];
            if (deds.includes(field)) {
                if (!value || value === '0' || value === '0.00' || value === '') {
                    db.prepare("DELETE FROM payroll_deductions WHERE payroll_id = ? AND deduction_name = ?").run(id, field);
                } else {
                    const existing = db.prepare("SELECT id FROM payroll_deductions WHERE payroll_id = ? AND deduction_name = ?").get(id, field);
                    if (existing) {
                        db.prepare("UPDATE payroll_deductions SET amount = ? WHERE id = ?").run(value, existing.id);
                    } else {
                        db.prepare("INSERT INTO payroll_deductions (payroll_id, deduction_name, amount) VALUES (?, ?, ?)").run(id, field, value);
                    }
                }
            } else {
                db.prepare(`UPDATE payroll SET ${field} = ? WHERE id = ?`).run(value, id);
            }
        }
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

app.get('/api/get-payroll', (req, res) => {
    const id = req.query.id;
    const data = db.prepare(`
        SELECT p.*, e.monthly_salary,
        MAX(CASE WHEN pd.deduction_name = 'tax' THEN pd.amount END) AS tax,
        MAX(CASE WHEN pd.deduction_name = 'gsis' THEN pd.amount END) AS gsis,
        MAX(CASE WHEN pd.deduction_name = 'philhealth' THEN pd.amount END) AS philhealth,
        MAX(CASE WHEN pd.deduction_name = 'pagibig' THEN pd.amount END) AS pagibig,
        MAX(CASE WHEN pd.deduction_name = 'mpl_lite' THEN pd.amount END) AS mpl_lite,
        MAX(CASE WHEN pd.deduction_name = 'emergency_loan' THEN pd.amount END) AS emergency_loan,
        MAX(CASE WHEN pd.deduction_name = 'mpl_gsis' THEN pd.amount END) AS mpl_gsis,
        MAX(CASE WHEN pd.deduction_name = 'computer_loan' THEN pd.amount END) AS computer_loan,
        MAX(CASE WHEN pd.deduction_name = 'rural_bank_loan' THEN pd.amount END) AS rural_bank_loan,
        MAX(CASE WHEN pd.deduction_name = 'fcb_loan' THEN pd.amount END) AS fcb_loan,
        MAX(CASE WHEN pd.deduction_name = 'fcb_loan_2' THEN pd.amount END) AS fcb_loan_2,
        MAX(CASE WHEN pd.deduction_name = 'gfal' THEN pd.amount END) AS gfal,
        MAX(CASE WHEN pd.deduction_name = 'sss_premium' THEN pd.amount END) AS sss_premium,
        MAX(CASE WHEN pd.deduction_name = 'mpl_lite_loan' THEN pd.amount END) AS mpl_lite_loan,
        MAX(CASE WHEN pd.deduction_name = 'policy_loan' THEN pd.amount END) AS policy_loan
        FROM payroll p 
        JOIN employees e ON p.employee_id = e.id 
        LEFT JOIN payroll_deductions pd ON p.id = pd.payroll_id
        WHERE p.id = ?
        GROUP BY p.id
    `).get(id);

    if (data) {
        data.month_year = data.period_start ? `${data.period_start} to ${data.period_end}` : '';
        res.json({ success: true, data });
    } else {
        res.json({ success: false, message: 'Record not found' });
    }
});

app.post('/api/save-payroll', (req, res) => {
    const d = req.body;
    try {
        let period_start = d.period_start || null;
        let period_end = d.period_end || null;
        
        let pid = d.payroll_id;

        db.transaction(() => {
            if (pid) {
                db.prepare(`
                    UPDATE payroll SET 
                    employee_id = ?, period_start = ?, period_end = ?, basic_pay = ?, pera = ?, rata = ?, clothing_allowance = ?
                    WHERE id = ?
                `).run(
                    d.employee_id, period_start, period_end, d.basic_pay, d.pera, d.rata, d.clothing_allowance || 0, pid
                );
                // Clear old deductions to easily replace 
                db.prepare("DELETE FROM payroll_deductions WHERE payroll_id = ?").run(pid);
            } else {
                const rs = db.prepare(`
                    INSERT INTO payroll (
                        employee_id, period_start, period_end, basic_pay, pera, rata, clothing_allowance
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                `).run(
                    d.employee_id, period_start, period_end, d.basic_pay, d.pera, d.rata, d.clothing_allowance || 0
                );
                pid = rs.lastInsertRowid;
            }

            const deds = ['tax', 'gsis', 'philhealth', 'pagibig', 'mpl_lite', 'emergency_loan', 'computer_loan', 'rural_bank_loan', 'fcb_loan', 'fcb_loan_2', 'mpl_gsis', 'gfal', 'sss_premium', 'mpl_lite_loan', 'policy_loan'];
            const insertDed = db.prepare("INSERT INTO payroll_deductions (payroll_id, deduction_name, amount) VALUES (?, ?, ?)");
            for (const key of deds) {
                if (d[key] && parseFloat(d[key]) > 0) {
                    insertDed.run(pid, key, parseFloat(d[key]));
                }
            }
        })();
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

app.post('/api/delete-payroll', (req, res) => {
    try {
        db.prepare('DELETE FROM payroll WHERE id = ?').run(req.body.id);
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

app.post('/api/update-sheet-name', (req, res) => {
    const { sheet, name } = req.body;
    try {
        const key = sheet === 'sheet1' ? 'sheet1_name' : 'sheet2_name';
        db.prepare('UPDATE settings SET meta_value = ? WHERE meta_key = ?').run(name, key);
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

app.post('/api/save-employee', (req, res) => {
    const d = req.body;
    try {
        let deptId = null;
        if (d.department) {
            let dept = db.prepare('SELECT id FROM departments WHERE TRIM(LOWER(name)) = TRIM(LOWER(?))').get(d.department);
            if (!dept) {
                const dr = db.prepare('INSERT INTO departments (name) VALUES (?)').run(d.department);
                deptId = dr.lastInsertRowid;
            } else {
                deptId = dept.id;
            }
        }

        if (d.id) {
            // Update existing — keep existing employee_id if none is provided
            const existing = db.prepare('SELECT employee_id FROM employees WHERE id = ?').get(d.id);
            const empId = d.employee_id || existing.employee_id;
            db.prepare('UPDATE employees SET employee_id = ?, name = ?, position = ?, department_id = ?, monthly_salary = ? WHERE id = ?')
              .run(empId, d.name, d.position, deptId, d.monthly_salary, d.id);
        } else {
            // Insert new — auto-generate employee_id from the new row's ID
            const tempId = 'TMP_' + Date.now();
            const result = db.prepare('INSERT INTO employees (employee_id, name, position, department_id, monthly_salary) VALUES (?, ?, ?, ?, ?)')
              .run(tempId, d.name, d.position, deptId, d.monthly_salary);
            const newId = result.lastInsertRowid;
            const empId = 'EMP' + String(newId).padStart(5, '0');
            db.prepare('UPDATE employees SET employee_id = ? WHERE id = ?').run(empId, newId);
        }
        res.redirect('/employees');
    } catch (err) {
        res.send('Error saving employee: ' + err.message);
    }
});

app.post('/api/delete-employee', (req, res) => {
    try {
        db.prepare('DELETE FROM employees WHERE id = ?').run(req.body.id);
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

// Admin Settings Update
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, 'public/uploads/'),
    filename: (req, file, cb) => cb(null, Date.now() + '_' + file.originalname)
});
const upload = multer({ storage: storage });

app.post('/admin/update', upload.single('logo'), (req, res) => {
    const d = req.body;
    try {
        const updateSetting = db.prepare('UPDATE settings SET meta_value = ? WHERE meta_key = ?');
        const insertSetting = db.prepare('INSERT INTO settings (meta_key, meta_value) VALUES (?, ?)');
        const checkSetting = db.prepare('SELECT id FROM settings WHERE meta_key = ?');

        const saveSetting = (key, val) => {
            if (checkSetting.get(key)) updateSetting.run(val, key);
            else insertSetting.run(key, val);
        };

        saveSetting('lgu_name', d.lgu_name);
        saveSetting('lgu_province', d.lgu_province);
        saveSetting('lgu_office', d.lgu_office);
        
        const fields = [
            'field_tax', 'field_philhealth', 'field_gsis', 'field_pagibig', 
            'field_mpl_lite', 'field_emergency_loan', 'field_mpl_gsis', 
            'field_computer_loan', 'field_rural_bank_loan', 'field_fcb_loan',
            'field_pera', 'field_rata', 'field_clothing'
        ];
        fields.forEach(f => { if (d[f]) saveSetting(f, d[f]); });

        if (req.file) {
            saveSetting('system_logo', req.file.filename);
        }
        
        res.redirect('/admin?message=Settings updated successfully');
    } catch (err) {
        res.send('Error updating settings: ' + err.message);
    }
});

app.post('/admin/clear-payroll', (req, res) => {
    try {
        db.prepare('DELETE FROM payroll').run();
        res.redirect('/admin?message=All payroll records have been successfully deleted.');
    } catch (err) {
        res.send('Error clearing payroll data: ' + err.message);
    }
});

app.post('/admin/clear-all', (req, res) => {
    try {
        db.transaction(() => {
            db.prepare('DELETE FROM payroll_deductions').run();
            db.prepare('DELETE FROM payroll').run();
            db.prepare('DELETE FROM employees').run();
            db.prepare('DELETE FROM departments').run();
        })();
        res.redirect('/admin?message=System has been fully reset. All data deleted.');
    } catch (err) {
        res.send('Error resetting system: ' + err.message);
    }
});

// Duplicate Management APIs
app.get('/api/admin/scan-duplicates', (req, res) => {
    try {
        const sql = `
            SELECT 
                e.name, 
                p.employee_id, 
                strftime('%Y-%m', p.period_start) as month_year,
                CASE WHEN strftime('%d', p.period_start) <= '15' THEN 1 ELSE 2 END as part,
                COUNT(*) as count
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            GROUP BY p.employee_id, month_year, part
            HAVING count > 1
        `;
        const duplicates = db.prepare(sql).all();
        res.json({ success: true, data: duplicates });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

app.post('/api/admin/delete-duplicate-group', (req, res) => {
    const { empId, monthYear, part } = req.body;
    try {
        db.transaction(() => {
            const dayCondition = part == 1 ? "<= '15'" : "> '15'";
            const sql = `
                SELECT id FROM payroll 
                WHERE employee_id = ? 
                AND strftime('%Y-%m', period_start) = ?
                AND strftime('%d', period_start) ${dayCondition}
                ORDER BY id DESC
            `;
            const records = db.prepare(sql).all(empId, monthYear);
            
            // Keep the first one (latest ID), delete the rest
            if (records.length > 1) {
                const toDelete = records.slice(1).map(r => r.id);
                const deleteDeds = db.prepare('DELETE FROM payroll_deductions WHERE payroll_id = ?');
                const deletePayroll = db.prepare('DELETE FROM payroll WHERE id = ?');
                
                for (const id of toDelete) {
                    deleteDeds.run(id);
                    deletePayroll.run(id);
                }
            }
        })();
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

app.post('/api/admin/delete-all-duplicates', (req, res) => {
    try {
        db.transaction(() => {
            const scanSql = `
                SELECT 
                    employee_id, 
                    strftime('%Y-%m', period_start) as month_year,
                    CASE WHEN strftime('%d', p.period_start) <= '15' THEN 1 ELSE 2 END as part
                FROM payroll p
                GROUP BY employee_id, month_year, part
                HAVING COUNT(*) > 1
            `;
            const groups = db.prepare(scanSql).all();
            
            const deleteDeds = db.prepare('DELETE FROM payroll_deductions WHERE payroll_id = ?');
            const deletePayroll = db.prepare('DELETE FROM payroll WHERE id = ?');

            for (const g of groups) {
                const dayCondition = g.part == 1 ? "<= '15'" : "> '15'";
                const fetchSql = `
                    SELECT id FROM payroll 
                    WHERE employee_id = ? 
                    AND strftime('%Y-%m', period_start) = ?
                    AND strftime('%d', period_start) ${dayCondition}
                    ORDER BY id DESC
                `;
                const records = db.prepare(fetchSql).all(g.employee_id, g.month_year);
                
                if (records.length > 1) {
                    const toDelete = records.slice(1).map(r => r.id);
                    for (const id of toDelete) {
                        deleteDeds.run(id);
                        deletePayroll.run(id);
                    }
                }
            }
        })();
        res.json({ success: true });
    } catch (err) {
        res.json({ success: false, message: err.message });
    }
});

// Full Data Export (SQL Format)
app.get('/api/export-data', (req, res) => {
    try {
        const lines = [];
        lines.push('-- LGU Payroll System -- Full Data Backup');
        lines.push('-- Generated: ' + new Date().toISOString());
        lines.push('-- Import this file via the "Import / Load Data" button.');
        lines.push('');

        // Settings
        lines.push('-- SETTINGS');
        const settings = db.prepare('SELECT * FROM settings ORDER BY id ASC').all();
        settings.forEach(s => {
            const val = s.meta_value ? s.meta_value.replace(/'/g, "''") : '';
            lines.push(`INSERT OR REPLACE INTO settings (id, meta_key, meta_value) VALUES (${s.id}, '${s.meta_key}', '${val}');`);
        });
        lines.push('');

        // Departments
        lines.push('-- DEPARTMENTS');
        lines.push('DELETE FROM departments;');
        const depts = db.prepare('SELECT * FROM departments ORDER BY id ASC').all();
        depts.forEach(d => {
            const name = d.name.replace(/'/g, "''");
            lines.push(`INSERT INTO departments (id, name) VALUES (${d.id}, '${name}');`);
        });
        lines.push('');

        // Employees
        lines.push('-- EMPLOYEES');
        lines.push('DELETE FROM employees;');
        const employees = db.prepare('SELECT * FROM employees ORDER BY id ASC').all();
        employees.forEach(e => {
            const name = e.name.replace(/'/g, "''");
            const pos = (e.position || '').replace(/'/g, "''");
            lines.push(`INSERT INTO employees (id, employee_id, name, position, department_id, monthly_salary) VALUES (${e.id}, '${e.employee_id}', '${name}', '${pos}', ${e.department_id || 'NULL'}, ${e.monthly_salary});`);
        });
        lines.push('');

        // Payroll
        lines.push('-- PAYROLL');
        lines.push('DELETE FROM payroll;');
        const payroll = db.prepare('SELECT * FROM payroll ORDER BY id ASC').all();
        payroll.forEach(p => {
            const cols = Object.keys(p).filter(k => k !== 'created_at');
            const vals = cols.map(c => {
                const val = p[c];
                if (val === null) return 'NULL';
                if (typeof val === 'number') return val;
                return `'${String(val).replace(/'/g, "''")}'`;
            });
            lines.push(`INSERT INTO payroll (${cols.join(',')}) VALUES (${vals.join(',')});`);
        });
        lines.push('');

        // Deductions
        lines.push('-- DEDUCTIONS');
        lines.push('DELETE FROM payroll_deductions;');
        const deductions = db.prepare('SELECT * FROM payroll_deductions ORDER BY id ASC').all();
        deductions.forEach(d => {
            const name = d.deduction_name.replace(/'/g, "''");
            lines.push(`INSERT INTO payroll_deductions (id, payroll_id, deduction_name, amount) VALUES (${d.id}, ${d.payroll_id}, '${name}', ${d.amount});`);
        });

        lines.push('');
        lines.push('-- End of backup');

        const filename = `payroll_backup_${new Date().toISOString().replace(/[:.]/g, '-')}.sql`;
        res.setHeader('Content-disposition', 'attachment; filename=' + filename);
        res.setHeader('Content-type', 'text/sql');
        res.send(lines.join('\n'));
    } catch (err) {
        console.error(err);
        res.status(500).send('Export failed: ' + err.message);
    }
});

// Full Data Import (SQL and Excel Format)
app.post('/api/import-data', upload.single('backup_file'), (req, res) => {
    if (!req.file) return res.json({ success: false, message: 'No file uploaded' });

    const ext = path.extname(req.file.originalname).toLowerCase();

    // ── Excel Import ─────────────────────────────────────────────────────────
    if (ext === '.xlsx' || ext === '.xls') {
        try {
            const workbook = XLSX.readFile(req.file.path);
            
            const getEmployee = db.prepare('SELECT id FROM employees WHERE TRIM(LOWER(name)) = TRIM(LOWER(?))');
            const getDeptAdmin = db.prepare('SELECT id FROM departments WHERE TRIM(LOWER(name)) = TRIM(LOWER(?))');
            const insertDept = db.prepare('INSERT INTO departments (name) VALUES (?)');
            const insertEmployee = db.prepare('INSERT INTO employees (employee_id, name, position, department_id, monthly_salary) VALUES (?, ?, ?, ?, ?)');
            const getPayroll = db.prepare('SELECT id FROM payroll WHERE employee_id = ? AND period_start IS ? AND period_end IS ?');
            const insertPayroll = db.prepare(`
                INSERT INTO payroll (employee_id, period_start, period_end, basic_pay, pera, rata, clothing_allowance) VALUES (?, ?, ?, ?, ?, ?, ?)
            `);
            const updatePayroll = db.prepare(`
                UPDATE payroll SET basic_pay = ?, pera = ?, rata = ?, clothing_allowance = ? WHERE id = ?
            `);
            const delDeds = db.prepare("DELETE FROM payroll_deductions WHERE payroll_id = ?");
            const insertDed = db.prepare("INSERT INTO payroll_deductions (payroll_id, deduction_name, amount) VALUES (?, ?, ?)");

            const n = (val) => parseFloat(val) || 0;

            const runImport = db.transaction(() => {
                let count = 0;
                let duplicates = 0;

                workbook.SheetNames.forEach(sheetName => {
                    const sheet = workbook.Sheets[sheetName];
                    if (!sheet || sheetName === 'Empty') return;

                    let p_start = req.body.period_start || null;
                    let p_end = req.body.period_end || null;
                    if (!p_start && !p_end && sheet['A3'] && sheet['A3'].v) {
                        const v = String(sheet['A3'].v);
                        if (v.includes(' to ')) {
                            const parts = v.replace('Period:', '').split(' to ');
                            p_start = parts[0].trim();
                            p_end = parts[1].trim();
                        } else {
                            // legacy extraction: "April 16-30, 2026"
                            const match = v.match(/([a-zA-Z]+)\s+(\d+)-(\d+),\s+(\d+)/);
                            if (match) {
                                const mNames = {january:'01',february:'02',march:'03',april:'04',may:'05',june:'06',july:'07',august:'08',september:'09',october:'10',november:'11',december:'12'};
                                const mm = mNames[match[1].toLowerCase()] || '01';
                                p_start = `${match[4]}-${mm}-${match[2].padStart(2, '0')}`;
                                p_end = `${match[4]}-${mm}-${match[3].padStart(2, '0')}`;
                            }
                        }
                    }

                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }); 
                    
                    let headerIdx = -1;
                    for (let i = 0; i < Math.min(20, rows.length); i++) {
                        if (rows[i] && rows[i].some(cell => typeof cell === 'string' && cell.trim() === 'Name')) {
                            headerIdx = i; break;
                        }
                    }
                    
                    if (headerIdx === -1) return; // skip weird sheets

                    const head1 = rows[headerIdx] || [];
                    const head2 = rows[headerIdx + 1] || [];
                    
                    const colMap = {};
                    for (let c = 0; c < Math.max(head1.length, head2.length); c++) {
                        let name = ((head1[c] || '') + ' ' + (head2[c] || '')).trim().toLowerCase().replace(/\s+/g, ' ');
                        if (name.includes('name')) colMap.name = c;
                        else if (name.includes('position')) colMap.position = c;
                        else if (name.includes('monthly') && name.includes('salary')) colMap.monthly = c;
                        else if (name.includes('amount') && name.includes('earned')) colMap.earned = c;
                        else if (name === 'basic pay') colMap.earned = c;
                        else if (name.includes('pera')) colMap.pera = c;
                        else if (name.includes('rata')) colMap.rata = c;
                        else if (name.includes('clothing')) colMap.clothing = c;
                        else if (name.includes('tax')) colMap.tax = c;
                        else if (name.includes('rural bank')) colMap.rural_bank_loan = c;
                        else if (name.includes('fcb')) colMap.fcb_loan = c;
                        else if (name.includes('mpl gsis') || name === 'mpl') colMap.mpl_gsis = c;
                        else if (name.includes('pag-ibig') || name.includes('pagibig')) colMap.pagibig = c;
                        else if (name.includes('gfal')) colMap.gfal = c;
                        else if (name.includes('computer')) colMap.computer_loan = c;
                        else if (name.includes('emergency')) colMap.emergency_loan = c;
                        else if (name.includes('sss')) colMap.sss_premium = c;
                        else if (name.includes('policy')) colMap.policy_loan = c;
                        else if (name.includes('mpl lite loan') || name.includes('mpl pi')) colMap.mpl_lite_loan = c;
                        else if (name.includes('mpl lite')) colMap.mpl_lite = c;
                        else if (name.includes('philhealth')) colMap.philhealth = c;
                        else if (name.includes('gsis')) colMap.gsis = c;
                    }

                    for (let i = headerIdx + 2; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row || !row[colMap.name]) continue;

                        const nameVal = String(row[colMap.name]).trim();
                        if (nameVal.replace(/\s+/g, '').toUpperCase() === 'TOTAL' || nameVal === '') continue;
                        
                        let emp = getEmployee.get(nameVal);
                        if (!emp) {
                            let deptId = null;
                            const deptName = sheetName === 'Sheet1' ? 'Unassigned' : sheetName;
                            let dRow = getDeptAdmin.get(deptName);
                            if (!dRow) {
                                const dr = insertDept.run(deptName);
                                deptId = dr.lastInsertRowid;
                            } else {
                                deptId = dRow.id;
                            }

                            const tempId = 'TMP_' + Date.now() + '_' + Math.floor(Math.random() * 999999);
                            const result = insertEmployee.run(tempId, nameVal, row[colMap.position] || '', deptId, n(row[colMap.monthly]));
                            const newId = result.lastInsertRowid;
                            const empId = 'EMP' + String(newId).padStart(5, '0');
                            db.prepare('UPDATE employees SET employee_id = ? WHERE id = ?').run(empId, newId);
                            emp = { id: newId };
                        }

                        const ext_earned = n(row[colMap.earned]);
                        const ext_pera = n(row[colMap.pera]);
                        const ext_rata = n(row[colMap.rata]);
                        const ext_clothing = n(row[colMap.clothing]);

                        let payId;
                        const existing = getPayroll.get(emp.id, p_start, p_end);
                        if (existing) {
                            duplicates++;
                            updatePayroll.run(ext_earned, ext_pera, ext_rata, ext_clothing, existing.id);
                            payId = existing.id;
                            delDeds.run(payId);
                        } else {
                            const pInsert = insertPayroll.run(emp.id, p_start, p_end, ext_earned, ext_pera, ext_rata, ext_clothing);
                            payId = pInsert.lastInsertRowid;
                        }

                        const finalDeds = {};
                        for (let c = 0; c < row.length; c++) {
                            const val = n(row[c]);
                            if (val <= 0) continue;

                            let name = ((head1[c] || '') + ' ' + (head2[c] || '')).trim().toLowerCase().replace(/\s+/g, ' ');
                            let dKey = null;

                            if (name.includes('tax')) dKey = 'tax';
                            else if (name.includes('rural bank')) dKey = 'rural_bank_loan';
                            else if (name.includes('fcb')) dKey = 'fcb_loan';
                            else if (name.includes('mpl gsis') || name === 'mpl') dKey = 'mpl_gsis';
                            else if (name.includes('pag-ibig') || name.includes('pagibig')) dKey = 'pagibig';
                            else if (name.includes('gfal')) dKey = 'gfal';
                            else if (name.includes('computer')) dKey = 'computer_loan';
                            else if (name.includes('emergency')) dKey = 'emergency_loan';
                            else if (name.includes('sss')) dKey = 'sss_premium';
                            else if (name.includes('policy')) dKey = 'policy_loan';
                            else if (name.includes('mpl lite loan') || name.includes('mpl pi')) dKey = 'mpl_lite_loan';
                            else if (name.includes('mpl lite')) dKey = 'mpl_lite';
                            else if (name.includes('philhealth')) dKey = 'philhealth';
                            else if (name.includes('gsis')) dKey = 'gsis';

                            if (dKey) {
                                if (dKey === 'fcb_loan' && finalDeds['fcb_loan'] !== undefined) {
                                    finalDeds['fcb_loan_2'] = (finalDeds['fcb_loan_2'] || 0) + val;
                                } else {
                                    finalDeds[dKey] = (finalDeds[dKey] || 0) + val;
                                }
                            }
                        }

                        for (const [dKey, dVal] of Object.entries(finalDeds)) {
                            insertDed.run(payId, dKey, dVal);
                        }
                        count++;
                    }
                });
                return { count, duplicates };
            });
            const result = runImport();
            fs.unlinkSync(req.file.path);
            let msg = `Excel import successful! ${result.count} payroll records extracted.`;
            if (result.duplicates > 0) {
                msg += ` Note: ${result.duplicates} records already existed and were updated.`;
            }
            return res.json({ success: true, message: msg });

        } catch (err) {
            console.error(err);
            return res.json({ success: false, message: 'Excel import failed: ' + err.message });
        }
    }

    // ── SQL Import ────────────────────────────────────────────────────────────
    try {
        const sql = fs.readFileSync(req.file.path, 'utf8');
        if (!sql.includes('LGU Payroll System')) {
            return res.json({ success: false, message: 'Invalid backup file. Please use a .sql file exported from this system, or an Excel (.xlsx) file.' });
        }

        const statements = [];
        let current = '';
        let inString = false;

        for (let i = 0; i < sql.length; i++) {
            const char = sql[i];
            if (char === "'" && sql[i-1] !== "\\") inString = !inString;
            if (char === ';' && !inString) {
                statements.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        if (current.trim()) statements.push(current.trim());

        const runTransaction = db.transaction((stmts) => {
            for (const s of stmts) {
                const trimmed = s.trim();
                if (!trimmed || trimmed.startsWith('--')) continue;
                db.prepare(trimmed).run();
            }
        });

        runTransaction(statements);
        fs.unlinkSync(req.file.path);
        res.json({ success: true, message: `SQL import successful! ${statements.length} statements processed.` });
    } catch (err) {
        console.error(err);
        res.json({ success: false, message: 'SQL import failed: ' + err.message });
    }
});

app.get('/api/export-excel', async (req, res) => {
    try {
        const periodParam = req.query.period; // e.g. '2026-04'
        
        let query = `
            SELECT p.*, e.name, e.position, d.name AS department, e.monthly_salary,
            MAX(CASE WHEN pd.deduction_name = 'tax' THEN pd.amount END) AS tax,
            MAX(CASE WHEN pd.deduction_name = 'gsis' THEN pd.amount END) AS gsis,
            MAX(CASE WHEN pd.deduction_name = 'philhealth' THEN pd.amount END) AS philhealth,
            MAX(CASE WHEN pd.deduction_name = 'pagibig' THEN pd.amount END) AS pagibig,
            MAX(CASE WHEN pd.deduction_name = 'mpl_lite' THEN pd.amount END) AS mpl_lite,
            MAX(CASE WHEN pd.deduction_name = 'emergency_loan' THEN pd.amount END) AS emergency_loan,
            MAX(CASE WHEN pd.deduction_name = 'mpl_gsis' THEN pd.amount END) AS mpl_gsis,
            MAX(CASE WHEN pd.deduction_name = 'computer_loan' THEN pd.amount END) AS computer_loan,
            MAX(CASE WHEN pd.deduction_name = 'rural_bank_loan' THEN pd.amount END) AS rural_bank_loan,
            MAX(CASE WHEN pd.deduction_name = 'fcb_loan' THEN pd.amount END) AS fcb_loan,
            MAX(CASE WHEN pd.deduction_name = 'gfal' THEN pd.amount END) AS gfal,
            MAX(CASE WHEN pd.deduction_name = 'sss_premium' THEN pd.amount END) AS sss_premium,
            MAX(CASE WHEN pd.deduction_name = 'mpl_lite_loan' THEN pd.amount END) AS mpl_lite_loan,
            MAX(CASE WHEN pd.deduction_name = 'policy_loan' THEN pd.amount END) AS policy_loan
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            LEFT JOIN payroll_deductions pd ON p.id = pd.payroll_id
        `;
        let rows = [];
        if (periodParam && periodParam.length === 7) {
            query += ` WHERE strftime('%Y-%m', p.period_start) = ? GROUP BY p.id ORDER BY d.name ASC, e.name ASC`;
            rows = db.prepare(query).all(periodParam);
        } else {
            query += ` GROUP BY p.id ORDER BY d.name ASC, e.name ASC`;
            rows = db.prepare(query).all();
        }

        // Group rows by Dept
        const byDept = {};
        for (const row of rows) {
            const dept = row.department || 'Unassigned';
            if (!byDept[dept]) byDept[dept] = [];
            byDept[dept].push(row);
        }

        const workbook = new ExcelJS.Workbook();
        
        const thickBorder = { style: 'thin' }; 
        const allBorders = {
            top: thickBorder, left: thickBorder, bottom: thickBorder, right: thickBorder
        };

        const createSheet = (deptName, items) => {
            const safeName = deptName.substring(0, 31).replace(/[\\*?:/\[\]]/g, '');
            const ws = workbook.addWorksheet(safeName, { pageSetup: { orientation: 'landscape' } });
            
            // Default Font for entire sheet
            ws.properties.defaultRowHeight = 20;

            ws.mergeCells('A1:V1');
            ws.getCell('A1').value = 'GENERAL PAYROLL';
            ws.getCell('A1').font = { name: 'Arial', size: 12, bold: true };
            ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

            ws.mergeCells('A2:V2');
            ws.getCell('A2').value = 'LOCAL GOVERNMENT UNIT OF MAMBUSAO';
            ws.getCell('A2').font = { name: 'Arial', size: 10, bold: true };
            ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };

            ws.mergeCells('A3:V3');
            let displayPeriod = 'All Periods';
            if (periodParam && periodParam.length === 7) {
                const [year, month] = periodParam.split('-');
                const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
                displayPeriod = "Period: " + monthNames[parseInt(month) - 1] + " " + year;
            }
            ws.getCell('A3').value = displayPeriod;
            ws.getCell('A3').font = { name: 'Arial', size: 10, bold: true, underline: true };
            ws.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };

            ws.mergeCells('A6:V6');
            ws.getCell('A6').value = 'We acknowledge receipt of the sum shown opposite our names as full compensation for services rendered for the period stated.';
            ws.getCell('A6').font = { name: 'Arial', size: 10, italic: true };

            const deductStartCol = 6;
            const deductEndCol = 20; 

            ws.mergeCells(8, deductStartCol, 8, deductEndCol);
            ws.getCell(8, deductStartCol).value = 'D E D U C T I O N S';
            ws.getCell(8, deductStartCol).font = { name: 'Arial', size: 10, bold: true };
            ws.getCell(8, deductStartCol).alignment = { horizontal: 'center', vertical: 'middle' };
            ws.getCell(8, deductStartCol).border = allBorders;

            const headers = [
                'No.', 'Name', 'Position', 'Monthly Salary', 'Amount Earned',
                'Tax', 'GSIS/LR', 'PhilHealth', 'Pag-IBIG', 'MPL Lite', 'Emerg. Loan', 'Computer', 'Rural Bank', 'FCB 1', 'FCB 2', 'MPL GSIS', 'GFAL', 'SSS', 'MPL Lite Loan', 'Policy Loan',
                'Amount Received', 'Signature'
            ];
            
            const headerRow = ws.getRow(9);
            headers.forEach((h, i) => {
                const cell = headerRow.getCell(i + 1);
                cell.value = h;
                cell.font = { name: 'Arial', size: 10, bold: true };
                cell.border = allBorders;
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            });

            ws.getColumn(2).width = 25; 
            ws.getColumn(3).width = 20; 
            ws.getColumn(22).width = 20; 
            for (let i = 4; i <= 21; i++) ws.getColumn(i).width = 11;

            let rowIdx = 10;
            let counter = 1;

            items.forEach(r => {
                const dataRow = ws.getRow(rowIdx);
                
                dataRow.getCell(1).value = counter++;
                dataRow.getCell(2).value = r.name;
                dataRow.getCell(3).value = r.position;
                
                dataRow.getCell(4).value = r.monthly_salary;
                dataRow.getCell(4).numFmt = '#,##0.00';
                
                const earned = (r.basic_pay || 0) + (r.pera || 0) + (r.rata || 0) + (r.clothing_allowance || 0);
                dataRow.getCell(5).value = earned;
                dataRow.getCell(5).numFmt = '#,##0.00';

                const deds = [
                    r.tax, r.gsis, r.philhealth, r.pagibig, r.mpl_lite, r.emergency_loan,
                    r.computer_loan, r.rural_bank_loan, r.fcb_loan, r.fcb_loan_2, r.mpl_gsis, r.gfal,
                    r.sss_premium, r.mpl_lite_loan, r.policy_loan
                ];

                for (let i = 0; i < 15; i++) {
                    const c = dataRow.getCell(6 + i);
                    c.value = deds[i] || 0;
                    c.numFmt = '#,##0.00';
                }

                const earnedRef = dataRow.getCell(5).address;
                const dedStart = dataRow.getCell(6).address;
                const dedEnd = dataRow.getCell(20).address;
                
                const netCell = dataRow.getCell(21);
                netCell.value = { formula: `${earnedRef}-SUM(${dedStart}:${dedEnd})` };
                netCell.numFmt = '#,##0.00';
                netCell.font = { name: 'Arial', size: 10, bold: true };

                for (let i = 1; i <= 21; i++) {
                    const c = dataRow.getCell(i);
                    c.border = allBorders;
                    c.font = c.font || { name: 'Arial', size: 10 };
                }

                rowIdx++;
            });

            const totRow = ws.getRow(rowIdx);
            const titleCell = totRow.getCell(2);
            titleCell.value = 'TOTAL';
            titleCell.font = { name: 'Arial', size: 10, bold: true };
            titleCell.alignment = { horizontal: 'right' };
            
            totRow.getCell(1).border = allBorders;
            totRow.getCell(2).border = allBorders;
            totRow.getCell(3).border = allBorders;

            for (let i = 4; i <= 20; i++) {
                const startCell = ws.getCell(10, i).address;
                const endCell = ws.getCell(rowIdx - 1, i).address;
                const c = totRow.getCell(i);
                if (rowIdx > 10) {
                    c.value = { formula: `SUM(${startCell}:${endCell})` };
                } else {
                    c.value = 0;
                }
                c.numFmt = '#,##0.00';
                c.font = { name: 'Arial', size: 10, bold: true };
                c.border = allBorders;
            }
            totRow.getCell(21).border = allBorders;

            rowIdx += 3;
            ws.getCell(`B${rowIdx}`).value = 'CERTIFIED: Services rendered as stated above.';
            ws.getCell(`B${rowIdx}`).font = { name: 'Arial', size: 10 };
            ws.getCell(`H${rowIdx}`).value = 'APPROVED: For payment.';
            ws.getCell(`H${rowIdx}`).font = { name: 'Arial', size: 10 };
            ws.getCell(`P${rowIdx}`).value = 'VERIFIED: Funds available.';
            ws.getCell(`P${rowIdx}`).font = { name: 'Arial', size: 10 };
            
            rowIdx += 3;
            ws.getCell(`B${rowIdx}`).value = '_______________________________';
            ws.getCell(`B${rowIdx}`).font = { name: 'Arial', size: 10, bold: true };
            ws.getCell(`H${rowIdx}`).value = '_______________________________';
            ws.getCell(`H${rowIdx}`).font = { name: 'Arial', size: 10, bold: true };
            ws.getCell(`P${rowIdx}`).value = '_______________________________';
            ws.getCell(`P${rowIdx}`).font = { name: 'Arial', size: 10, bold: true };

            rowIdx += 1;
            ws.getCell(`B${rowIdx}`).value = 'Department Head';
            ws.getCell(`B${rowIdx}`).font = { name: 'Arial', size: 10 };
            ws.getCell(`H${rowIdx}`).value = 'Municipal Mayor';
            ws.getCell(`H${rowIdx}`).font = { name: 'Arial', size: 10 };
            ws.getCell(`P${rowIdx}`).value = 'Treasurer';
            ws.getCell(`P${rowIdx}`).font = { name: 'Arial', size: 10 };
        };

        if (Object.keys(byDept).length === 0) {
            workbook.addWorksheet('Empty');
        } else {
            for (const dept in byDept) {
                createSheet(dept, byDept[dept]);
            }
        }

        const dateStr = periodParam ? periodParam.replace(/-/g, '_') : 'All_Periods';
        const filename = `Payroll_${dateStr}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        
        await workbook.xlsx.write(res);
        res.end();

    } catch (err) {
        console.error(err);
        res.status(500).send('Excel Export failed: ' + err.message);
    }
});

app.get('/print_payslip', (req, res) => {
    const settings = getGlobalSettings();
    const id = req.query.id;
    const payroll = db.prepare(`
        SELECT p.*, e.name, e.position, d.name AS department, e.monthly_salary,
        MAX(CASE WHEN pd.deduction_name = 'tax' THEN pd.amount END) AS tax,
        MAX(CASE WHEN pd.deduction_name = 'gsis' THEN pd.amount END) AS gsis,
        MAX(CASE WHEN pd.deduction_name = 'philhealth' THEN pd.amount END) AS philhealth,
        MAX(CASE WHEN pd.deduction_name = 'pagibig' THEN pd.amount END) AS pagibig,
        MAX(CASE WHEN pd.deduction_name = 'mpl_lite' THEN pd.amount END) AS mpl_lite,
        MAX(CASE WHEN pd.deduction_name = 'emergency_loan' THEN pd.amount END) AS emergency_loan,
        MAX(CASE WHEN pd.deduction_name = 'mpl_gsis' THEN pd.amount END) AS mpl_gsis,
        MAX(CASE WHEN pd.deduction_name = 'computer_loan' THEN pd.amount END) AS computer_loan,
        MAX(CASE WHEN pd.deduction_name = 'rural_bank_loan' THEN pd.amount END) AS rural_bank_loan,
        MAX(CASE WHEN pd.deduction_name = 'fcb_loan' THEN pd.amount END) AS fcb_loan,
        MAX(CASE WHEN pd.deduction_name = 'fcb_loan_2' THEN pd.amount END) AS fcb_loan_2,
        MAX(CASE WHEN pd.deduction_name = 'gfal' THEN pd.amount END) AS gfal,
        MAX(CASE WHEN pd.deduction_name = 'sss_premium' THEN pd.amount END) AS sss_premium,
        MAX(CASE WHEN pd.deduction_name = 'mpl_lite_loan' THEN pd.amount END) AS mpl_lite_loan,
        MAX(CASE WHEN pd.deduction_name = 'policy_loan' THEN pd.amount END) AS policy_loan
        FROM payroll p
        JOIN employees e ON p.employee_id = e.id
        LEFT JOIN departments d ON e.department_id = d.id
        LEFT JOIN payroll_deductions pd ON p.id = pd.payroll_id
        WHERE p.id = ?
        GROUP BY p.id
    `).get(id);

    if (!payroll) return res.send('Record not found');
    res.render('print_payslip', { settings, row: payroll });
});

const portfinder = require('portfinder');

// Start Server with dynamic port
portfinder.basePort = 3000;
portfinder.getPort((err, port) => {
    if (err) {
        console.error('Could not find an open port:', err);
        return;
    }
    app.listen(port, '0.0.0.0', () => {
        const url = `http://localhost:${port}`;
        console.log(`Server running at ${url}`);
        // Create a file to notify the user of the port
        fs.writeFileSync('server_port.txt', port.toString());
        fs.writeFileSync('server_pid.txt', process.pid.toString());
    });
});
