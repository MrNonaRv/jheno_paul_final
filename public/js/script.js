$(document).ready(function () {

    let additionFieldCount  = 0;
    let deductionFieldCount = 0;

    // ── Number formatter ───────────────────────────────────────────────────
    function fmt(n) {
        if (n === null || n === undefined || n === '' || parseFloat(n) === 0) return '0.00';
        const num = parseFloat(String(n).replace(/,/g, ''));
        if (isNaN(num)) return '0.00';
        return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    function rawNum(s) {
        if (s === null || s === undefined || s === '') return 0;
        const num = parseFloat(String(s).replace(/,/g, ''));
        return isNaN(num) ? 0 : num;
    }

    // ── Orientation toggle ─────────────────────────────────────────────────
    function setOrientation(o) {
        const m = '0.4in 0.3in 0.3in 0.3in';
        document.getElementById('printOrientationStyle').textContent =
            '@page { size: letter ' + o + '; margin: ' + m + '; }';
        $('#btnPortrait').toggleClass('active', o === 'portrait');
        $('#btnLandscape').toggleClass('active', o === 'landscape');
    }
    $('#btnPortrait').click(function ()  { setOrientation('portrait');  });
    $('#btnLandscape').click(function () { setOrientation('landscape'); });
    $('#printTableBtn').click(function () { window.print(); });

    // ── Calculate table footer totals ──────────────────────────────────────
    function calculateFooterTotals(tableId) {
        const table = document.getElementById(tableId);
        if (!table) return;
        const tfoot = table.querySelector('tfoot');
        if (!tfoot) return;
        
        const totalCells = tfoot.querySelectorAll('.total-cell');
        if (totalCells.length === 0) return;

        const tbody = table.querySelector('tbody');
        const rows = tbody.querySelectorAll('tr.payroll-row');
        
        // Initialize totals array
        const totals = new Array(totalCells.length).fill(0);
        
        // Sum all visible rows
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (row.style.display === 'none') continue;
            
            const cells = row.querySelectorAll('td.ec-num, td.cell-net');
            for (let j = 0; j < cells.length; j++) {
                if (j < totals.length) {
                    totals[j] += rawNum(cells[j].textContent);
                }
            }
        }
        
        // Update total cells
        for (let i = 0; i < totalCells.length; i++) {
            totalCells[i].textContent = fmt(totals[i]);
        }
    }
    
    // Recalculate totals on filter change
    function recalculateAllTotals() {
        calculateFooterTotals('payrollTable1');
        calculateFooterTotals('payrollTable2');
    }

    // ── Inline Editing ─────────────────────────────────────────────────────
    $(document).on('click', '.editable-cell', function (e) {
        if ($(this).hasClass('editing') || $(this).closest('.modal').length) return;

        const cell    = $(this);
        const isNum   = cell.hasClass('ec-num');
        const origText = cell.text().trim();
        const rawVal  = isNum ? rawNum(origText) : origText;

        cell.addClass('editing').empty();

        const input = $('<input class="cell-input">')
            .attr('type', isNum ? 'number' : 'text')
            .attr('step', isNum ? '0.01' : undefined)
            .val(isNum && rawVal !== 0 ? parseFloat(rawVal).toFixed(2) : (isNum ? '' : rawVal));

        cell.append(input);
        input.focus();
        if (isNum) input.select();

        input.on('blur', function () {
            const newVal  = $(this).val().trim();
            const changed = (newVal !== (isNum && rawVal !== 0 ? parseFloat(rawVal).toFixed(2) : String(rawVal)));

            if (isNum) {
                cell.text(newVal === '' ? '0.00' : fmt(newVal));
            } else {
                cell.text(newVal);
            }
            cell.removeClass('editing');

            if (!changed) return;

            cell.addClass('saving');
            $.ajax({
                url: '/api/update-cell', type: 'POST',
                data: { type: cell.data('type'), field: cell.data('field'), id: cell.data('id'), value: newVal },
                dataType: 'json',
                success: function (r) {
                    cell.removeClass('saving');
                    if (r.success) {
                        cell.addClass('saved');
                        setTimeout(function () { cell.removeClass('saved'); }, 900);
                        
                        // Update the amount received (net) for the row WITHOUT page refresh
                        const tr = cell.closest('tr');
                        const type = cell.data('type');
                        if (type === 'payroll' || type === 'employee') {
                            let gross = 0;
                            let deds = 0;
                            
                            // Logic depends on which sheet we are in
                            const tableId = cell.closest('table').attr('id');
                            if (tableId === 'payrollTable1') {
                                const basic = rawNum(tr.find('[data-field="basic_pay"]').text());
                                const pera = rawNum(tr.find('[data-field="pera"]').text());
                                const rata = rawNum(tr.find('[data-field="rata"]').text());
                                gross = basic + pera + rata;
                                
                                const tax = rawNum(tr.find('[data-field="tax"]').text());
                                const ph = rawNum(tr.find('[data-field="philhealth"]').text());
                                const lr = rawNum(tr.find('[data-field="gsis"]').text());
                                const pi = rawNum(tr.find('[data-field="pagibig"]').text());
                                const ml = rawNum(tr.find('[data-field="mpl_lite"]').text());
                                const emg = rawNum(tr.find('[data-field="emergency_loan"]').text());
                                const mg = rawNum(tr.find('[data-field="mpl_gsis"]').text());
                                const comp = rawNum(tr.find('[data-field="computer_loan"]').text());
                                const rb = rawNum(tr.find('[data-field="rural_bank_loan"]').text());
                                const fcb = rawNum(tr.find('[data-field="fcb_loan"]').text());
                                deds = tax + ph + lr + pi + ml + emg + mg + comp + rb + fcb;
                            } else if (tableId === 'payrollTable2') {
                                const basic = rawNum(tr.find('[data-field="basic_pay"]').text());
                                const cloth = rawNum(tr.find('[data-field="clothing_allowance"]').text());
                                gross = basic + cloth;
                                
                                const gfal = rawNum(tr.find('[data-field="gfal"]').text());
                                const sss = rawNum(tr.find('[data-field="sss_premium"]').text());
                                const mll = rawNum(tr.find('[data-field="mpl_lite_loan"]').text());
                                const mg = rawNum(tr.find('[data-field="mpl_gsis"]').text());
                                const fcb = rawNum(tr.find('[data-field="fcb_loan"]').text());
                                const fcb2 = rawNum(tr.find('[data-field="fcb_loan_2"]').text());
                                const pol = rawNum(tr.find('[data-field="policy_loan"]').text());
                                const comp = rawNum(tr.find('[data-field="computer_loan"]').text());
                                const rb = rawNum(tr.find('[data-field="rural_bank_loan"]').text());
                                const emg = rawNum(tr.find('[data-field="emergency_loan"]').text());
                                deds = gfal + sss + mll + mg + fcb + fcb2 + pol + comp + rb + emg;
                            }
                            
                            tr.find('.cell-net').text(fmt(gross - deds));
                        }
                        
                        // Recalculate totals for the current visible table
                        recalculateAllTotals();
                        
                        // Update fundsAmount in signatories based on visible rows
                        let totalVisibleNet = 0;
                        $('.payroll-row:visible').each(function() {
                            totalVisibleNet += rawNum($(this).find('.cell-net').text());
                        });
                        $('#fundsAmount').text(fmt(totalVisibleNet));
                    } else {
                        alert('Save failed: ' + (r.message || 'unknown error'));
                        cell.text(origText);
                    }
                },
                error: function () {
                    cell.removeClass('saving');
                    alert('Network error — cell could not be saved.');
                    cell.text(origText);
                }
            });
        });

        input.on('keydown', function (e) {
            if (e.key === 'Enter')  { $(this).blur(); }
            if (e.key === 'Escape') {
                cell.removeClass('editing').text(origText);
            }
        });
    });

    // ── Search & Filter ────────────────────────────────────────────────────
    function filterTable() {
        const term = document.getElementById('searchInput').value.toLowerCase().trim();
        const dept = document.getElementById('filterDepartment').value;
        const month = document.getElementById('filterMonth').value;
        let visibleNet = 0;
        
        const rows = document.querySelectorAll('.payroll-row');
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const nameMatch = !term || row.dataset.name.includes(term) || row.dataset.position.includes(term);
            const deptMatch = !dept || row.dataset.department === dept;
            const monthMatch = !month || String(row.dataset.month) === month;
            
            const day = parseInt(row.dataset.day);
            const isSheet1 = row.closest('#payrollTable1') !== null;
            const isSheet2 = row.closest('#payrollTable2') !== null;
            
            let dayMatch = true;
            if (isSheet1) dayMatch = day <= 15;
            else if (isSheet2) dayMatch = day >= 16;
            
            const visible = nameMatch && deptMatch && monthMatch && dayMatch;
            
            row.style.display = visible ? '' : 'none';
            
            if (visible) {
                const netCell = row.querySelector('.cell-net');
                if (netCell) {
                    visibleNet += rawNum(netCell.textContent);
                }
            }
        }
        
        // Update signatories amount based on filter
        document.getElementById('fundsAmount').textContent = fmt(visibleNet);
        
        // Update dynamic period display
        const dynamicDisplays = document.querySelectorAll('.dynamic-period-display');
        const sheet1Names = document.querySelectorAll('.sheet-name[data-id="sheet1"]');
        const sheet2Names = document.querySelectorAll('.sheet-name[data-id="sheet2"]');
        
        if (month) {
            const [year, mNum] = month.split('-');
            const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            const monthLabel = monthNames[parseInt(mNum) - 1] + " " + year;
            
            dynamicDisplays.forEach(el => el.textContent = 'Period: ' + monthLabel);
            sheet1Names.forEach(el => el.textContent = 'Part 1 - ' + monthLabel + ' (1-15)');
            sheet2Names.forEach(el => el.textContent = 'Part 2 - ' + monthLabel + ' (16-31)');
        } else {
            dynamicDisplays.forEach(el => el.textContent = 'All Periods');
            sheet1Names.forEach(el => el.textContent = 'Part 1 - All Periods');
            sheet2Names.forEach(el => el.textContent = 'Part 2 - All Periods');
        }
        
        // Recalculate footer totals
        recalculateAllTotals();
    }

    $('#searchBtn').click(filterTable);
    $('#filterDepartment').change(filterTable);
    $('#filterMonth').change(filterTable);
    $('#searchInput').keypress(function (e) { if (e.which === 13) filterTable(); });

    // ── Modal Handlers ─────────────────────────────────────────────────────
    $('#addRecordBtn').click(function () {
        $('#modalTitle').text('Add Payroll Record');
        $('#payroll_id').val('');
        $('#payrollForm')[0].reset();
        
        // Auto-fill dates with current 15-day period logic if possible, or leave blank
        const now = new Date();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const y = now.getFullYear();
        $('#period_start').val(`${y}-${m}-01`);
        $('#period_end').val(`${y}-${m}-15`);

        $('#addRecordForm').show();
    });

    $('#cancelBtn').click(function () {
        $('#addRecordForm').hide();
    });

    $('#employee_select').change(function () {
        const salary = $(this).find(':selected').data('salary');
        if (salary) {
            $('#monthly_salary').val(parseFloat(salary).toFixed(2));
            $('#amount_earned').val((parseFloat(salary) / 2).toFixed(2));
        } else {
            $('#monthly_salary').val('');
            $('#amount_earned').val('');
        }
    });

    $('#payrollForm').submit(function (e) {
        e.preventDefault();
        $.ajax({
            url: '/api/save-payroll', type: 'POST',
            data: $(this).serialize(), dataType: 'json',
            success: function (r) {
                if (r.success) { alert('Record saved!'); location.reload(); }
                else alert('Error: ' + r.message);
            },
            error: function () { alert('Error saving record.'); }
        });
    });

    // Import modal logic
    $('#openImportBtn').click(function () {
        $('#importModal').show();
        $('#importPeriodFields').hide();
        $('#importFileInput').val('');
        $('#importDropZone span').html('&#128190; Choose .sql or .xlsx file');
        $('#doImportBtn').prop('disabled', true);
    });
    $('#cancelImportBtn').click(function () {
        $('#importModal').hide();
    });
    $('#importFileInput').change(function() {
        if (this.files.length > 0) {
            const file = this.files[0];
            $('#doImportBtn').prop('disabled', false);
            $('#importDropZone span').text(file.name);
            
            if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
                $('#importPeriodFields').show();
            } else {
                $('#importPeriodFields').hide();
            }
        }
    });
    $('#doImportBtn').click(function() {
        const file = $('#importFileInput')[0].files[0];
        if (!file) return;
        const fd = new FormData();
        fd.append('backup_file', file);
        
        if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
            const month = $('#import_month').val();
            const year = $('#import_year').val();
            const part = $('#import_part').val();
            
            // Calculate representative dates for the backend
            let pStart, pEnd;
            if (part === '1') {
                pStart = `${year}-${month}-01`;
                pEnd = `${year}-${month}-15`;
            } else {
                pStart = `${year}-${month}-16`;
                // Get last day of month
                const lastDay = new Date(year, month, 0).getDate();
                pEnd = `${year}-${month}-${lastDay}`;
            }
            
            fd.append('period_start', pStart);
            fd.append('period_end', pEnd);
        }
        
        $(this).prop('disabled', true).text('Importing...');
        $.ajax({
            url: '/api/import-data', type: 'POST', data: fd,
            processData: false, contentType: false, dataType: 'json',
            success: function(r) {
                if (r.success) { alert('Import successful!'); location.reload(); }
                else { alert('Import failed: ' + r.message); $('#doImportBtn').prop('disabled', false).text('Import Data'); }
            },
            error: function() { alert('Network error during import.'); $('#doImportBtn').prop('disabled', false).text('Import Data'); }
        });
    });

    // Initialize footer totals on page load
    recalculateAllTotals();

});

// ── Global helpers ─────────────────────────────────────────────────────────
function exportExcel() {
    const period = $('#filterMonth').val();
    let url = '/api/export-excel';
    if (period) {
        url += '?period=' + encodeURIComponent(period);
    }
    window.location.href = url;
}

function deleteRecord(id) {
    if (!confirm('Delete this payroll record?')) return;
    $.post('/api/delete-payroll', { id: id }, function (r) {
        if (r.success) location.reload();
        else alert('Error: ' + r.message);
    }, 'json');
}

function editRecord(id) {
    $.ajax({
        url: '/api/get-payroll', type: 'GET', data: { id: id }, dataType: 'json',
        success: function (r) {
            if (!r.success) { alert('Error: ' + r.message); return; }
            const d = r.data;
            $('#modalTitle').text('Edit Payroll Record');
            $('#payroll_id').val(d.id);
            $('#employee_select').val(d.employee_id);
            $('#period_start').val(d.period_start || '');
            $('#period_end').val(d.period_end || '');
            $('#monthly_salary').val(parseFloat(d.monthly_salary).toFixed(2));
            $('#amount_earned').val(parseFloat(d.basic_pay).toFixed(2));
            $('#pera').val(parseFloat(d.pera).toFixed(2));
            $('#rata').val(parseFloat(d.rata).toFixed(2));
            $('#tax').val(parseFloat(d.tax).toFixed(2));
            $('#gsis').val(parseFloat(d.gsis).toFixed(2));
            $('#philhealth').val(parseFloat(d.philhealth).toFixed(2));
            $('#pagibig').val(parseFloat(d.pagibig).toFixed(2));
            $('#mpl_lite').val(parseFloat(d.mpl_lite).toFixed(2));
            $('#emergency_loan').val(parseFloat(d.emergency_loan).toFixed(2));
            $('#computer_loan').val(parseFloat(d.computer_loan).toFixed(2));
            $('#rural_bank_loan').val(parseFloat(d.rural_bank_loan).toFixed(2));
            $('#fcb_loan').val(parseFloat(d.fcb_loan).toFixed(2));
            $('#fcb_loan_2').val(parseFloat(d.fcb_loan_2).toFixed(2));
            $('#mpl_gsis').val(parseFloat(d.mpl_gsis).toFixed(2));
            $('#gfal').val(parseFloat(d.gfal).toFixed(2));
            $('#sss_premium').val(parseFloat(d.sss_premium).toFixed(2));
            $('#mpl_lite_loan').val(parseFloat(d.mpl_lite_loan).toFixed(2));
            $('#policy_loan').val(parseFloat(d.policy_loan).toFixed(2));
            $('#addRecordForm').show();
        },
        error: function () { alert('Error fetching record.'); }
    });
}

function escHtml(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
