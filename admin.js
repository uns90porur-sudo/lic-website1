// Global array to hold the parsed Excel data
let excelData = [];

let currentQuickFilter = 'all';

document.getElementById('excel-file').addEventListener('change', handleFile, false);
document.getElementById('search-input').addEventListener('input', applyFilters, false);
document.getElementById('filter-month').addEventListener('change', applyFilters, false);
document.getElementById('sort-by').addEventListener('change', applyFilters, false);

document.querySelectorAll('.quick-filter-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.quick-filter-btn').forEach(b => {
            b.style.background = 'transparent';
            b.style.color = 'var(--text-main)';
            b.style.border = '1px solid var(--border-color)';
        });
        e.target.style.background = 'var(--brand-blue)';
        e.target.style.color = 'white';
        e.target.style.border = 'none';
        
        currentQuickFilter = e.target.getAttribute('data-filter');
        applyFilters();
    });
});

document.getElementById('dark-mode-toggle').addEventListener('click', () => {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('darkMode', isDark);
    document.getElementById('dark-mode-toggle').innerHTML = isDark ? '<i class="fa-solid fa-sun"></i> Light Mode' : '<i class="fa-solid fa-moon"></i> Dark Mode';
});

window.addEventListener('DOMContentLoaded', () => {
    if (localStorage.getItem('darkMode') === 'true') {
        document.body.classList.add('dark-mode');
        document.getElementById('dark-mode-toggle').innerHTML = '<i class="fa-solid fa-sun"></i> Light Mode';
    }

    if (typeof AGENT_CONFIG !== 'undefined') {
        document.getElementById('cfg-agent-display').innerHTML = `<i class="fa-solid fa-user-tie"></i> ${AGENT_CONFIG.name} (${AGENT_CONFIG.agentCode})`;
        document.getElementById('cfg-pdf-agent-name').textContent = AGENT_CONFIG.name;
        document.getElementById('cfg-pdf-agent-code').textContent = AGENT_CONFIG.agentCode;
        document.getElementById('cfg-pdf-agent-phone').textContent = AGENT_CONFIG.phoneFormatted;
        document.getElementById('cfg-pdf-agent-sign').textContent = AGENT_CONFIG.name;
    }
});

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('file-name').innerHTML = `<i class="fa-solid fa-check-circle"></i> Loaded: ${file.name}`;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        
        try {
            // Read workbook
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Get first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            // We use header: 1 to get an array of arrays, so we can find where the actual table headers start
            const rawData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            processData(rawData);
            
            // Show dashboard
            document.getElementById('upload-card').style.display = 'none';
            document.getElementById('dashboard-content').classList.remove('hidden');
            
        } catch (error) {
            alert('Error parsing the Excel file. Please make sure it is a valid format.');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function processData(rows) {
    // The Excel file seems to have a few header rows (Agent Name, etc.) before the actual table header.
    // Let's find the row that contains 'PolicyNo' or 'S.No'
    let headerRowIndex = -1;
    
    for (let i = 0; i < rows.length; i++) {
        if (rows[i] && rows[i].some(cell => typeof cell === 'string' && (cell.includes('PolicyNo') || cell.includes('S.No')))) {
            headerRowIndex = i;
            break;
        }
    }
    
    if (headerRowIndex === -1) {
        alert("Could not find table headers in the Excel file.");
        return;
    }
    
    // Extract headers (remove empty columns if any)
    const headers = rows[headerRowIndex].map(h => h ? h.trim() : '');
    
    // Map data
    excelData = [];
    let totalPolicies = 0;
    let totalPremium = 0;
    let totalCommission = 0;
    
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || !row[0]) continue; // Skip empty rows
        
        // Map row array to object based on headers index
        const rowObj = {};
        for(let j = 0; j < headers.length; j++) {
            if (headers[j]) {
                rowObj[headers[j]] = row[j];
            }
        }
        
        // Only add valid policy rows
        if (rowObj['PolicyNo']) {
            totalPolicies++;
            
            // Parse premium and commission if they exist
            const prem = parseFloat(rowObj['TotPrem']) || parseFloat(rowObj['InstPrem']) || 0;
            const comm = parseFloat(rowObj['EstCom']) || 0;
            
            totalPremium += prem;
            totalCommission += comm;

            const name = (rowObj['Name of Assured'] || 'Unknown').toString().trim();
            const dobRaw = rowObj['DOB'] || rowObj['Date of Birth'] || rowObj['D.o.B'] || rowObj['Dob'] || '';
            const dobString = formatExcelDate(dobRaw);
            const dueString = (rowObj['Due'] || '').toString().trim();
            
            let docKey = Object.keys(rowObj).find(k => {
                let lower = k.toLowerCase().replace(/[^a-z]/g, '');
                return lower === 'doc' || lower === 'dateofcommencement' || lower === 'commencement' || lower === 'commencementdate';
            });
            const docRaw = docKey ? (rowObj[docKey] || '') : '';
            const docString = formatExcelDate(docRaw);
            
            let existingEntry = excelData.find(item => item.name === name);
            if (!existingEntry) {
                existingEntry = {
                    name: name,
                    count: 0,
                    policyNumbers: [],
                    policies: [],
                    totalPrem: 0,
                    totalCom: 0,
                    dob: dobString,
                    dueMonths: []
                };
                excelData.push(existingEntry);
            } else if (!existingEntry.dob && dobString) {
                existingEntry.dob = dobString;
            }
            
            // Extract Due Month (format usually MM/YYYY or similar)
            if (dueString) {
                const parts = dueString.split(/[-/]/);
                if (parts.length >= 1) {
                    let month = parts[0];
                    if (month.length === 1) month = '0' + month; // pad with zero
                    if (!existingEntry.dueMonths.includes(month)) {
                        existingEntry.dueMonths.push(month);
                    }
                }
            }
            
            existingEntry.count++;
            existingEntry.policyNumbers.push(rowObj['PolicyNo']);
            existingEntry.policies.push({ no: rowObj['PolicyNo'], prem: prem, doc: docString });
            existingEntry.totalPrem += prem;
            existingEntry.totalCom += comm;
        }
    }
    
    // Update Summary Cards
    document.getElementById('total-policies').textContent = totalPolicies;
    document.getElementById('total-premium').textContent = formatCurrency(totalPremium);
    document.getElementById('total-commission').textContent = formatCurrency(totalCommission);
    
    // Render Table
    renderTable(excelData);
    
    // Render Birthdays
    renderBirthdays(excelData);
    
    // Render Chart
    renderChart(totalPremium, totalCommission);
    
    // Save to LocalStorage
    saveDashboardData(totalPolicies, totalPremium, totalCommission);
}

function saveDashboardData(totalPolicies, totalPremium, totalCommission) {
    try {
        const dataToSave = {
            excelData: excelData,
            totalPolicies: totalPolicies,
            totalPremium: totalPremium,
            totalCommission: totalCommission
        };
        const jsonStr = JSON.stringify(dataToSave);
        // Encrypt the data using the current password
        const encrypted = CryptoJS.AES.encrypt(jsonStr, currentCryptoKey).toString();
        
        fetch('/api/saveData', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ payload: encrypted })
        }).then(response => response.json())
          .then(data => {
              if (data.success) {
                  console.log("Data securely synced to MongoDB.");
                  alert("Success! Data synced to the cloud and available on all devices.");
              } else {
                  console.error("Backend error:", data.error);
                  alert("Warning: Could not sync to cloud. Have you added MONGODB_URI to Vercel and redeployed?");
                  localStorage.setItem('licDashboardData', encrypted);
              }
          }).catch(err => {
              console.error("MongoDB sync error:", err);
              alert("Warning: Could not connect to backend server. Saving locally only.");
              localStorage.setItem('licDashboardData', encrypted);
          });
    } catch (e) {
        console.error("Local storage error:", e);
    }
}

let commissionChartInstance = null;
function renderChart(premium, commission) {
    const ctx = document.getElementById('commissionChart').getContext('2d');
    
    if (commissionChartInstance) {
        commissionChartInstance.destroy();
    }
    
    commissionChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Total Pending Premium', 'Estimated Total Commission'],
            datasets: [{
                label: 'Amount (₹)',
                data: [premium, commission],
                backgroundColor: [
                    'rgba(0, 51, 102, 0.7)',
                    'rgba(40, 167, 69, 0.7)'
                ],
                borderColor: [
                    'rgba(0, 51, 102, 1)',
                    'rgba(40, 167, 69, 1)'
                ],
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return '₹' + value;
                        }
                    }
                }
            },
            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });
}

function renderTable(data) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';
    
    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 30px;">No pending dues found.</td></tr>';
        return;
    }
    
    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        
        const name = row.name;
        const count = row.count;
        
        const policyNumbers = row.policies.map(p => {
            return p.doc ? `${p.no} <span style="font-size:0.8em; color:var(--text-muted);">(D.O.C: ${formatExcelDate(p.doc)})</span>` : p.no;
        }).join('<br>');
        
        const totPrem = formatCurrency(row.totalPrem);
        const estCom = formatCurrency(row.totalCom);
        
        let agentName = (typeof AGENT_CONFIG !== 'undefined') ? AGENT_CONFIG.name : 'R. Neelakandan';
        const message = `Hello ${name} Sir/Madam,\n\nThis is a gentle reminder from your LIC Advisor, ${agentName}.\n\nYou have ${count} pending LIC policies with a total premium due of ${totPrem}. (Policy Nos: ${row.policyNumbers.join(', ')}).\n\nPlease pay the premium at the earliest to keep your life cover active. You can securely pay online at https://licindia.in/\n\nIf you have any questions, feel free to contact me.\n\nThank you!`;
        const whatsappLink = `https://wa.me/?text=${encodeURIComponent(message)}`;
        
        const encodedRow = encodeURIComponent(JSON.stringify(row));
        
        tr.innerHTML = `
            <td>${index + 1}</td>
            <td class="policy-name" onclick="showClientModal('${encodedRow}')"><span class="clickable-name">${name}</span></td>
            <td><span class="fup-badge" style="background-color: var(--brand-blue); color: white;">${count}</span></td>
            <td><strong>${policyNumbers}</strong></td>
            <td class="amount text-red">${totPrem}</td>
            <td class="amount text-green">${estCom}</td>
            <td style="display: flex; gap: 5px;">
                <a href="${whatsappLink}" target="_blank" class="btn" style="background-color: #25D366; color: white; padding: 8px 12px; font-size: 0.85rem;"><i class="fa-brands fa-whatsapp"></i></a>
                <button onclick="generatePDFNotice('${encodeURIComponent(JSON.stringify(row))}')" class="btn btn-secondary" style="padding: 8px 12px; font-size: 0.85rem;"><i class="fa-solid fa-file-pdf"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function applyFilters() {
    const searchTerm = document.getElementById('search-input').value.toLowerCase();
    const filterMonth = document.getElementById('filter-month').value;
    const sortBy = document.getElementById('sort-by').value;
    
    // 1. Filter
    let filteredData = excelData.filter(row => {
        const nameMatch = row.name.toLowerCase().includes(searchTerm);
        const policyMatch = row.policyNumbers.join(' ').toLowerCase().includes(searchTerm);
        const searchMatch = !searchTerm || nameMatch || policyMatch;
        
        let monthMatch = true;
        if (filterMonth !== 'all') {
            monthMatch = row.dueMonths.includes(filterMonth);
        }
        
        let quickMatch = true;
        if (currentQuickFilter !== 'all') {
            const today = new Date();
            const thisMonth = (today.getMonth() + 1).toString().padStart(2, '0');
            
            if (currentQuickFilter === 'overdue') {
                quickMatch = row.dueMonths.some(m => parseInt(m) < parseInt(thisMonth));
            } else if (currentQuickFilter === 'due-today' || currentQuickFilter === 'due-7days') {
                quickMatch = row.dueMonths.includes(thisMonth);
            }
        }
        
        return searchMatch && monthMatch && quickMatch;
    });
    
    // 2. Sort
    if (sortBy === 'prem-desc') {
        filteredData.sort((a, b) => b.totalPrem - a.totalPrem);
    } else if (sortBy === 'prem-asc') {
        filteredData.sort((a, b) => a.totalPrem - b.totalPrem);
    } else if (sortBy === 'name-asc') {
        filteredData.sort((a, b) => a.name.localeCompare(b.name));
    }
    
    // 3. Render
    renderTable(filteredData);
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('en-IN', {
        style: 'currency',
        currency: 'INR',
        maximumFractionDigits: 0
    }).format(amount);
}

function formatExcelDate(serial) {
    if (!serial) return '';
    if (typeof serial === 'string') {
        if (!isNaN(serial) && serial.trim() !== '') {
            serial = parseFloat(serial);
        } else {
            return serial.toString().trim();
        }
    }
    if (typeof serial === 'number') {
        const offset = serial > 59 ? 25569 : 25568;
        const utc_days = Math.floor(serial - offset);
        const date = new Date(utc_days * 86400 * 1000);
        
        const day = date.getUTCDate().toString().padStart(2, '0');
        const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
        const year = date.getUTCFullYear();
        return `${day}/${month}/${year}`;
    }
    return serial.toString().trim();
}

function renderBirthdays(data) {
    const section = document.getElementById('birthdays-section');
    const tbody = document.getElementById('birthdays-table-body');
    tbody.innerHTML = '';
    
    const upcoming = [];
    const today = new Date();
    today.setHours(0,0,0,0);
    
    data.forEach(row => {
        if (!row.dob) return;
        
        let parts = row.dob.split(/[-/]/);
        if (parts.length >= 2) {
            let month = parseInt(parts[1], 10) - 1; // 0-11
            let day = parseInt(parts[0], 10);
            
            // Next birthday this year
            let nextBday = new Date(today.getFullYear(), month, day);
            
            // If already passed this year, next birthday is next year
            if (nextBday < today) {
                nextBday.setFullYear(today.getFullYear() + 1);
            }
            
            // Calculate difference in days
            const diffTime = nextBday - today;
            const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)); 
            
            if (diffDays >= 0 && diffDays <= 30) {
                upcoming.push({
                    name: row.name,
                    dob: row.dob,
                    diffDays: diffDays
                });
            }
        }
    });
    
    if (upcoming.length === 0) {
        section.classList.add('hidden');
        return;
    }
    
    section.classList.remove('hidden');
    
    // Sort by nearest
    upcoming.sort((a, b) => a.diffDays - b.diffDays);
    
    upcoming.forEach(person => {
        const tr = document.createElement('tr');
        
        let agentName = (typeof AGENT_CONFIG !== 'undefined') ? AGENT_CONFIG.name : 'your LIC Advisor';
        const message = `Happy Birthday ${person.name}! Wishing you a wonderful year ahead. Best wishes from ${agentName}.`;
        const whatsappLink = `https://wa.me/?text=${encodeURIComponent(message)}`;
        
        const daysText = person.diffDays === 0 ? 'Today!' : `In ${person.diffDays} days`;
        
        tr.innerHTML = `
            <td class="policy-name"><strong>${person.name}</strong> <span class="text-sm" style="color: var(--text-muted);">(${daysText})</span></td>
            <td>${person.dob}</td>
            <td><a href="${whatsappLink}" target="_blank" class="btn" style="background-color: #25D366; color: white; padding: 6px 12px; font-size: 0.8rem; border-radius: 6px;"><i class="fa-solid fa-gift"></i> Send Wish</a></td>
        `;
        tbody.appendChild(tr);
    });
}

function generatePDFNotice(encodedRowData) {
    const row = JSON.parse(decodeURIComponent(encodedRowData));
    
    // Populate Notice
    const notice = document.getElementById('pdf-notice');
    notice.classList.remove('hidden');
    
    document.getElementById('pdf-client-name').textContent = row.name;
    document.getElementById('pdf-total-amount').textContent = formatCurrency(row.totalPrem);
    
    const tbody = document.getElementById('pdf-policy-list');
    tbody.innerHTML = '';
    
    if (row.policies && row.policies.length > 0) {
        row.policies.forEach(policy => {
            const docText = policy.doc ? `<br><small style="color: #666; font-size: 0.85em;">D.O.C: ${formatExcelDate(policy.doc)}</small>` : '';
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding: 12px; border: 1px solid #ccc;">${policy.no}${docText}</td>
                <td style="padding: 12px; border: 1px solid #ccc; text-align: right;">${formatCurrency(policy.prem)}</td>
            `;
            tbody.appendChild(tr);
        });
    } else {
        row.policyNumbers.forEach(pNo => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding: 12px; border: 1px solid #ccc;">${pNo}</td>
                <td style="padding: 12px; border: 1px solid #ccc; text-align: right;">-</td>
            `;
            tbody.appendChild(tr);
        });
    }
    
    const opt = {
      margin:       10,
      filename:     `Premium_Due_Notice_${row.name.replace(/\s+/g, '_')}.pdf`,
      image:        { type: 'jpeg', quality: 0.98 },
      html2canvas:  { scale: 2, useCORS: true },
      jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    
    html2pdf().set(opt).from(notice).save().then(() => {
        notice.classList.add('hidden');
    });
}

let currentCryptoKey = ''; // Holds password for session
function handleLogin() {
    const user = document.getElementById('username').value;
    const pass = document.getElementById('password').value;
    const errorMsg = document.getElementById('login-error');
    
    // Check credentials
    if (user === 'admin' && pass === 'LIC123') {
        errorMsg.classList.add('hidden');
        currentCryptoKey = pass; // Save for encryption
        
        const loginWrapper = document.getElementById('login-wrapper');
        loginWrapper.style.opacity = '0';
        loginWrapper.style.visibility = 'hidden';
        loginWrapper.style.pointerEvents = 'none';
        
        setTimeout(() => {
            loginWrapper.classList.add('hidden');
        }, 500);
        
        const appWrapper = document.getElementById('app-wrapper');
        appWrapper.classList.remove('hidden');
        
        // NO sessionStorage! Must login every time.
        
        // Attempt to load encrypted data
        loadFromDatabase(pass);
    } else {
        errorMsg.classList.remove('hidden');
    }
}

function processEncryptedData(decryptedData) {
    const parsed = JSON.parse(decryptedData);
    excelData = parsed.excelData;
    
    // Re-render UI
    document.getElementById('total-policies').textContent = parsed.totalPolicies;
    document.getElementById('total-premium').textContent = formatCurrency(parsed.totalPremium);
    document.getElementById('total-commission').textContent = formatCurrency(parsed.totalCommission);
    
    renderTable(excelData);
    renderBirthdays(excelData);
    renderChart(parsed.totalPremium, parsed.totalCommission);
    
    // Hide upload screen, show dashboard
    document.getElementById('upload-card').style.display = 'none';
    document.getElementById('dashboard-content').classList.remove('hidden');
}

function loadFromDatabase(key) {
    fetch('/api/loadData', {
        method: 'GET'
    }).then(response => response.json())
      .then(data => {
          if (data.success && data.payload) {
              try {
                  const encrypted = data.payload;
                  const bytes = CryptoJS.AES.decrypt(encrypted, key);
                  const decryptedData = bytes.toString(CryptoJS.enc.Utf8);
                  if (!decryptedData) throw new Error("Decryption failed");
                  processEncryptedData(decryptedData);
              } catch(e) {
                  console.error("MongoDB data decryption failed. Incorrect password?", e);
              }
          } else {
              loadFromLocalStorage(key);
          }
      }).catch(err => {
          console.error("Error fetching from MongoDB:", err);
          loadFromLocalStorage(key);
      });
}

function loadFromLocalStorage(key) {
    try {
        const savedData = localStorage.getItem('licDashboardData');
        if (savedData) {
            // Decrypt using password
            const bytes = CryptoJS.AES.decrypt(savedData, key);
            const decryptedData = bytes.toString(CryptoJS.enc.Utf8);
            
            if (!decryptedData) throw new Error("Decryption failed");
            processEncryptedData(decryptedData);
        }
    } catch (e) {
        console.warn('No valid encrypted data found or decryption failed.');
        // Don't show error, just means they need to upload file
    }
}

function clearData() {
    if (confirm('Are you sure you want to clear the current list and upload a new one?')) {
        fetch('/api/clearData', {
            method: 'POST'
        }).catch(err => console.error("Error clearing MongoDB:", err));
        
        localStorage.removeItem('licDashboardData');
        excelData = [];
        
        // Reset UI
        document.getElementById('dashboard-content').classList.add('hidden');
        document.getElementById('upload-card').style.display = 'block';
        document.getElementById('file-name').innerHTML = '';
        document.getElementById('excel-file').value = '';
        document.getElementById('search-input').value = '';
        document.getElementById('filter-month').value = 'all';
        document.getElementById('sort-by').value = 'default';
        
        if (commissionChartInstance) {
            commissionChartInstance.destroy();
        }
        
        alert('Data cleared successfully. You can now upload a new Excel file.');
    }
}

function logout() {
    currentCryptoKey = ''; // Clear key memory
    
    // Show login screen
    const loginWrapper = document.getElementById('login-wrapper');
    loginWrapper.classList.remove('hidden');
    loginWrapper.style.opacity = '1';
    loginWrapper.style.visibility = 'visible';
    loginWrapper.style.pointerEvents = 'auto';
    
    // Hide app
    const appWrapper = document.getElementById('app-wrapper');
    appWrapper.classList.add('hidden');
    
    // Clear password input
    document.getElementById('password').value = '';
}

// Client Profile Modal Logic
function showClientModal(encodedRow) {
    const row = JSON.parse(decodeURIComponent(encodedRow));
    
    document.getElementById('modal-client-name').textContent = row.name;
    document.getElementById('modal-client-dob').textContent = row.dob || 'Unknown';
    document.getElementById('modal-client-total').textContent = formatCurrency(row.totalPrem);
    
    const tbody = document.getElementById('modal-policy-list');
    tbody.innerHTML = '';
    
    if (row.policies && row.policies.length > 0) {
        row.policies.forEach(p => {
            const docText = p.doc ? `<br><small style="color: #64748b;">D.O.C: ${formatExcelDate(p.doc)}</small>` : '';
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0;">${p.no}${docText}</td>
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0; text-align: right;">${formatCurrency(p.prem)}</td>
            `;
            tbody.appendChild(tr);
        });
    } else {
        row.policyNumbers.forEach(pNo => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0;">${pNo}</td>
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0; text-align: right;">-</td>
            `;
            tbody.appendChild(tr);
        });
    }
    
    // Setup WhatsApp Button
    let agentName = (typeof AGENT_CONFIG !== 'undefined') ? AGENT_CONFIG.name : 'R. Neelakandan';
    const message = `Hello ${row.name} Sir/Madam,\n\nThis is a gentle reminder from your LIC Advisor, ${agentName}.\n\nYou have ${row.count} pending LIC policies with a total premium due of ${formatCurrency(row.totalPrem)}. (Policy Nos: ${row.policyNumbers.join(', ')}).\n\nPlease pay the premium at the earliest to keep your life cover active.\n\nThank you!`;
    document.getElementById('modal-whatsapp-btn').href = `https://wa.me/?text=${encodeURIComponent(message)}`;
    
    // Show Modal
    const modal = document.getElementById('client-modal');
    modal.classList.remove('hidden');
    modal.style.pointerEvents = 'auto';
}

function closeClientModal() {
    const modal = document.getElementById('client-modal');
    modal.classList.add('hidden');
    modal.style.pointerEvents = 'none';
}

function exportToCSV() {
    if (excelData.length === 0) {
        alert("No data available to export.");
        return;
    }
    
    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += "Name,Policies Count,Policy Numbers,Total Premium,Estimated Commission\n";
    
    excelData.forEach(row => {
        const name = `"${row.name}"`;
        const count = row.count;
        const pNos = `"${row.policyNumbers.join('; ')}"`;
        const prem = row.totalPrem;
        const comm = row.totalCom;
        csvContent += `${name},${count},${pNos},${prem},${comm}\n`;
    });
    
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "LIC_Pending_Dues.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
