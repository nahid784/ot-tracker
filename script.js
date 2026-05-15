let basicSalary = 10000;
let currentOTDays = [];
let editingIndex = -1;

document.getElementById('basic').addEventListener('input', (e) => {
  basicSalary = parseFloat(e.target.value) || 10000;
  if (currentOTDays.length > 0) renderResult();
});

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

uploadArea.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length > 0) {
    processFile(e.target.files[0]);
  }
});

function processFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    currentOTDays = [];

    rows.forEach(row => {
      const completedRaw = parseFloat(row['COMPLETED']);
      if (!completedRaw) return;

      const totalHours = Math.floor(completedRaw);
      const status = (row['STATUS'] || '').toString().toUpperCase();
      const inTime = row['IN TIME'] || 'N/A';
      const outTime = row['OUT TIME'] || 'N/A';
      const date = row['DATE'];

      let payableOT = 0;
      let dayType = '';

      if (completedRaw >= 12) {
        const extra = totalHours - 9;
        const breaks = (totalHours >= 16) ? 1.0 : 0.5;
        payableOT = Math.max(0, extra - breaks);
        dayType = 'Normal OT';
      } else if (status.includes('HOLIDAY') && inTime !== 'N/A' && outTime !== 'N/A') {
        payableOT = Math.max(0, totalHours - 1.0);
        dayType = 'Holiday OT';
      }

      if (payableOT > 0) {
        currentOTDays.push({
          date: date,
          dayType: dayType,
          inTime: inTime,
          outTime: outTime,
          completed: completedRaw.toFixed(2),
          totalHours: totalHours,
          payableOT: payableOT
        });
      }
    });

    renderResult();
  };
  reader.readAsArrayBuffer(file);
}

// === SUPER ROBUST TIME CALCULATOR ===
function calculateHours(inTimeStr, outTimeStr) {
  if (!inTimeStr || !outTimeStr || inTimeStr === 'N/A' || outTimeStr === 'N/A') return 0;

  try {
    // Remove date part if present (e.g., "30-APR-26 06:49:25 AM" → "06:49:25 AM")
    inTimeStr = inTimeStr.toString().split(' ').slice(-2).join(' ').trim();
    outTimeStr = outTimeStr.toString().split(' ').slice(-2).join(' ').trim();

    const inDate = parseFlexibleTime(inTimeStr);
    const outDate = parseFlexibleTime(outTimeStr);

    if (!inDate || !outDate) return 0;

    let diff = (outDate - inDate) / (1000 * 60 * 60);

    if (diff < 0) diff += 24;

    return parseFloat(diff.toFixed(2));
  } catch (e) {
    console.error("Time calculation failed:", e);
    return 0;
  }
}

function parseFlexibleTime(timeStr) {
  if (!timeStr) return null;
  timeStr = timeStr.trim();

  const testDate = new Date(`1970-01-01 ${timeStr}`);
  if (!isNaN(testDate.getTime())) return testDate;

  // Try without seconds
  const withoutSec = timeStr.replace(/:\d{2}\s*(AM|PM)/i, ' $1');
  const test2 = new Date(`1970-01-01 ${withoutSec}`);
  if (!isNaN(test2.getTime())) return test2;

  return null;
}

function editDay(index) {
  editingIndex = index;
  const day = currentOTDays[index];

  const modalHTML = `
    <div class="modal" id="editModal" style="display:flex;">
      <div class="modal-content">
        <h2>Edit Day: ${day.date}</h2>
        
        <div class="modal-label">Check In Time</div>
        <input type="text" id="editInTime" class="modal-input" value="${day.inTime}" placeholder="06:49:25 AM">

        <div class="modal-label">Check Out Time</div>
        <input type="text" id="editOutTime" class="modal-input" value="${day.outTime}" placeholder="07:04:20 PM">

        <small style="color:#64748b;">Example: 06:49:25 AM or 7:04 PM</small>

        <div class="modal-actions">
          <button class="cancel-btn" onclick="closeModal()">Cancel</button>
          <button class="save-btn" onclick="saveEdit()">Save & Recalculate</button>
        </div>
      </div>
    </div>
  `;

  document.getElementById('result').insertAdjacentHTML('beforeend', modalHTML);
}

function closeModal() {
  const modal = document.getElementById('editModal');
  if (modal) modal.remove();
  editingIndex = -1;
}

function saveEdit() {
  if (editingIndex === -1) return;

  const newInTime = document.getElementById('editInTime').value.trim();
  const newOutTime = document.getElementById('editOutTime').value.trim();

  const day = currentOTDays[editingIndex];

  day.inTime = newInTime || 'N/A';
  day.outTime = newOutTime || 'N/A';

  if (newInTime && newOutTime && newInTime !== 'N/A' && newOutTime !== 'N/A') {
    const calculatedHours = calculateHours(newInTime, newOutTime);
    
    day.completed = calculatedHours.toFixed(2);
    day.totalHours = Math.floor(calculatedHours);

    if (day.dayType === 'Normal OT') {
      const extra = day.totalHours - 9;
      const breaks = (day.totalHours >= 16) ? 1.0 : 0.5;
      day.payableOT = Math.max(0, extra - breaks);
    } else if (day.dayType === 'Holiday OT') {
      day.payableOT = Math.max(0, day.totalHours - 1.0);
    }
  }

  closeModal();
  renderResult();
}

function deleteDay(index) {
  if (confirm("Delete this day?")) {
    currentOTDays.splice(index, 1);
    renderResult();
  }
}

function renderResult() {
  let totalOTHours = 0;
  currentOTDays.forEach(day => totalOTHours += day.payableOT);

  const otRate = basicSalary / 104;
  const totalMoney = Math.round(totalOTHours * otRate);

  let html = `
    <div class="total-card">
      <h2>Total Payable Overtime</h2>
      <div class="total-amount">${totalOTHours.toFixed(1)} Hours</div>
      <h2 style="margin-top: 15px;">Total OT Amount</h2>
      <div class="total-amount">${totalMoney} Taka</div>
    </div>
  `;

  if (currentOTDays.length === 0) {
    html += `<p style="text-align:center; color:#64748b; font-size:1.1rem;">No overtime days found.</p>`;
  } else {
    html += `<table>
      <thead>
        <tr>
          <th>Date</th>
          <th>Type</th>
          <th>Check In</th>
          <th>Check Out</th>
          <th>Completed</th>
          <th>Floored</th>
          <th>Payable OT</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>`;

    currentOTDays.forEach((day, index) => {
      html += `
        <tr>
          <td><strong>${day.date}</strong></td>
          <td><strong>${day.dayType}</strong></td>
          <td>${day.inTime}</td>
          <td>${day.outTime}</td>
          <td>${day.completed}</td>
          <td>${day.totalHours}</td>
          <td><strong style="color:#10b981;">${day.payableOT.toFixed(1)}</strong></td>
          <td>
            <button onclick="editDay(${index})" style="background:#3b82f6; color:white; border:none; padding:6px 12px; border-radius:8px; margin-right:5px; cursor:pointer;">
              <i class="fas fa-edit"></i> Edit
            </button>
            <button class="delete-btn" onclick="deleteDay(${index})">
              <i class="fas fa-trash"></i>
            </button>
          </td>
        </tr>`;
    });

    html += `</tbody></table>`;
  }

  document.getElementById('result').innerHTML = html;
}
