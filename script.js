let basicSalary = 10000;
let currentOTDays = [];

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
        // Normal Working Day OT
        const extra = totalHours - 9;
        const breaks = (totalHours >= 16) ? 1.0 : 0.5;
        payableOT = Math.max(0, extra - breaks);
        dayType = 'Normal OT';
      } 
      else if (status.includes('HOLIDAY') && inTime !== 'N/A' && outTime !== 'N/A') {
        // Holiday Overtime
        payableOT = Math.max(0, totalHours - 1.0);   // 1 hour total break
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

function deleteDay(index) {
  currentOTDays.splice(index, 1);
  renderResult();
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