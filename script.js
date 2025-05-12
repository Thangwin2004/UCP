const technicalFactors = [
  ["Hệ thống phân tán", 2.0], ["Hiệu suất yêu cầu cao", 1.0],
  ["Hiệu quả xử lý giao dịch", 1.0], ["Khả năng sử dụng lại", 1.0],
  ["Dễ cài đặt", 0.5], ["Khả năng sử dụng", 0.5],
  ["Tính di động", 2.0], ["Khả năng thay đổi", 1.0],
  ["Giao tiếp đồng thời", 1.0], ["Tính bảo mật", 1.0],
  ["Truy cập bên thứ ba", 1.0], ["Yêu cầu đào tạo đặc biệt", 1.0],
  ["Hệ thống thân thiện", 1.0]
];

const envFactors = [
  ["Kinh nghiệm lập trình", 1.5], ["Kinh nghiệm ứng dụng", 0.5],
  ["Mức độ ổn định yêu cầu", 1.0], ["Kỹ năng phân tích", 0.5],
  ["Kinh nghiệm với mô hình UCP", 1.0], ["Sử dụng công cụ phát triển", 0.5],
  ["Mức độ phát triển theo chuẩn", 1.0], ["Sự hỗ trợ của khách hàng", 2.0]
];

function renderFactors(id, factors) {
  const tbody = document.getElementById(id);
  tbody.innerHTML = "";
  factors.forEach(([label, weight]) => {
    const row = document.createElement("tr");
    const inputClass = id === "technical-factors" ? "technical-factors-input" : "env-factors-input";
    row.innerHTML = `<td>${label}</td><td>${weight}</td>
    <td><input type="number" min="0" max="5" step="1" value="0" data-weight="${weight}" class="${inputClass}" /></td>`;
    tbody.appendChild(row);
  });
}

function sumWeightedValues(selector) {
  let total = 0;
  document.querySelectorAll(selector).forEach(input => {
    const val = parseFloat(input.value);
    const score = isNaN(val) ? 0 : val;
    const weight = parseFloat(input.dataset.weight);
    total += score * weight;
  });
  return total;
}

function updateAll() {
  const uawVal = parseFloat(document.getElementById("uaw").value);
  const uaw = isNaN(uawVal) ? 0 : uawVal;

  const uucwVal = parseFloat(document.getElementById("uucw").value);
  const uucw = isNaN(uucwVal) ? 0 : uucwVal;

  const uucp = uaw + uucw;
  document.getElementById("uucp").textContent = uucp.toFixed(2);

  const techTotal = sumWeightedValues(".technical-factors-input");
  document.getElementById("tech-total").textContent = techTotal.toFixed(2);
  const tcf = 0.6 + 0.01 * techTotal;
  document.getElementById("tcf").textContent = tcf.toFixed(2);

  const envTotal = sumWeightedValues(".env-factors-input");
  document.getElementById("env-total").textContent = envTotal.toFixed(2);
  const ef = 1.4 - 0.03 * envTotal;
  document.getElementById("ef").textContent = ef.toFixed(2);

  const ucp = uucp * tcf * ef;
  document.getElementById("ucp").textContent = ucp.toFixed(2);
}

function exportToExcel() {
  const wb = XLSX.utils.book_new();
  const ws_data = [];

  ws_data.push(["UAW", document.getElementById("uaw").value]);
  ws_data.push(["UUCW", document.getElementById("uucw").value]);

  ws_data.push(["Technical Factors"]);
  technicalFactors.forEach((item, i) => {
    const score = document.querySelectorAll(".technical-factors-input")[i].value;
    ws_data.push([item[0], item[1], score]);
  });

  ws_data.push(["Environmental Factors"]);
  envFactors.forEach((item, i) => {
    const score = document.querySelectorAll(".env-factors-input")[i].value;
    ws_data.push([item[0], item[1], score]);
  });

  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, "UCP Data");
  XLSX.writeFile(wb, "ucp_data.xlsx");
}

document.getElementById("excelFile").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    let techStart = -1, envStart = -1;
    let invalidEntries = [];

    function validateNumeric(value, fieldName, itemName) {
      if (value === undefined || value === null || value === "") return 0;
      if (typeof value === "number") return value;
      const num = parseFloat(value);
      if (isNaN(num)) {
        invalidEntries.push(`${fieldName}${itemName ? ' for "' + itemName + '"' : ''} has invalid non-numeric data: "${value}"`);
        return 0;
      }
      return num;
    }

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0]) continue;

      if (row[0] === "UAW") {
        const val = validateNumeric(row[1], "UAW");
        document.getElementById("uaw").value = val;
      }
      else if (row[0] === "UUCW") {
        const val = validateNumeric(row[1], "UUCW");
        document.getElementById("uucw").value = val;
      }
      else if (row[0] === "Technical Factors") {
        techStart = i + 1;
      }
      else if (row[0] === "Environmental Factors") {
        envStart = i + 1;
      }
    }

    if (techStart !== -1) {
      technicalFactors.forEach(([label], i) => {
        const score = rows[techStart + i]?.[2];
        const val = validateNumeric(score, "Technical Factor", label);
        document.querySelectorAll(".technical-factors-input")[i].value = val;
      });
    }

    if (envStart !== -1) {
      envFactors.forEach(([label], i) => {
        const score = rows[envStart + i]?.[2];
        const val = validateNumeric(score, "Environmental Factor", label);
        document.querySelectorAll(".env-factors-input")[i].value = val;
      });
    }

    if (invalidEntries.length > 0) {
      alert("Có lỗi dữ liệu trong file Excel:\n" + invalidEntries.join("\n"));
    }

    updateAll();
  };

  reader.readAsArrayBuffer(file);
});

// Initialization
renderFactors("technical-factors", technicalFactors);
renderFactors("env-factors", envFactors);

// Add event listeners
document.getElementById("uaw").addEventListener("input", updateAll);
document.getElementById("uucw").addEventListener("input", updateAll);

document.addEventListener("input", (e) => {
  if (
    e.target.classList.contains("technical-factors-input") ||
    e.target.classList.contains("env-factors-input")
  ) {
    updateAll();
  }
});

updateAll();
