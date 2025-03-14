const SPREADSHEET_ID = "1FTir6myrkrU1KP7AyqRaKaQiKmX2jKIx4iKVx28QM0w"; // ใส่ Sheet ID
const SHEET_NAME = "ชีต1"; // ชื่อชีตของคุณ
const TOKEN_URL = "https://script.google.com/macros/s/AKfycby3W1ArMO5Gyz_7WTeayqNYMyUYZF0d9nhps8XcviHj6fw_so6MLQG4DquWurA88KxVGQ/exec"; // 🔥 ใส่ URL ของ Apps Script

// ✅ ดึง Access Token จาก Google Apps Script
async function getAccessToken() {
    const response = await fetch(TOKEN_URL);
    const token = await response.text();
    return token;
}

// ✅ ฟังก์ชันส่งข้อมูลไปยัง Google Sheets
async function sendData() {
    const accessToken = await getAccessToken();
    if (!accessToken) return;

    const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SHEET_NAME}!A:G:append?valueInputOption=RAW`;

    fetch(API_URL, {
        method: "POST",
        body: JSON.stringify({
            values: [[
                new Date().toLocaleString('th-TH'),
                document.getElementById("fn").value,
                document.getElementById("tl").value,
                document.getElementById("rm").value,
                document.getElementById("am").value,
                "กำลังดำเนินการ",
                ""
            ]]
        }),
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        }
    })
    .then(res => res.json())
    .then(response => {
        Swal.fire({
            icon: 'success',
            title: 'แจ้งซ่อมสำเร็จ',
            text: 'ข้อมูลถูกบันทึกแล้ว',
            timer: 2000
        });
        document.querySelector("form").reset();
    })
    .catch(error => {
        console.error("❌ Error:", error);
    });
}

// ✅ ฟังก์ชันค้นหาข้อมูลจาก Google Sheets
async function searchData() {
    const keyword = document.getElementById("keyword").value.trim();
    const result = document.getElementById("result");
    const accessToken = await getAccessToken();
    if (!accessToken) return;

    result.innerHTML = "🔄 กำลังค้นหา...";

    const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SHEET_NAME}!A:G`;

    fetch(API_URL, {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        }
    })
    .then(res => res.json())
    .then(data => {
        console.log("📌 Data from API:", data);
        const rows = data.values || [];

        const filteredRows = rows.filter(row =>
            row.slice(1, 4).some(cell => (cell || "").toLowerCase().includes(keyword.toLowerCase()))
        );

        if (filteredRows.length === 0) {
            result.innerHTML = `<div class="alert alert-warning">❌ ไม่พบข้อมูล</div>`;
            return;
        }

        let html = `<table class="table table-striped"><thead><tr>
                        <th>ชื่อ</th><th>เบอร์โทร</th><th>อุปกรณ์</th><th>อาการ</th><th>สถานะ</th><th>ผู้ปิดงาน</th><th>อัปเดต</th>
                    </tr></thead><tbody>`;

        filteredRows.forEach((row, index) => {
            html += `<tr>
                        <td>${row[1] || "-"}</td>
                        <td>${row[2] || "-"}</td>
                        <td>${row[3] || "-"}</td>
                        <td>${row[4] || "-"}</td>
                        <td>${row[5] || "-"}</td>
                        <td>${row[6] || "-"}</td>
                        <td>
                            <button class="btn btn-success btn-sm" onclick="updateStatus(${index + 1}, prompt('ใส่ชื่อผู้ปิดงาน'))">✅ ปิดงาน</button>
                        </td>
                    </tr>`;
        });

        html += `</tbody></table>`;
        result.innerHTML = html;
    })
    .catch(error => {
        console.error("❌ Error:", error);
        result.innerHTML = `<div class="alert alert-danger">❌ เกิดข้อผิดพลาดในการค้นหา</div>`;
    });
}

// ✅ ฟังก์ชันอัปเดตสถานะเป็น "เสร็จสิ้น" พร้อมชื่อผู้ปิดงาน
async function updateStatus(rowNumber, staffName) {
    if (!staffName) return;
    const accessToken = await getAccessToken();
    if (!accessToken) return;

    const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SHEET_NAME}!F${rowNumber + 1}:G${rowNumber + 1}?valueInputOption=USER_ENTERED`;

    fetch(API_URL, {
        method: "PUT",
        body: JSON.stringify({
            values: [["เสร็จสิ้น", staffName]]
        }),
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        }
    })
    .then(res => res.json())
    .then(response => {
        Swal.fire({
            icon: 'success',
            title: '📌 ปิดงานสำเร็จ',
            text: 'ข้อมูลถูกอัปเดตแล้ว',
            timer: 2000
        });
        searchData(); // รีเฟรชผลการค้นหา
    })
    .catch(error => {
        console.error("❌ Error:", error);
    });
}

	// ✅ ฟังก์ชัน Export ข้อมูลเป็น Excel
async function exportToExcel() {
    const accessToken = await getAccessToken();
    if (!accessToken) return;

    const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SHEET_NAME}!A:G`;

    fetch(API_URL, {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        }
    })
    .then(res => res.json())
    .then(data => {
        const rows = data.values || [];

        if (rows.length === 0) {
            Swal.fire({
                icon: "warning",
                title: "❌ ไม่มีข้อมูล",
                text: "ไม่พบข้อมูลสำหรับดาวน์โหลด"
            });
            return;
        }

        // ✅ สร้างไฟล์ Excel
        let ws = XLSX.utils.aoa_to_sheet(rows);
        let wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "ข้อมูลแจ้งซ่อม");

        // ✅ ดาวน์โหลดไฟล์
        XLSX.writeFile(wb, "repair_data.xlsx");
    })
    .catch(error => {
        console.error("❌ Error exporting data:", error);
        Swal.fire({
            icon: "error",
            title: "❌ ดาวน์โหลดล้มเหลว",
            text: "มีข้อผิดพลาดในการดาวน์โหลดไฟล์"
        });
    });
}


