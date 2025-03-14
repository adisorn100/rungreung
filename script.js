const SPREADSHEET_ID = "1FTir6myrkrU1KP7AyqRaKaQiKmX2jKIx4iKVx28QM0w"; // Sheet ID
const SHEET_NAME = "ชีต1"; // ชื่อชีต
const TOKEN_URL = "https://script.google.com/macros/s/AKfycby3W1ArMO5Gyz_7WTeayqNYMyUYZF0d9nhps8XcviHj6fw_so6MLQG4DquWurA88KxVGQ/exec"; // Apps Script Token URL

// ✅ ดึง Access Token จาก Apps Script
async function getAccessToken() {
    try {
        const response = await fetch(TOKEN_URL);
        const token = await response.text();
        if (!token) throw new Error("❌ ไม่สามารถรับ Access Token");
        return token;
    } catch (error) {
        console.error("❌ เกิดข้อผิดพลาดในการรับ Access Token:", error);
        return null;
    }
}

// ✅ ฟังก์ชันส่งข้อมูลไปยัง Google Sheets
async function sendData() {
    const accessToken = await getAccessToken();
    if (!accessToken) return;

    // 🔥 ตรวจสอบเบอร์โทรให้มีแค่ตัวเลขและต้องเป็น 10 หลัก
    let phoneNumber = document.getElementById("tl").value.trim();
    if (!/^\d{10}$/.test(phoneNumber)) {
        Swal.fire({ icon: "warning", title: "❌ เบอร์โทรไม่ถูกต้อง", text: "กรุณากรอกเบอร์โทร 10 หลัก" });
        return;
    }

    const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SHEET_NAME}!A:G:append?valueInputOption=RAW`;

    fetch(API_URL, {
        method: "POST",
        body: JSON.stringify({
            values: [[
                new Date().toLocaleString('th-TH'),
                document.getElementById("fn").value,
                phoneNumber,
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
    .then(() => {
        Swal.fire({
            icon: 'success',
            title: '✅ แจ้งซ่อมสำเร็จ',
            text: 'ข้อมูลถูกบันทึกแล้ว',
            timer: 2000
        });
        document.querySelector("form").reset();
    })
    .catch(error => {
        console.error("❌ Error:", error);
        Swal.fire({
            icon: "error",
            title: "❌ ไม่สามารถบันทึกข้อมูลได้",
            text: "โปรดตรวจสอบระบบอีกครั้ง"
        });
    });
}

// ✅ ฟังก์ชันค้นหาข้อมูลจาก Google Sheets
async function searchData() {
    const keyword = document.getElementById("keyword").value.trim().toLowerCase();
    const statusFilter = document.getElementById("statusFilter").value;
    const result = document.getElementById("result");

    result.innerHTML = "🔄 กำลังค้นหา...";

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

        // 🔥 ค้นหาจากคีย์เวิร์ดและตัวกรองสถานะ
        const filteredRows = rows.filter(row => {
            const matchesKeyword = row.slice(1, 4).some(cell => (cell || "").toLowerCase().includes(keyword));
            const matchesStatus = statusFilter ? row[5] === statusFilter : true;
            return matchesKeyword && matchesStatus;
        });

        if (filteredRows.length === 0) {
            result.innerHTML = `<div class="alert alert-warning">❌ ไม่พบข้อมูล</div>`;
            return;
        }

        let html = `<table class="table table-striped"><thead><tr>
                        <th>ชื่อ</th><th>เบอร์โทร</th><th>อุปกรณ์</th><th>อาการ</th><th>สถานะ</th><th>ผู้ปิดงาน</th><th>อัปเดต</th>
                    </tr></thead><tbody>`;

        filteredRows.forEach((row, index) => {
            let statusClass = row[5] === "เสร็จสิ้น" ? "badge bg-success" : "badge bg-danger";
            html += `<tr>
                        <td>${row[1] || "-"}</td>
                        <td>${row[2] || "-"}</td>
                        <td>${row[3] || "-"}</td>
                        <td>${row[4] || "-"}</td>
                        <td><span class="${statusClass}">${row[5] || "-"}</span></td>
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

#result 
{
    max-height: 400px;
    overflow-y: auto;
}

// ✅ ฟังก์ชันล้างการค้นหา
function clearSearch() {
    document.getElementById("keyword").value = "";
    document.getElementById("statusFilter").value = "";
    document.getElementById("result").innerHTML = "";
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
    .then(() => {
        Swal.fire({
            icon: 'success',
            title: '📌 ปิดงานสำเร็จ',
            text: 'ข้อมูลถูกอัปเดตแล้ว',
            timer: 2000
        });
        searchData();
    })
    .catch(error => {
        console.error("❌ Error:", error);
    });
}
