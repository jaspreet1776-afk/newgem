// Global Storage
let finalReportData = {};
let isExcelReady = false;
let isFormReady = false;

// --- 1. Navigation Logic ---
document.querySelectorAll('.menu-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        document.querySelectorAll('.menu-item').forEach(i => i.classList.remove('active'));
        e.target.classList.add('active');
        
        const targetId = e.target.getAttribute('data-target');
        document.querySelectorAll('.view-section').forEach(sec => sec.classList.add('hidden'));
        document.getElementById(targetId).classList.remove('hidden');

        document.getElementById('sidebar').classList.remove('open');
    });
});

document.getElementById('openSidebar').addEventListener('click', () => {
    document.getElementById('sidebar').classList.add('open');
});
document.getElementById('closeSidebar').addEventListener('click', () => {
    document.getElementById('sidebar').classList.remove('open');
});

// --- 2. Excel Upload & Parsing ---
const excelInput = document.getElementById('excelFile');
const uploadFeedback = document.getElementById('uploadFeedback');

excelInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Function to map excel sheet rows into objects for Word template loops
            const mapSheetToObjects = (sheetName, keys) => {
                let objects = [];
                if (workbook.SheetNames.includes(sheetName)) {
                    const ws = workbook.Sheets[sheetName];
                    const json = XLSX.utils.sheet_to_json(ws, {header: 1});
                    for (let i = 2; i < json.length; i++) { // Skip headers (assuming row 0 & 1 are headers/titles)
                        const row = json[i];
                        if (row.length > 0 && row.some(cell => cell !== undefined && cell !== "")) {
                            let obj = {};
                            keys.forEach((key, index) => {
                                obj[key] = row[index] || "";
                            });
                            objects.push(obj);
                        }
                    }
                }
                return objects;
            };

            // Read Accounts Sheets (Based on your Excel structure)
            finalReportData.income_items = mapSheetToObjects("आय व्यय", ["income_name", "income_amount", "expense_name", "expense_amount"]);
            finalReportData.trading_items = mapSheetToObjects("व्यापार खाता", ["purchase_name", "purchase_amount", "sale_name", "sale_amount"]);
            finalReportData.pnl_items = mapSheetToObjects("लाभ हानि", ["exp_details", "exp_amount", "inc_details", "inc_amount"]);
            finalReportData.bs_items = mapSheetToObjects("तुलन पत्र", ["liability_name", "liability_amount", "asset_name", "asset_amount"]);

            uploadFeedback.innerHTML = `<span style="color: green;">✅ Accounts Data Parsed! Now go to 'Edit Details' to fill profile info.</span>`;
            document.getElementById('statusText').innerText = "Excel Data Loaded.";
            
            isExcelReady = true;
            document.getElementById('dataStatus').innerHTML = "✅ Ready";
            checkReadyState();

        } catch (error) {
            uploadFeedback.innerHTML = `<span style="color: red;">❌ Error parsing Excel: Check file format.</span>`;
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
});

// --- 3. Form Save Logic ---
document.getElementById('saveDataBtn').addEventListener('click', () => {
    const inputs = document.querySelectorAll('#gssForm input');
    inputs.forEach(input => {
        finalReportData[input.dataset.key] = input.value || "";
    });
    
    isFormReady = true;
    document.getElementById('formStatus').innerHTML = "✅ Saved";
    alert("Profile Data Saved! Proceed to 'Generate Report'.");
    checkReadyState();
});

function checkReadyState() {
    if(isExcelReady && isFormReady) {
        document.getElementById('generateDocBtn').disabled = false;
    }
}

// --- 4. Word Document Generation ---
document.getElementById('generateDocBtn').addEventListener('click', () => {
    if (!isExcelReady || !isFormReady) {
        alert("Please complete both Excel upload and Profile Form.");
        return;
    }

    const templateUrl = "templates/gss/report-template.docx";

    fetch(templateUrl)
        .then(response => {
            if (!response.ok) throw new Error("Template not found at templates/gss/report-template.docx");
            return response.arrayBuffer();
        })
        .then(content => {
            const zip = new PizZip(content);
            const doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            doc.render(finalReportData);

            const out = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });
            
            const samitiName = finalReportData["Samiti_Name"] ? finalReportData["Samiti_Name"].replace(/ /g, "_") : "Report";
            saveAs(out, `GSS_Audit_${samitiName}.docx`);
        })
        .catch(error => {
            alert("Generation Error: " + error.message);
            console.error(error);
        });
});
