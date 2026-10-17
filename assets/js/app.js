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

document.getElementById('openSidebar').addEventListener('click', () => document.getElementById('sidebar').classList.add('open'));
document.getElementById('closeSidebar').addEventListener('click', () => document.getElementById('sidebar').classList.remove('open'));

// --- 2. Excel Upload & Parsing ---
const excelInput = document.getElementById('excelFile');
const uploadFeedback = document.getElementById('uploadFeedback');
const dataCheck = document.getElementById('dataCheck');

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
                    
                    // Aapke Excel format (Financial_Statements_2025-26 jaita.xlsx) ke mutabik:
                    // Row 0, 1, 2 are usually headers or titles (e.g., "आय व्यय विवरण", Dates, Column Names)
                    // Start reading data from row 3 (index 3)
                    for (let i = 3; i < json.length; i++) { 
                        const row = json[i];
                        // Check if row has some data
                        if (row && row.length > 0 && row.some(cell => cell !== undefined && cell !== null && cell !== "")) {
                            let obj = {};
                            keys.forEach((key, index) => { 
                                // Clean up NaN or undefined values, map array indexes to keys
                                obj[key] = (row[index] !== undefined && row[index] !== null && row[index] !== "NaN") ? row[index] : ""; 
                            });
                            objects.push(obj);
                        }
                    }
                }
                return objects;
            };

            // Read Accounts Sheets based on your format
            // Col 0: आय (Income), Col 1: राशि (Amount), Col 2: व्यय (Expense), Col 3: राशि (Amount)
            finalReportData.income_items = mapSheetToObjects("आय व्यय", ["income_name", "income_amount", "expense_name", "expense_amount"]);
            
            // Col 0: क्रय (Purchase), Col 1: राशि (Amount), Col 2: विक्रय (Sale), Col 3: राशि (Amount)
            finalReportData.trading_items = mapSheetToObjects("व्यापार खाता", ["purchase_name", "purchase_amount", "sale_name", "sale_amount"]);
            
            // Col 0: विवरण (Expenses), Col 1: राशि (Amount), Col 2: विवरण (Income), Col 3: राशि (Amount)
            finalReportData.pnl_items = mapSheetToObjects("लाभ हानि", ["exp_details", "exp_amount", "inc_details", "inc_amount"]);
            
            // Col 0: देनदारियाँ (Liabilities), Col 1,2: amounts, Col 3: लेनदारियाँ (Assets), Col 4,5: amounts
            // Adjust indexes if your columns are different. Assuming Col 0, 2 are names and 1, 5 are amounts for current year.
            finalReportData.bs_items = mapSheetToObjects("तुलन पत्र", ["liability_name", "liability_amount_prev", "liability_amount", "asset_name", "asset_amount_prev", "asset_amount"]);

            uploadFeedback.innerHTML = `<span style="color: var(--success);">✅ Accounts Data Parsed successfully! Proceed to 'Edit Profile'.</span>`;
            document.getElementById('statusText').innerText = "Excel Data Loaded";
            document.getElementById('statusText').className = "status-badge success";
            
            isExcelReady = true;
            dataCheck.innerText = "✅";
            checkReadyState();

        } catch (error) {
            uploadFeedback.innerHTML = `<span style="color: red;">❌ Error parsing Excel. Please ensure correct format.</span>`;
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
});

// --- 3. Form Save Logic ---
const formCheck = document.getElementById('formCheck');
document.getElementById('saveDataBtn').addEventListener('click', (e) => {
    e.preventDefault();
    const inputs = document.querySelectorAll('#gssForm input');
    inputs.forEach(input => {
        finalReportData[input.dataset.key] = input.value || "";
    });
    
    isFormReady = true;
    formCheck.innerText = "✅";
    alert("Profile Data Saved Successfully! You can now Generate the Report.");
    checkReadyState();
});

function checkReadyState() {
    if(isExcelReady && isFormReady) {
        document.getElementById('generateDocBtn').disabled = false;
    }
}

// --- 4. Word Document Generation (FIXED) ---
document.getElementById('generateDocBtn').addEventListener('click', () => {
    if (!isExcelReady || !isFormReady) {
        alert("Please complete both Excel upload and Profile Form.");
        return;
    }

    if (typeof window.PizZip === "undefined" || typeof window.docxtemplater === "undefined") {
        alert("System Error: Required libraries did not load properly. Check internet connection.");
        return;
    }

    const templateUrl = "templates/gss/report-template.docx";

    fetch(templateUrl)
        .then(response => {
            if (!response.ok) throw new Error("Could not find Word Template. Ensure it is uploaded at templates/gss/report-template.docx");
            return response.arrayBuffer();
        })
        .then(content => {
            const zip = new window.PizZip(content);
            const doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            // Set the data
            doc.render(finalReportData);

            const out = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });
            
            const samitiName = finalReportData["Samiti_Name"] ? finalReportData["Samiti_Name"].replace(/ /g, "_") : "Report";
            
            window.saveAs(out, `GSS_Audit_Report_${samitiName}.docx`);
        })
        .catch(error => {
            alert("Generation Error: " + error.message);
            console.error(error);
        });
});
