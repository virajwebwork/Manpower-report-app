<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📊 Daily Manpower Report - FREESIA</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(45deg, #2c3e50, #34495e);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2rem;
            margin-bottom: 10px;
        }
        
        .stage {
            padding: 30px;
            border-bottom: 1px solid #eee;
        }
        
        .stage:last-of-type {
            border-bottom: none;
        }
        
        .stage h2 {
            color: #2c3e50;
            margin-bottom: 20px;
            font-size: 1.5rem;
        }
        
        .upload-area {
            border: 2px dashed #bdc3c7;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            background: #f8f9fa;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        
        .upload-area:hover {
            border-color: #3498db;
            background: #ebf3fd;
        }
        
        .upload-area.dragover {
            border-color: #2ecc71;
            background: #d5f4e6;
        }
        
        input[type="file"] {
            display: none;
        }
        
        .btn {
            background: linear-gradient(45deg, #3498db, #2980b9);
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
            margin: 10px 5px;
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(52, 152, 219, 0.4);
        }
        
        .btn-success {
            background: linear-gradient(45deg, #27ae60, #2ecc71);
        }
        
        .btn-success:hover {
            box-shadow: 0 5px 15px rgba(46, 204, 113, 0.4);
        }
        
        .table-container {
            margin: 20px 0;
            overflow-x: auto;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
        }
        
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #eee;
        }
        
        th {
            background: #f8f9fa;
            font-weight: 600;
            color: #2c3e50;
        }
        
        tr:hover {
            background: #f8f9fa;
        }
        
        .total-row {
            background: #e8f4f8 !important;
            font-weight: bold;
        }
        
        .alert {
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
        }
        
        .alert-info {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .alert-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .footer {
            background: #2c3e50;
            color: white;
            text-align: center;
            padding: 20px;
        }
        
        .footer a {
            color: #3498db;
        }
        
        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .stat-card {
            background: linear-gradient(45deg, #f39c12, #e67e22);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }
        
        .stat-number {
            font-size: 2rem;
            font-weight: bold;
            margin-bottom: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Daily Manpower Report - FREESIA</h1>
            <p>Process attendance data and generate group-wise manpower reports</p>
        </div>
        
        <!-- Stage 1 -->
        <div class="stage">
            <h2>Stage 1: Upload Original Attendance Sheet</h2>
            <div class="upload-area" onclick="document.getElementById('stage1-file').click()" id="upload1">
                <p>📥 Click here or drag & drop your Excel file</p>
                <p style="color: #7f8c8d; margin-top: 10px;">Supported: .xlsx files</p>
            </div>
            <input type="file" id="stage1-file" accept=".xlsx" onchange="processStage1(event)">
            
            <div id="stage1-results"></div>
        </div>
        
        <!-- Stage 2 -->
        <div class="stage">
            <h2>Stage 2: Group-wise Summary</h2>
            <div id="stage2-content">
                <div class="alert alert-info">
                    👆 Complete Stage 1 first to enable group-wise processing
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>Developed by: Viraj Niroshan Gunarathna</strong></p>
            <p>Contact: <a href="mailto:Viraj.se@gmail.com">Viraj.se@gmail.com</a> | 📞 0586804392</p>
            <p style="margin-top: 10px; color: #bdc3c7;">&copy; 2025 Viraj Niroshan Gunarathna. All rights reserved.</p>
        </div>
    </div>

    <script>
        let stage1Data = null;
        
        // Trade to group mapping
        const tradeGroupMap = {
            'Asst Electrician': 'ELE',
            'Electrician': 'ELE',
            'Ac Tech': 'HVAC',
            'Ac-Pipe-Fitter': 'HVAC',
            'Asst Ductman': 'HVAC',
            'Ductman': 'HVAC',
            'Chw-Pipe-Fitter': 'HVAC',
            'Gi Duct Fabricator': 'HVAC',
            'Insulator': 'HVAC',
            'Welder': 'Welder',
            'Asst Plumber': 'PLU',
            'Plumber': 'PLU',
            'Fire Alarm-Helper': 'FA',
            'Fire Alarm & Emergency Technician': 'FA',
            'Fire Alarm Technician': 'FA',
            'Fire Fighting Technician-Helper': 'FF',
            'Fire Fighting - Pipe Fitter': 'FF',
            'Fire Fighting Technicans': 'FF',
            'Fire Sealant Technician': 'F/S',
            'Elv Technician': 'ELV',
            'Lpg Technician-Pipe Fitter': 'LPG Technician/Pipe Fitter',
            'Lpg  Helper': 'LPG Helper',
            'Welder-Cs-Lpg-Technician': 'LPG Welder'
        };

        function processStage1(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const workbook = XLSX.read(e.target.result, { type: 'binary' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);

                    processAttendanceData(jsonData);
                } catch (error) {
                    showError('stage1-results', 'Error reading Excel file: ' + error.message);
                }
            };
            reader.readAsBinaryString(file);
        }

        function processAttendanceData(data) {
            try {
                // Clean and filter data
                const cleanData = data.filter(row => 
                    row['Working as'] && 
                    row['Building No'] && 
                    row['Status']
                );

                if (cleanData.length === 0) {
                    showError('stage1-results', 'No valid data found. Please check column names: "Working as", "Building No", "Status"');
                    return;
                }

                // Separate day and night data
                const dayData = cleanData.filter(row => 
                    row['Status'] && row['Status'].toString().toLowerCase().includes('day present')
                );
                
                const nightData = cleanData.filter(row => 
                    row['Status'] && row['Status'].toString().toLowerCase().includes('night present')
                );

                // Create pivot tables
                const dayPivot = createPivotTable(dayData);
                const nightPivot = createPivotTable(nightData);

                // Store for stage 2
                stage1Data = { dayPivot, nightPivot, dayData, nightData };

                // Display results
                displayStage1Results(dayPivot, nightPivot, dayData.length, nightData.length);
                
                // Enable stage 2
                enableStage2();

            } catch (error) {
                showError('stage1-results', 'Error processing data: ' + error.message);
            }
        }

        function createPivotTable(data) {
            const pivot = {};
            const buildings = new Set();

            // Collect all buildings and create structure
            data.forEach(row => {
                const trade = row['Working as'];
                const building = row['Building No'];
                buildings.add(building);
                
                if (!pivot[trade]) {
                    pivot[trade] = {};
                }
                if (!pivot[trade][building]) {
                    pivot[trade][building] = 0;
                }
                pivot[trade][building]++;
            });

            // Fill missing buildings with 0
            Object.keys(pivot).forEach(trade => {
                buildings.forEach(building => {
                    if (!pivot[trade][building]) {
                        pivot[trade][building] = 0;
                    }
                });
            });

            return { data: pivot, buildings: Array.from(buildings).sort() };
        }

        function displayStage1Results(dayPivot, nightPivot, dayCount, nightCount) {
            const resultsDiv = document.getElementById('stage1-results');
            
            resultsDiv.innerHTML = `
                <div class="stats">
                    <div class="stat-card">
                        <div class="stat-number">${dayCount}</div>
                        <div>Day Present</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">${nightCount}</div>
                        <div>Night Present</div>
                    </div>
                </div>
                
                <h3>📊 Day Present - Trade wise Building Count</h3>
                ${createPivotTableHTML(dayPivot)}
                
                <h3 style="margin-top: 30px;">🌙 Night Present - Trade wise Building Count</h3>
                ${createPivotTableHTML(nightPivot)}
                
                <button class="btn btn-success" onclick="downloadStage1Excel()">
                    ⬇️ Download Stage 1 Excel Report
                </button>
            `;
        }

        function createPivotTableHTML(pivot) {
            if (Object.keys(pivot.data).length === 0) {
                return '<div class="alert alert-info">No data available</div>';
            }

            let html = '<div class="table-container"><table><thead><tr>';
            html += '<th>Working as</th>';
            pivot.buildings.forEach(building => {
                html += `<th>${building}</th>`;
            });
            html += '<th>Total</th></tr></thead><tbody>';

            Object.keys(pivot.data).forEach(trade => {
                html += '<tr>';
                html += `<td><strong>${trade}</strong></td>`;
                let rowTotal = 0;
                pivot.buildings.forEach(building => {
                    const count = pivot.data[trade][building] || 0;
                    html += `<td>${count}</td>`;
                    rowTotal += count;
                });
                html += `<td><strong>${rowTotal}</strong></td>`;
                html += '</tr>';
            });

            // Add totals row
            html += '<tr class="total-row">';
            html += '<td><strong>Total</strong></td>';
            let grandTotal = 0;
            pivot.buildings.forEach(building => {
                let buildingTotal = 0;
                Object.keys(pivot.data).forEach(trade => {
                    buildingTotal += pivot.data[trade][building] || 0;
                });
                html += `<td><strong>${buildingTotal}</strong></td>`;
                grandTotal += buildingTotal;
            });
            html += `<td><strong>${grandTotal}</strong></td>`;
            html += '</tr></tbody></table></div>';

            return html;
        }

        function enableStage2() {
            const stage2Div = document.getElementById('stage2-content');
            stage2Div.innerHTML = `
                <div class="alert alert-info">
                    ✅ Stage 1 completed. Now processing group-wise summary...
                </div>
                <button class="btn" onclick="processStage2()">
                    📊 Generate Group-wise Summary
                </button>
                <div id="stage2-results"></div>
            `;
        }

        function processStage2() {
            if (!stage1Data) return;

            try {
                // Create group-wise summaries
                const dayGroupSummary = createGroupSummary(stage1Data.dayPivot);
                const nightGroupSummary = createGroupSummary(stage1Data.nightPivot);

                displayStage2Results(dayGroupSummary, nightGroupSummary);
            } catch (error) {
                showError('stage2-results', 'Error creating group summary: ' + error.message);
            }
        }

        function createGroupSummary(pivot) {
            const groupData = {};
            const buildings = pivot.buildings;

            // Initialize groups
            Object.values(tradeGroupMap).forEach(group => {
                if (!groupData[group]) {
                    groupData[group] = {};
                    buildings.forEach(building => {
                        groupData[group][building] = 0;
                    });
                }
            });

            // Sum trades into groups
            Object.keys(pivot.data).forEach(trade => {
                const group = tradeGroupMap[trade];
                if (group) {
                    buildings.forEach(building => {
                        groupData[group][building] += pivot.data[trade][building] || 0;
                    });
                }
            });

            return { data: groupData, buildings };
        }

        function displayStage2Results(dayGroupSummary, nightGroupSummary) {
            const resultsDiv = document.getElementById('stage2-results');
            
            resultsDiv.innerHTML = `
                <h3>📊 Day Present - Group wise Building Count</h3>
                ${createGroupTableHTML(dayGroupSummary)}
                
                <h3 style="margin-top: 30px;">🌙 Night Present - Group wise Building Count</h3>
                ${createGroupTableHTML(nightGroupSummary)}
                
                <button class="btn btn-success" onclick="downloadStage2Excel()">
                    📥 Download Complete Excel Report
                </button>
            `;
        }

        function createGroupTableHTML(groupSummary) {
            if (Object.keys(groupSummary.data).length === 0) {
                return '<div class="alert alert-info">No group data available</div>';
            }

            let html = '<div class="table-container"><table><thead><tr>';
            html += '<th>Main Group</th>';
            groupSummary.buildings.forEach(building => {
                html += `<th>${building}</th>`;
            });
            html += '<th>Total</th></tr></thead><tbody>';

            Object.keys(groupSummary.data).forEach(group => {
                html += '<tr>';
                html += `<td><strong>${group}</strong></td>`;
                let rowTotal = 0;
                groupSummary.buildings.forEach(building => {
                    const count = groupSummary.data[group][building] || 0;
                    html += `<td>${count}</td>`;
                    rowTotal += count;
                });
                html += `<td><strong>${rowTotal}</strong></td>`;
                html += '</tr>';
            });

            // Add totals row
            html += '<tr class="total-row">';
            html += '<td><strong>Total</strong></td>';
            let grandTotal = 0;
            groupSummary.buildings.forEach(building => {
                let buildingTotal = 0;
                Object.keys(groupSummary.data).forEach(group => {
                    buildingTotal += groupSummary.data[group][building] || 0;
                });
                html += `<td><strong>${buildingTotal}</strong></td>`;
                grandTotal += buildingTotal;
            });
            html += `<td><strong>${grandTotal}</strong></td>`;
            html += '</tr></tbody></table></div>';

            return html;
        }

        function downloadStage1Excel() {
            if (!stage1Data) return;
            
            const wb = XLSX.utils.book_new();
            
            // Add Day Present sheet
            const daySheet = createExcelSheet(stage1Data.dayPivot);
            XLSX.utils.book_append_sheet(wb, daySheet, 'Day_Present');
            
            // Add Night Present sheet
            const nightSheet = createExcelSheet(stage1Data.nightPivot);
            XLSX.utils.book_append_sheet(wb, nightSheet, 'Night_Present');
            
            XLSX.writeFile(wb, 'Manpower_Report-Day&Night.xlsx');
        }

        function downloadStage2Excel() {
            if (!stage1Data) return;
            
            const wb = XLSX.utils.book_new();
            
            // Add original sheets
            const daySheet = createExcelSheet(stage1Data.dayPivot);
            XLSX.utils.book_append_sheet(wb, daySheet, 'Day_Present');
            
            const nightSheet = createExcelSheet(stage1Data.nightPivot);
            XLSX.utils.book_append_sheet(wb, nightSheet, 'Night_Present');
            
            // Add group summaries
            const dayGroupSummary = createGroupSummary(stage1Data.dayPivot);
            const nightGroupSummary = createGroupSummary(stage1Data.nightPivot);
            
            const dayGroupSheet = createExcelSheet(dayGroupSummary);
            XLSX.utils.book_append_sheet(wb, dayGroupSheet, 'Day_Groupwise');
            
            const nightGroupSheet = createExcelSheet(nightGroupSummary);
            XLSX.utils.book_append_sheet(wb, nightGroupSheet, 'Night_Groupwise');
            
            XLSX.writeFile(wb, 'Manpower_Report_Complete.xlsx');
        }

        function createExcelSheet(pivotData) {
            const data = [];
            
            // Header row
            const header = ['Working as / Main Group', ...pivotData.buildings, 'Total'];
            data.push(header);
            
            // Data rows
            Object.keys(pivotData.data).forEach(trade => {
                const row = [trade];
                let rowTotal = 0;
                pivotData.buildings.forEach(building => {
                    const count = pivotData.data[trade][building] || 0;
                    row.push(count);
                    rowTotal += count;
                });
                row.push(rowTotal);
                data.push(row);
            });
            
            // Total row
            const totalRow = ['Total'];
            let grandTotal = 0;
            pivotData.buildings.forEach(building => {
                let buildingTotal = 0;
                Object.keys(pivotData.data).forEach(trade => {
                    buildingTotal += pivotData.data[trade][building] || 0;
                });
                totalRow.push(buildingTotal);
                grandTotal += buildingTotal;
            });
            totalRow.push(grandTotal);
            data.push(totalRow);
            
            return XLSX.utils.aoa_to_sheet(data);
        }

        function showError(elementId, message) {
            document.getElementById(elementId).innerHTML = `
                <div class="alert alert-error">❌ ${message}</div>
            `;
        }

        // Drag and drop functionality
        document.getElementById('upload1').addEventListener('dragover', function(e) {
            e.preventDefault();
            this.classList.add('dragover');
        });

        document.getElementById('upload1').addEventListener('dragleave', function(e) {
            e.preventDefault();
            this.classList.remove('dragover');
        });

        document.getElementById('upload1').addEventListener('drop', function(e) {
            e.preventDefault();
            this.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                document.getElementById('stage1-file').files = files;
                processStage1({ target: { files: files } });
            }
        });
    </script>
</body>
</html>