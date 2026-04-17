// Google Apps Script API for Employee Evaluation System
// Backend for Google Sheets integration

// =========================================================================
// CONFIGURATION
// =========================================================================
var CONFIG = {
  SHEET_MASTER: "Master_List",
  SHEET_DB: "DB_Penilaian New"
};

// =========================================================================
// CORS Helper
// =========================================================================
function createCORSResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  
  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  output.setHeader("Access-Control-Allow-Headers", "Content-Type");
  
  return output;
}

// =========================================================================
// 1. GET METHOD (Frontend minta daftar nama untuk di-render)
// Endpoint: URL_WEB_APP?action=getUsers
// =========================================================================
function doGet(e) {
  // Handle undefined event object
  if (!e || !e.parameter) {
    return createCORSResponse({
      status: "error",
      message: "Invalid request: no parameters provided"
    });
  }
  
  // Handle CORS preflight
  if (e.parameter.cors === "true") {
    return createCORSResponse({status: "ok"});
  }
  
  var action = e.parameter.action;
  var penilai = e.parameter.penilai; 
  
  if (action === "getUsers") {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
      
      if (!sheet) {
        return createCORSResponse({
          status: "error",
          message: "Master sheet not found"
        });
      }
      
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var users = [];
      
      // Find column indices
      var namaIdx = headers.indexOf("Nama");
      var posisiIdx = headers.indexOf("Posisi");
      var outletIdx = headers.indexOf("Outlet");
      var statusIdx = headers.indexOf("Status");
      
      // Skip header row
      for (var i = 1; i < data.length; i++) {
        var nama = data[i][namaIdx];
        var posisi = data[i][posisiIdx];
        var outlet = data[i][outletIdx] || "-";
        var status = data[i][statusIdx] || "";
        
        // Filter out blacklisted roles/outlets
        if (["Manager", "HRD", "Owner", "Head Office"].includes(posisi)) continue;
        if (["Head Office", "HO"].includes(outlet)) continue;
        
        // If penilai provided, filter users based on evaluator's role
        if (penilai) {
          var penilaiData = data.find(row => row[namaIdx] === penilai);
          if (penilaiData) {
            var penilaiPosisi = penilaiData[posisiIdx];
            var penilaiOutlet = penilaiData[outletIdx];
            
            // Manager can see all
            if (penilaiPosisi === "Manager") {
              // Show all
            } else if (penilaiPosisi === "SPV") {
              // SPV can see same outlet
              if (outlet !== penilaiOutlet) continue;
            } else {
              // Staff can only see themselves
              if (nama !== penilai) continue;
            }
          }
        }
        
        users.push({
          nama: nama,
          posisi: posisi,
          outlet: outlet,
          status: status
        });
      }
      
      return createCORSResponse({
        status: "success",
        data: users
      });
    } catch (error) {
      Logger.log("Error in getUsers: " + error.toString());
      return createCORSResponse({
        status: "error",
        message: error.toString()
      });
    }
  }
  
  if (action === "getEvaluations") {
    try {
      Logger.log("getEvaluations called");
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.SHEET_DB);
      
      if (!sheet) {
        Logger.log("Sheet not found: " + CONFIG.SHEET_DB);
        return createCORSResponse({
          status: "error",
          message: "Database sheet not found: " + CONFIG.SHEET_DB
        });
      }
      
      var data = sheet.getDataRange().getValues();
      Logger.log("Total rows in sheet: " + data.length);
      
      if (data.length <= 1) {
        Logger.log("No data in sheet (only header or empty)");
        return createCORSResponse({
          status: "success",
          data: []
        });
      }
      
      // Get header row
      var headers = data[0];
      Logger.log("Headers: " + JSON.stringify(headers));
      
      // Map Indonesian headers to field names
      var headerToField = {
        'tanggal': 'timestamp',
        'nama penilai': 'penilai',
        'karyawan yang dinilai': 'yangDinilai',
        'posisi': 'posisi',
        'outlet': 'outlet',
        'status': 'category',
        'komunikasi dengan rekan & atasan': 'ss1',
        'kerja sama tim': 'ss2',
        'tangung jawab & manajemen waktu': 'ss3',
        'inisiatif & penyelesaian masalah': 'ss4',
        'penguasaan tugas dan sop': 'hs1',
        'ketelitian & kecepatan kerja': 'hs2',
        'kemampuan menggunakan alat dan sistem': 'hs3',
        'konsistensi hasil kerja': 'hs4',
        'kedisiplinan dan kehadiran': 'at1',
        'kepatuhan aturan & arahan': 'at2',
        'etika & profesionalitas': 'at3',
        'tanggung jawab linkungan': 'at4',
        'ramah terhadap pelanggan': 'at5',
        'melaksanakan sholat': 'sholat',
        'melaksanakan puasa': 'puasa'
      };
      
      Logger.log("Header to field mapping: " + JSON.stringify(headerToField));
      
      var evaluations = [];
      
      // Skip header row, start from index 1
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        
        Logger.log("Processing row " + i + ": " + JSON.stringify(row));
        
        var evaluation = {
          timestamp: null,
          penilai: "",
          yangDinilai: "",
          posisi: "",
          outlet: "",
          category: "",
          ss1: "", ss2: "", ss3: "", ss4: "",
          hs1: "", hs2: "", hs3: "", hs4: "",
          at1: "", at2: "", at3: "", at4: "", at5: "",
          sholat: "",
          puasa: ""
        };
        
        // Map row data using Indonesian headers
        var mapped = false;
        for (var h = 0; h < headers.length; h++) {
          var header = headers[h] ? headers[h].toString().toLowerCase() : "";
          var fieldName = headerToField[header];
          if (fieldName && row[h]) {
            if (fieldName === 'timestamp') {
              evaluation[fieldName] = new Date(row[h]).toISOString();
            } else {
              evaluation[fieldName] = row[h].toString();
            }
            mapped = true;
          }
        }
        
        // If no headers matched, try position-based reading
        if (!mapped) {
          Logger.log("No headers matched, trying position-based reading");
          evaluation.timestamp = row[0] ? new Date(row[0]).toISOString() : null;
          evaluation.penilai = row[1] ? row[1].toString() : "";
          evaluation.yangDinilai = row[2] ? row[2].toString() : "";
          evaluation.posisi = row[3] ? row[3].toString() : "";
          evaluation.outlet = row[4] ? row[4].toString() : "";
          evaluation.category = row[5] ? row[5].toString() : "";
          evaluation.ss1 = row[6] ? row[6].toString() : "";
          evaluation.ss2 = row[7] ? row[7].toString() : "";
          evaluation.ss3 = row[8] ? row[8].toString() : "";
          evaluation.ss4 = row[9] ? row[9].toString() : "";
          evaluation.hs1 = row[10] ? row[10].toString() : "";
          evaluation.hs2 = row[11] ? row[11].toString() : "";
          evaluation.hs3 = row[12] ? row[12].toString() : "";
          evaluation.hs4 = row[13] ? row[13].toString() : "";
          evaluation.at1 = row[14] ? row[14].toString() : "";
          evaluation.at2 = row[15] ? row[15].toString() : "";
          evaluation.at3 = row[16] ? row[16].toString() : "";
          evaluation.at4 = row[17] ? row[17].toString() : "";
          evaluation.at5 = row[18] ? row[18].toString() : "";
          evaluation.sholat = row[19] ? row[19].toString() : "";
          evaluation.puasa = row[20] ? row[20].toString() : "";
        }
        
        Logger.log("Mapped evaluation: " + JSON.stringify(evaluation));
        
        // Skip empty rows based on mapped penilai
        if (!evaluation.penilai) {
          Logger.log("Row " + i + " skipped (empty penilai)");
          continue;
        }
        
        // If position/outlet not in row, get from master data
        if (!evaluation.posisi || !evaluation.outlet) {
          var masterSheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
          if (masterSheet) {
            var masterData = masterSheet.getDataRange().getValues();
            for (var j = 1; j < masterData.length; j++) {
              if (masterData[j][1] === evaluation.yangDinilai) {
                if (!evaluation.posisi) evaluation.posisi = masterData[j][2] || "";
                if (!evaluation.outlet) evaluation.outlet = masterData[j][3] || "";
                break;
              }
            }
          }
        }
        
        // Categorize position if not set
        if (!evaluation.category) {
          evaluation.category = categorizeByPosition(evaluation.posisi || '');
        }
        
        evaluations.push(evaluation);
      }
      
      Logger.log("Returning " + evaluations.length + " evaluations");
      
      return createCORSResponse({
        status: "success",
        data: evaluations
      });
    } catch (error) {
      Logger.log("Error in getEvaluations: " + error.toString());
      return createCORSResponse({
        status: "error",
        message: error.toString()
      });
    }
  }
  
  return createCORSResponse({
    status: "error",
    message: "Unknown action"
  });
}

// =========================================================================
// 2. POST METHOD (Frontend kirim data untuk disimpan)
// Endpoint: URL_WEB_APP (POST with action parameter)
// =========================================================================
function doPost(e) {
  try {
    var requestBody = e.postData.contents;
    var payload = JSON.parse(requestBody);
    
    var action = payload.action;
    
    if (action === "saveEvaluation") {
      return saveSingleEvaluation(payload.data);
    }
    
    if (action === "saveBatchEvaluations") {
      return saveBatchEvaluations(payload.data);
    }
    
    return createCORSResponse({
      status: "error",
      message: "Unknown action"
    });
  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    return createCORSResponse({
      status: "error",
      message: error.toString()
    });
  }
}

// =========================================================================
// Helper Functions
// =========================================================================
function saveSingleEvaluation(evaluation) {
  Logger.log("saveSingleEvaluation called with: " + JSON.stringify(evaluation));
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_DB);
  
  if (!sheet) {
    Logger.log("Creating new sheet: " + CONFIG.SHEET_DB);
    sheet = ss.insertSheet(CONFIG.SHEET_DB);
    
    // Add headers
    sheet.appendRow([
      "Tanggal", "Nama Penilai", "Karyawan yang dinilai", "Posisi", "Outlet", "Status",
      "Komunikasi dengan rekan & atasan", "Kerja sama tim", "Tangung jawab & manajemen waktu", "Inisiatif & penyelesaian masalah",
      "Penguasaan tugas dan SOP", "Ketelitian & kecepatan kerja", "Kemampuan menggunakan alat dan sistem", "Konsistensi hasil kerja",
      "Kedisiplinan dan kehadiran", "Kepatuhan aturan & arahan", "Etika & Profesionalitas", "Tanggung jawab linkungan", "Ramah terhadap pelanggan",
      "Melaksanakan sholat", "Melaksanakan Puasa"
    ]);
  }
  
  // Append row
  var row = [
    evaluation.timestamp || new Date().toISOString(),
    evaluation.penilai,
    evaluation.yangDinilai,
    evaluation.posisi,
    evaluation.outlet || "-",
    evaluation.category,
    evaluation.ss1, evaluation.ss2, evaluation.ss3, evaluation.ss4,
    evaluation.hs1, evaluation.hs2, evaluation.hs3, evaluation.hs4,
    evaluation.at1, evaluation.at2, evaluation.at3, evaluation.at4, evaluation.at5,
    evaluation.sholat,
    evaluation.puasa
  ];
  
  sheet.appendRow(row);
  
  Logger.log("Row appended successfully");
  
  return createCORSResponse({
    status: "success",
    message: "Evaluation saved successfully"
  });
}

function saveBatchEvaluations(records) {
  Logger.log("saveBatchEvaluations called with " + records.length + " records");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_DB);
  
  if (!sheet) {
    Logger.log("Creating new sheet: " + CONFIG.SHEET_DB);
    sheet = ss.insertSheet(CONFIG.SHEET_DB);
    
    // Add headers
    sheet.appendRow([
      "Tanggal", "Nama Penilai", "Karyawan yang dinilai", "Posisi", "Outlet", "Status",
      "Komunikasi dengan rekan & atasan", "Kerja sama tim", "Tangung jawab & manajemen waktu", "Inisiatif & penyelesaian masalah",
      "Penguasaan tugas dan SOP", "Ketelitian & kecepatan kerja", "Kemampuan menggunakan alat dan sistem", "Konsistensi hasil kerja",
      "Kedisiplinan dan kehadiran", "Kepatuhan aturan & arahan", "Etika & Profesionalitas", "Tanggung jawab linkungan", "Ramah terhadap pelanggan",
      "Melaksanakan sholat", "Melaksanakan Puasa"
    ]);
  }
  
  var rowsToInsert = [];
  
  for (var i = 0; i < records.length; i++) {
    var record = records[i];
    rowsToInsert.push([
      record.timestamp || new Date().toISOString(),
      record.penilai,
      record.yangDinilai,
      record.posisi,
      record.outlet || "-",
      record.category,
      record.ss1, record.ss2, record.ss3, record.ss4,
      record.hs1, record.hs2, record.hs3, record.hs4,
      record.at1, record.at2, record.at3, record.at4, record.at5,
      record.sholat,
      record.puasa
    ]);
  }
  
  if (rowsToInsert.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    Logger.log("Batch save completed: " + rowsToInsert.length + " rows inserted");
  }
  
  return createCORSResponse({
    status: "success",
    message: "Batch save completed successfully"
  });
}

function categorizeByPosition(position) {
  position = position ? position.toString().toLowerCase() : "";
  
  if (position.includes("spv") || position.includes("supervisor") || position.includes("manager")) {
    return "SPV";
  } else if (position.includes("karyawan") || position.includes("staff") || position.includes("crew")) {
    return "Karyawan";
  } else if (position.includes("freelance") || position.includes("part-time")) {
    return "Freelance";
  } else {
    return "Karyawan";
  }
}
