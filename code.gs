//------------ haha
// function xuatPDF_TungNguoi_VaoDungThuMuc() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheetNguon = ss.getSheetByName("danh sach ki ten print");
//   const sheetDich = ss.getSheetByName("danh sach lam them");

//   // 👉 CHỈNH 2 DÒNG DƯỚI ĐÂY MỖI LẦN CHẠY:
//   const hangBatDau = 331;    // dòng bắt đầu (bao gồm dòng này)
//   const hangKetThuc = 336;  // dòng kết thúc (bao gồm dòng này)

//   const folderGoc = getSubFolderByName(["google sheet", "hmsg", "chc", "ngoai vien", "coop pdf v2"]);

//   const data = sheetNguon.getRange(`A${hangBatDau}:E${hangKetThuc}`).getValues();

//   for (let i = 0; i < data.length; i++) {
//     const [stt, hoTen, ngaySinh, doiTuong, giaTriGoi] = data[i];

//     sheetDich.getRange("C10").setValue(stt || "");
//     sheetDich.getRange("C8").setValue(hoTen || "");
//     sheetDich.getRange("C9").setValue(ngaySinh || "");
//     sheetDich.getRange("C11").setValue(doiTuong || "");
//     sheetDich.getRange("D37").setValue(giaTriGoi || "");

//     SpreadsheetApp.flush();
//     Utilities.sleep(300);

//     const blob = xuatSheetThanhPDF(sheetDich);
//     blob.setName(`${stt || "no-stt"} - ${hoTen || "no-name"}.pdf`);
//     folderGoc.createFile(blob);

//     Utilities.sleep(3000);
//   }
// }
//------------




//------------
function xuatPDF_TungNguoi_VaoDungThuMuc_TuyChonHang() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNguon = ss.getSheetByName("danh sach ki ten print");
  const sheetDich = ss.getSheetByName("danh sach lam them");

  // 👉 CHỈNH DÒNG DƯỚI ĐỂ NHẬP CÁC HÀNG CẦN XUẤT
  const danhSachHang = [17,20,23,25,28,30,44,47,49,52,54,57,60,73,76,78,80,84,86,88,92,108,110,113,116,118,121,123,137,140,142];

  const folderGoc = getSubFolderByName(["google sheet", "hmsg", "chc", "ngoai vien", "coop pdf v2"]);

  for (let i = 0; i < danhSachHang.length; i++) {
    const hang = danhSachHang[i];
    const [stt, hoTen, ngaySinh, doiTuong, giaTriGoi] = sheetNguon.getRange(`A${hang}:E${hang}`).getValues()[0];

    sheetDich.getRange("C10").setValue(stt || "");
    sheetDich.getRange("C8").setValue(hoTen || "");
    sheetDich.getRange("C9").setValue(ngaySinh || "");
    sheetDich.getRange("C11").setValue(doiTuong || "");
    sheetDich.getRange("D37").setValue(giaTriGoi || "");

    SpreadsheetApp.flush();
    Utilities.sleep(300);

    const blob = xuatSheetThanhPDF(sheetDich);
    blob.setName(`${stt || "no-stt"} - ${hoTen || "no-name"}.pdf`);
    folderGoc.createFile(blob);

    Utilities.sleep(3000);
  }
}
//------------


function getSubFolderByName(pathArr) {
  let folder = DriveApp.getRootFolder();
  for (const name of pathArr) {
    const folders = folder.getFoldersByName(name);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      throw new Error("Không tìm thấy thư mục: " + name);
    }
  }
  return folder;
}

function xuatSheetThanhPDF(sheet) {
  const ss = sheet.getParent();
  const sheetId = sheet.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?`;

  const exportOptions = {
    format: "pdf",
    portrait: true,
    size: "A4",
    top_margin: 0.2,
    bottom_margin: 0.2,
    left_margin: 0.2,
    right_margin: 0.2,
    gridlines: false,
    printtitle: false,
    sheetnames: false,
    pagenumbers: false,
    fzr: false,
    gid: sheetId,
    fitw: true
  };

  const params = Object.entries(exportOptions)
    .map(([key, value]) => `${key}=${value}`)
    .join("&");

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url + params, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  return response.getBlob();
}
