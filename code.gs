function xuatPDF_TungNguoi_VaoDungThuMuc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNguon = ss.getSheetByName("danh sach ki ten print");
  const sheetDich = ss.getSheetByName("danh sach lam them");

  // 👉 CHỈNH 2 DÒNG DƯỚI ĐÂY MỖI LẦN CHẠY:
  const hangBatDau = 63;    // dòng bắt đầu (bao gồm dòng này)
  const hangKetThuc = 93;  // dòng kết thúc (bao gồm dòng này)

  const folderGoc = getSubFolderByName(["google sheet", "hmsg", "chc", "ngoai vien", "coop pdf"]);

  const data = sheetNguon.getRange(`A${hangBatDau}:E${hangKetThuc}`).getValues();

  for (let i = 0; i < data.length; i++) {
    const [stt, hoTen, ngaySinh, doiTuong, giaTriGoi] = data[i];
    if (!hoTen) continue;

    sheetDich.getRange("C11").setValue(stt);
    sheetDich.getRange("C9").setValue(hoTen);
    sheetDich.getRange("C10").setValue(ngaySinh);
    sheetDich.getRange("C12").setValue(doiTuong);
    sheetDich.getRange("D30").setValue(giaTriGoi);

    SpreadsheetApp.flush();
    Utilities.sleep(300);

    const blob = xuatSheetThanhPDF(sheetDich);
    blob.setName(`${stt} - ${hoTen}.pdf`);
    folderGoc.createFile(blob);

    Utilities.sleep(3000);
  }
}


// CHAY THEO CHỈ ĐỊNH HÀNG
//-------
// function xuatPDF_TungNguoi_VaoDungThuMuc_TuyChonHang() { 
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheetNguon = ss.getSheetByName("danh sach ki ten print");
//   const sheetDich = ss.getSheetByName("danh sach lam them");

//   // 👉 CHỈNH DÒNG DƯỚI ĐỂ NHẬP CÁC HÀNG CẦN XUẤT
//   const danhSachHang = [29,43,46,48,50,53,55,57,60,62]; // <== điền các số dòng bạn muốn chạy

//   const folderGoc = getSubFolderByName(["google sheet", "hmsg", "chc", "ngoai vien", "coop pdf"]);

//   for (let i = 0; i < danhSachHang.length; i++) {
//     const hang = danhSachHang[i];
//     const [stt, hoTen, ngaySinh, doiTuong, giaTriGoi] = sheetNguon.getRange(`A${hang}:E${hang}`).getValues()[0];
    
//     if (!hoTen) continue;

//     sheetDich.getRange("C11").setValue(stt);
//     sheetDich.getRange("C9").setValue(hoTen);
//     sheetDich.getRange("C10").setValue(ngaySinh);
//     sheetDich.getRange("C12").setValue(doiTuong);
//     sheetDich.getRange("D30").setValue(giaTriGoi);

//     SpreadsheetApp.flush();
//     Utilities.sleep(300); // Cho xử lý ổn định trước khi xuất PDF

//     const blob = xuatSheetThanhPDF(sheetDich);
//     blob.setName(`${stt} - ${hoTen}.pdf`);
//     folderGoc.createFile(blob);

//     Utilities.sleep(3000); // Nghỉ 3 giây để tránh lỗi "quá nhiều yêu cầu"
//   }
// }
//-------


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
    portrait: true,  // Chiều dọc A4
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
    fitw: true  // 👉 Fit to width (tự co nội dung cho vừa khổ giấy)
    // 👉 KHÔNG CẦN `range=...` nữa
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


