/* global CustomFunctions */

const CHU_SO = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];

function docBaChuSo(baso, isFirst) {
  if (baso === 0) return "";
  const tram = Math.floor(baso / 100);
  const chuc = Math.floor((baso % 100) / 10);
  const donvi = baso % 10;
  const hasHundredText = tram > 0 || !isFirst;
  const parts = [];

  if (tram > 0) {
    parts.push(CHU_SO[tram] + " trăm");
  } else if (hasHundredText) {
    parts.push("không trăm");
  }

  if (chuc >= 2) {
    parts.push(CHU_SO[chuc] + " mươi");
  } else if (chuc === 1) {
    parts.push("mười");
  } else if (chuc === 0 && donvi > 0 && hasHundredText) {
    parts.push("linh");
  }

  if (donvi > 0) {
    if (donvi === 1 && chuc >= 2) {
      parts.push("mốt");
    } else if (donvi === 5 && chuc >= 1) {
      parts.push("lăm");
    } else {
      parts.push(CHU_SO[donvi]);
    }
  }

  return parts.join(" ");
}

function docSoTienBangChu(soTien) {
  const num = Number(soTien);
  if (isNaN(num)) return "Giá trị không hợp lệ";
  if (num === 0) return "Không đồng";

  const isNegative = num < 0;
  let absNum = Math.round(Math.abs(num));
  let s = String(absNum);

  const groups = [];
  while (s.length > 0) {
    if (s.length > 3) {
      groups.unshift(s.slice(-3));
      s = s.slice(0, -3);
    } else {
      groups.unshift(s);
      s = "";
    }
  }

  const unitNames = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ", "tỷ tỷ"];
  const n = groups.length;
  const resultParts = [];

  for (let i = 0; i < n; i++) {
    const groupVal = parseInt(groups[i], 10);
    if (groupVal === 0) continue;
    const isFirst = i === 0;
    const text = docBaChuSo(groupVal, isFirst);
    const unit = unitNames[n - 1 - i] || "";
    resultParts.push(unit ? text + " " + unit : text);
  }

  let result = resultParts.join(" ").replace(/\s+/g, " ").trim();
  let full = (isNegative ? "âm " : "") + result + " đồng";
  return full.charAt(0).toUpperCase() + full.slice(1);
}

/**
 * Đọc số tiền bằng chữ tiếng Việt (VNĐ).
 * @customfunction DOCSO
 * @param {number} soTien Số tiền cần đọc bằng chữ.
 * @returns {string} Số tiền viết bằng chữ tiếng Việt.
 */
function DOCSO(soTien) {
  return docSoTienBangChu(soTien);
}

CustomFunctions.associate("DOCSO", DOCSO);
