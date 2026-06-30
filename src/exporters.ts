import JSZip from "jszip";
import * as XLSX from "xlsx";
import { downloadBlob, photoCaseOf, safeFilePart } from "./logic";
import type { AppSettings, BackupPayload, SurveyItem, SurveyPhoto, SurveyStore } from "./types";

const stamp = () => new Date().toISOString().slice(0, 10).replaceAll("-", "");
const stampTime = () => new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");

function photoOf(photos: SurveyPhoto[], type: SurveyPhoto["type"], item?: SurveyItem, store?: SurveyStore) {
  return photos.find((photo) => photo.type === type && (item ? photo.itemId === item.id : true) && (store ? photo.storeId === store.id : true));
}

export async function exportRegionExcel(region: string, items: SurveyItem[]) {
  const group = [
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "실물 확인가능",
    "실물 확인가능",
    "실물 확인가능",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "정상진열 안될 시",
    "정상진열 안될 시",
    "정상진열 안될 시",
    "정상진열 안될 시",
    "특이사항",
    "조사일",
  ];
  const headers = [
    "순번",
    "업체명",
    "업체연락처",
    "마트명",
    "바코드",
    "물품명",
    "규격",
    "기준가격",
    "정상진열",
    "규격일치",
    "바코드일치",
    "정상가격",
    "할인가격",
    "시작",
    "종료",
    "기간구분",
    "바코드등록여부",
    "미판매",
    "미진열",
    "비정상",
    "특이사항",
    "조사일",
  ];
  const rows = items.map((item) => [
    item.itemNo,
    item.companyName,
    item.companyTel,
    item.martName,
    item.barcode,
    item.productName,
    item.spec,
    item.basePrice ?? "",
    item.normalDisplay,
    item.specMatch,
    item.barcodeMatch,
    item.normalPrice ?? "",
    item.discountPrice ?? "",
    item.discountStartDate,
    item.discountEndDate,
    `${item.discountType.replace("구두", "")}${item.discountOral || item.discountType.includes("구두") ? "구두" : ""}`,
    item.barcodeRegistered,
    item.abnormalStatus === "미판매" ? "O" : item.abnormalStatus ? "-" : "",
    item.abnormalStatus === "미진열" ? "O" : item.abnormalStatus ? "-" : "",
    item.abnormalDisplay === "O" ? "O" : "",
    item.memo,
    item.surveyDate,
  ]);
  const ws = XLSX.utils.aoa_to_sheet([group, headers, ...rows]);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
    { s: { r: 0, c: 8 }, e: { r: 0, c: 10 } },
    { s: { r: 0, c: 11 }, e: { r: 0, c: 15 } },
    { s: { r: 0, c: 16 }, e: { r: 0, c: 19 } },
  ];
  ws["!cols"] = headers.map((header) => ({ wch: Math.max(10, header.length + 4) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "조사결과");
  const buffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  await downloadBlob(new Blob([buffer], { type: "application/octet-stream" }), `price_survey_${safeFilePart(region)}_${stamp()}.xlsx`);
}

export async function exportRegionZip(region: string, stores: SurveyStore[], items: SurveyItem[], photos: SurveyPhoto[]) {
  const zip = new JSZip();
  items.forEach((item) => {
    const store = stores.find((candidate) => candidate.id === item.storeId);
    const front = store ? photoOf(photos, "STORE_FRONT", undefined, store) : undefined;
    const display = photoOf(photos, "PRODUCT_DISPLAY", item);
    const info = photoOf(photos, "PRODUCT_INFO_BARCODE", item);
    const pos = photoOf(photos, "POS_RECEIPT", item);
    const photoCase = photoCaseOf(item);
    if (front && photoCase !== "MISSING") zip.file(`${item.itemNo}.1.jpg`, front.blob);
    if (photoCase === "POS_ONLY") {
      if (pos) zip.file(`${item.itemNo}.2.jpg`, pos.blob);
    } else if (photoCase === "NORMAL") {
      if (display) zip.file(`${item.itemNo}.2.jpg`, display.blob);
      if (info) zip.file(`${item.itemNo}.3.jpg`, info.blob);
    }
  });
  await downloadBlob(await zip.generateAsync({ type: "blob", mimeType: "application/zip" }), `price_photos_${safeFilePart(region)}_${stamp()}.zip`);
}

const blobToDataUrl = (blob: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });

export async function exportBackup(region: string | undefined, stores: SurveyStore[], items: SurveyItem[], photos: SurveyPhoto[], settings: AppSettings) {
  const photoPayload = await Promise.all(
    photos.map(async ({ blob, ...photo }) => ({
      ...photo,
      dataUrl: await blobToDataUrl(blob),
    })),
  );
  const payload: BackupPayload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    scope: region ? "region" : "all",
    region,
    regions: Array.from(new Set(items.map((item) => item.region))).map((name) => ({ name, updatedAt: new Date().toISOString() })),
    stores,
    items,
    photos: photoPayload,
    settings,
  };
  const suffix = region ? safeFilePart(region) : "전체";
  await downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }), `price_backup_${suffix}_${stampTime()}.json`);
}

export async function dataUrlToBlob(dataUrl: string) {
  const response = await fetch(dataUrl);
  return await response.blob();
}
