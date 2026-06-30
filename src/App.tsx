import { Camera, CheckCircle2, ChevronDown, ChevronUp, Download, Info as InfoIcon, MapPin, Menu, MoreVertical, Phone, SlidersHorizontal, Search, Upload, X } from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import type { DragEvent } from "react";
import { clearAllData, deletePhoto, getItems, getPhotosByRegion, getPhotosByStore, getRegions, getSettings, getStores, importRegionData, now, putItem, putPhoto, putStore, saveParsedData, saveSettings, today, uid } from "./db";
import { parseContactRows, parseSurveyWorkbook, mergeContacts, rebuildStoresAndRegions } from "./excel";
import { dataUrlToBlob, exportBackup, exportRegionExcel, exportRegionZip } from "./exporters";
import { mapSearchAddress, requiredPhotoLabels, summarize } from "./logic";
import type { AppSettings, BackupPayload, PhotoType, Region, RegionStats, SurveyItem, SurveyPhoto, SurveyStore } from "./types";

type View = "upload" | "regions" | "workspace" | "store" | "items" | "item" | "backup" | "validation";
type Filter = "전체" | "미완료" | "미조사" | "조사중" | "완료" | "사진누락";
type StoreSort = "방문순서" | "주소순" | "품목 많은 순" | "미완료 많은 순" | "사진누락 많은 순";

const mapLinks = (address: string) => [
  ["구글", `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(mapSearchAddress(address))}`],
  ["네이버", `https://map.naver.com/p/search/${encodeURIComponent(mapSearchAddress(address))}`],
  ["카카오", `https://map.kakao.com/link/search/${encodeURIComponent(mapSearchAddress(address))}`],
];

const emptyStats: RegionStats = { total: 0, completed: 0, inProgress: 0, notStarted: 0, photoMissing: 0 };
const num = (value: string) => {
  const digits = value.replace(/\D/g, "");
  return digits === "" ? null : Number(digits);
};
const EXCEL_ACCEPT = ".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv,application/octet-stream";

function App() {
  const [view, setView] = useState<View>("upload");
  const [settings, setSettingsState] = useState<AppSettings>({ defaultSurveyDate: today() });
  const [regions, setRegions] = useState<Region[]>([]);
  const [stores, setStores] = useState<SurveyStore[]>([]);
  const [items, setItems] = useState<SurveyItem[]>([]);
  const [photos, setPhotos] = useState<SurveyPhoto[]>([]);
  const [selectedStoreId, setSelectedStoreId] = useState("");
  const [selectedItemId, setSelectedItemId] = useState("");
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<Filter>("미완료");
  const [surveyFile, setSurveyFile] = useState<File | null>(null);
  const [contactFile, setContactFile] = useState<File | null>(null);
  const [analysis, setAnalysis] = useState("");
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [storeSort, setStoreSort] = useState<StoreSort>("주소순");
  const [orderEditing, setOrderEditing] = useState(false);
  const [dragStoreId, setDragStoreId] = useState("");
  const [workspaceToolsOpen, setWorkspaceToolsOpen] = useState(false);
  const [itemToolsOpen, setItemToolsOpen] = useState(false);
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [contactStoreId, setContactStoreId] = useState("");

  const currentRegion = settings.currentRegion;
  const regionItems = useMemo(() => items.filter((item) => item.region === currentRegion), [items, currentRegion]);
  const regionStores = useMemo(() => stores.filter((store) => store.region === currentRegion), [stores, currentRegion]);
  const storeVisitCompare = (a: SurveyStore, b: SurveyStore) =>
    (a.visitOrder ?? 999999) - (b.visitOrder ?? 999999) || `${a.storeAddress} ${a.storeName}`.localeCompare(`${b.storeAddress} ${b.storeName}`, "ko");
  const sortedRegionStores = useMemo(() => {
    const statsOf = (storeId: string) => summarize(regionItems.filter((item) => item.storeId === storeId), photos.filter((photo) => photo.storeId === storeId));
    return [...regionStores].sort((a, b) => {
      const as = statsOf(a.id);
      const bs = statsOf(b.id);
      if (storeSort === "방문순서") return storeVisitCompare(a, b);
      if (storeSort === "품목 많은 순") return bs.total - as.total;
      if (storeSort === "사진누락 많은 순") return bs.photoMissing - as.photoMissing;
      if (storeSort === "주소순") return `${a.storeAddress} ${a.storeName}`.localeCompare(`${b.storeAddress} ${b.storeName}`, "ko");
      return (bs.notStarted + bs.inProgress) - (as.notStarted + as.inProgress);
    });
  }, [regionStores, regionItems, photos, storeSort]);
  const visibleRegionStores = useMemo(() => sortedRegionStores.filter((store) => {
    if (!`${store.storeName} ${store.storeAddress}`.includes(query)) return false;
    const ownItems = regionItems.filter((item) => item.storeId === store.id);
    const ownPhotos = photos.filter((photo) => photo.storeId === store.id);
    const ownStats = summarize(ownItems, ownPhotos);
    if (filter === "미완료" && ownStats.completed >= ownStats.total) return false;
    if (filter !== "전체" && filter !== "미완료" && filter !== "사진누락" && !ownItems.some((item) => item.status === filter)) return false;
    if (filter === "사진누락" && ownStats.photoMissing === 0) return false;
    return true;
  }), [sortedRegionStores, query, regionItems, photos, filter]);
  const selectedStore = stores.find((store) => store.id === selectedStoreId);
  const storeItems = useMemo(() => items.filter((item) => item.storeId === selectedStoreId), [items, selectedStoreId]);
  const selectedItem = items.find((item) => item.id === selectedItemId);
  const stats = useMemo(() => summarize(regionItems, photos), [regionItems, photos]);

  async function refresh(region = currentRegion) {
    const [nextSettings, nextRegions, allStores, allItems] = await Promise.all([getSettings(), getRegions(), getStores(), getItems()]);
    const photoRegion = region ?? nextSettings.currentRegion;
    const nextPhotos = photoRegion ? await getPhotosByRegion(photoRegion) : [];
    setSettingsState(nextSettings);
    setRegions(nextRegions);
    setStores(allStores);
    setItems(allItems);
    setPhotos(nextPhotos);
    if (nextRegions.length && view === "upload") setView("regions");
  }

  useEffect(() => {
    refresh();
  }, []);

  useEffect(() => {
    window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }, [view, currentRegion, selectedStoreId, selectedItemId]);

  useEffect(() => {
    if (view !== "workspace") {
      setOrderEditing(false);
      setDragStoreId("");
    }
  }, [view]);

  async function updateSettings(patch: Partial<AppSettings>) {
    const next = { ...settings, ...patch };
    setSettingsState(next);
    await saveSettings(next);
  }

  async function analyzeFiles() {
    if (!surveyFile) {
      alert("조사표 엑셀을 먼저 선택하세요.");
      return;
    }
    if ((items.length || stores.length || photos.length) && !confirm("새 자료 분석을 시작하면 기존 입력 데이터와 사진을 초기화합니다. 계속할까요?")) {
      return;
    }
    if (items.length || stores.length || photos.length) {
      await clearAllData();
    }
    setIsAnalyzing(true);
    setAnalysis("자료 분석 중입니다. 모바일에서는 10~30초 정도 걸릴 수 있습니다.");
    await new Promise((resolve) => setTimeout(resolve, 50));
    try {
      const parsed = await parseSurveyWorkbook(surveyFile);
      let parsedItems = parsed.items;
      let matched = 0;
      if (contactFile) {
        const before = parsedItems.filter((item) => item.companyTel).length;
        parsedItems = mergeContacts(parsedItems, await parseContactRows(contactFile));
        matched = parsedItems.filter((item) => item.companyTel).length - before;
      }
      const rebuilt = rebuildStoresAndRegions(parsedItems);
      parsedItems = rebuilt.items;
      const parsedStores = rebuilt.stores.map((store) => {
        const first = parsedItems.find((item) => item.storeId === store.id);
        return first ? { ...store, storeAddress: first.storeAddress || store.storeAddress, storeName: first.storeName || store.storeName } : store;
      });
      await saveParsedData(rebuilt.regions, parsedStores, parsedItems);
      setAnalysis(`자료 분석 완료: 전체 품목 ${parsedItems.length.toLocaleString()}개 / 지역 ${rebuilt.regions.length}개 / 방문지 ${parsedStores.length}개 / 업체 연락처 매칭 ${Math.max(0, matched)}개`);
      await refresh(rebuilt.regions[0]?.name);
      setView("regions");
    } catch (error) {
      console.error(error);
      setAnalysis("자료 분석 실패: 엑셀 파일을 확인해 주세요. 모바일에서는 고화질 바코드 파일 대신 기본 조사표를 먼저 사용해 보세요.");
      alert("자료 분석에 실패했습니다. 고화질 바코드 파일은 모바일에서 무거울 수 있어 기본 조사표로 먼저 시도해 주세요.");
    } finally {
      setIsAnalyzing(false);
    }
  }

  async function chooseRegion(region: string) {
    await updateSettings({ currentRegion: region });
    setSelectedStoreId("");
    setSelectedItemId("");
    setQuery("");
    setFilter("미완료");
    setOrderEditing(false);
    setDragStoreId("");
    await refresh(region);
    setView("workspace");
  }

  async function openStore(store: SurveyStore) {
    setSelectedStoreId(store.id);
    await updateSettings({ lastOpenedStoreId: store.id, currentRegion: store.region });
    setView("store");
  }

  async function saveVisitOrder(store: SurveyStore, value: string) {
    const order = Number(value.replace(/\D/g, ""));
    await putStore({ ...store, visitOrder: order || undefined, updatedAt: now() });
    await refresh(store.region);
  }

  async function moveVisitOrder(store: SurveyStore, direction: -1 | 1) {
    const ordered = [...regionStores].sort(storeVisitCompare);
    const index = ordered.findIndex((candidate) => candidate.id === store.id);
    const nextIndex = index + direction;
    if (index < 0 || nextIndex < 0 || nextIndex >= ordered.length) return;
    [ordered[index], ordered[nextIndex]] = [ordered[nextIndex], ordered[index]];
    await Promise.all(ordered.map((candidate, orderIndex) => putStore({ ...candidate, visitOrder: orderIndex + 1, updatedAt: now() })));
    await refresh(store.region);
  }

  async function reorderVisitOrder(draggedId: string, targetId: string) {
    if (!draggedId || draggedId === targetId) return;
    const ordered = [...regionStores].sort(storeVisitCompare);
    const from = ordered.findIndex((candidate) => candidate.id === draggedId);
    const to = ordered.findIndex((candidate) => candidate.id === targetId);
    if (from < 0 || to < 0) return;
    const [moved] = ordered.splice(from, 1);
    ordered.splice(to, 0, moved);
    await Promise.all(ordered.map((candidate, orderIndex) => putStore({ ...candidate, visitOrder: orderIndex + 1, updatedAt: now() })));
    setDragStoreId("");
    await refresh(currentRegion);
  }

  async function saveStorePhoto(file: File) {
    if (!selectedStore) return;
    if (selectedStore.frontPhotoId) await deletePhoto(selectedStore.frontPhotoId);
    const photo: SurveyPhoto = { id: uid("photo"), region: selectedStore.region, storeId: selectedStore.id, type: "STORE_FRONT", blob: file, originalName: file.name, mimeType: file.type, takenAt: now() };
    await putPhoto(photo);
    await putStore({ ...selectedStore, frontPhotoId: photo.id, status: "진행중", startedAt: selectedStore.startedAt ?? now(), updatedAt: now() });
    await refresh(selectedStore.region);
  }

  async function removeStorePhoto() {
    if (!selectedStore?.frontPhotoId) return;
    await deletePhoto(selectedStore.frontPhotoId);
    await putStore({ ...selectedStore, frontPhotoId: undefined, updatedAt: now() });
    await refresh(selectedStore.region);
  }

  async function saveItemPhoto(item: SurveyItem, type: PhotoType, file: File) {
    const oldPhotos = await getPhotosByStore(item.storeId);
    await Promise.all(oldPhotos.filter((photo) => photo.itemId === item.id && photo.type === type).map((photo) => deletePhoto(photo.id)));
    const photo: SurveyPhoto = { id: uid("photo"), region: item.region, storeId: item.storeId, itemId: item.id, type, blob: file, originalName: file.name, mimeType: file.type, takenAt: now() };
    await putPhoto(photo);
    await refresh(item.region);
  }

  async function removeItemPhoto(photo: SurveyPhoto) {
    await deletePhoto(photo.id);
    await refresh(photo.region);
  }

  async function saveItem(next: SurveyItem) {
    const storePhotos = await getPhotosByStore(next.storeId);
    const missing = requiredPhotoLabels(next, storePhotos);
    if (missing.length) {
      const label = next.normalDisplay === "X" ? "정상진열 X 품목" : next.normalDisplay === "O" ? "정상진열 품목" : "정상진열 여부가 선택되지 않아 기본 사진 기준";
      const ok = confirm(`사진이 부족합니다.\n\n${label}은 아래 사진이 필요합니다.\n- ${missing.join("\n- ")}\n\n그래도 저장하시겠습니까?`);
      if (!ok) return false;
    }
    const saved: SurveyItem = { ...next, status: "완료", updatedAt: now() };
    await putItem(saved);
    await refresh(saved.region);
    return true;
  }

  async function doExportExcel(region = currentRegion) {
    if (!region) return;
    await exportRegionExcel(region, items.filter((item) => item.region === region));
  }

  async function doExportZip(region = currentRegion) {
    if (!region) return;
    const regionStoresForExport = stores.filter((store) => store.region === region);
    const regionItemsForExport = items.filter((item) => item.region === region);
    const regionPhotos = region === currentRegion ? photos : await getPhotosByRegion(region);
    await exportRegionZip(region, regionStoresForExport, regionItemsForExport, regionPhotos);
  }

  async function doBackup(region = currentRegion) {
    const sourceStores = region ? stores.filter((store) => store.region === region) : stores;
    const sourceItems = region ? items.filter((item) => item.region === region) : items;
    const sourcePhotos = region ? (region === currentRegion ? photos : await getPhotosByRegion(region)) : photos;
    await exportBackup(region, sourceStores, sourceItems, sourcePhotos, settings);
  }

  async function restoreBackup(file: File) {
    const payload = JSON.parse(await file.text()) as BackupPayload;
    const region = payload.region ?? payload.regions[0]?.name;
    if (!region) return;
    if (!confirm(`${region} 지역 데이터를 백업 파일 내용으로 덮어씁니다. 계속할까요?`)) return;
    const restoredPhotos = await Promise.all(payload.photos.map(async ({ dataUrl, ...photo }) => ({ ...photo, blob: await dataUrlToBlob(dataUrl) })));
    await importRegionData(region, payload.stores, payload.items, restoredPhotos);
    await updateSettings({ currentRegion: region });
    await refresh(region);
    setView("regions");
  }

  const regionSummary = (region: string) => {
    const regionStoresForSummary = stores.filter((store) => store.region === region);
    const regionItemsForSummary = items.filter((item) => item.region === region);
    const completed = regionStoresForSummary.filter((store) => {
      const own = regionItemsForSummary.filter((item) => item.storeId === store.id);
      return own.length > 0 && own.every((item) => item.status === "완료");
    }).length;
    const inProgress = regionStoresForSummary.filter((store) => {
      const own = regionItemsForSummary.filter((item) => item.storeId === store.id);
      return own.some((item) => item.status === "완료" || item.status === "조사중") && !(own.length > 0 && own.every((item) => item.status === "완료"));
    }).length;
    return {
      total: regionStoresForSummary.length,
      completed,
      inProgress,
      notStarted: Math.max(0, regionStoresForSummary.length - completed - inProgress),
      photoMissing: summarize(regionItemsForSummary, region === currentRegion ? photos : []).photoMissing,
    };
  };
  const canGoBack = view !== "upload" && !(view === "regions" && regions.length > 0);
  const goBack = () => {
    setMenuOpen(false);
    if (view === "workspace") setView("regions");
    else if (view === "store") setView("workspace");
    else if (view === "items") setView("store");
    else if (view === "item") setView("items");
    else if (view === "validation") setView(currentRegion ? "workspace" : "regions");
    else if (view === "backup") setView(regions.length ? "regions" : "upload");
    else if (view === "regions") setView("upload");
  };
  const screenTitle =
    view === "regions" ? "지역리스트"
    : view === "workspace" ? "업체리스트"
    : view === "store" ? "업체정보"
    : view === "items" ? "물품리스트"
    : view === "item" ? "가격정보"
    : view === "validation" ? "검증"
    : view === "backup" ? "백업/복원"
    : "국군복지단 가격조사";

  return (
    <div className="app">
      <header className={`topbar ${menuOpen ? "menu-open" : ""}`}>
        <div className="top-main">
          <button className="top-back icon-button" onClick={goBack} disabled={!canGoBack} aria-label="뒤로가기">←</button>
          <button className="brand" onClick={() => setView(regions.length ? "regions" : "upload")}>{screenTitle}</button>
          <span className="current">{currentRegion ? `현재 지역: ${currentRegion}` : "지역 미선택"}</span>
          <button className="top-toggle icon-button" onClick={() => setMenuOpen((value) => !value)} aria-expanded={menuOpen} aria-label="메뉴 열기">
            <Menu size={20} />
          </button>
        </div>
        <div className="top-actions">
          <button onClick={() => { setView("regions"); setMenuOpen(false); }}>지역 선택</button>
          <button disabled={!currentRegion} onClick={() => { setView("validation"); setMenuOpen(false); }}>검증</button>
          <button onClick={() => { setView("backup"); setMenuOpen(false); }}>백업/복원</button>
        </div>
      </header>

      {view === "upload" && (
        <main className="page narrow">
          <h1>국군복지단 가격조사 PWA</h1>
          <section className="panel">
            <label>조사표 엑셀 업로드<input type="file" accept={EXCEL_ACCEPT} onChange={(event) => setSurveyFile(event.target.files?.[0] ?? null)} /></label>
            <label>업체 연락처 엑셀 업로드<input type="file" accept={EXCEL_ACCEPT} onChange={(event) => setContactFile(event.target.files?.[0] ?? null)} /></label>
            <label>바코드 포함 조사표 업로드<input type="file" accept={EXCEL_ACCEPT} /></label>
            <p className="hint">모바일에서 파일이 안 보이면 파일 앱의 다운로드/OneDrive/Google 드라이브에서 엑셀 파일을 먼저 내려받은 뒤 선택하세요.</p>
            <button className="primary" onClick={analyzeFiles} disabled={isAnalyzing}><Upload size={18} />{isAnalyzing ? "자료 분석 중..." : "자료 분석 시작"}</button>
            {analysis && <p className="notice">{analysis}</p>}
          </section>
          <button disabled={!regions.length} onClick={() => setView("regions")}>지역 선택으로 이동</button>
        </main>
      )}

      {view === "regions" && (
        <main className="page">
          <SearchBox value={query} onChange={setQuery} placeholder="지역명 검색" />
          {currentRegion && regions.some((region) => region.name === currentRegion) && (
            <div className="recent-region">
              <span>최근 지역</span>
              <button onClick={() => chooseRegion(currentRegion)}>{currentRegion}</button>
            </div>
          )}
          <div className="grid">
            {regions.filter((region) => region.name.includes(query)).map((region) => {
              const summary = regionSummary(region.name);
              return (
                <article className="card" key={region.name}>
                  <h2>{region.name}</h2>
                  <p className="area-summary">{region.areaSummary || region.city || "-"}</p>
                  <p className="muted">담당부서: {region.department || "-"}</p>
                  <RegionSummary stats={summary.total ? summary : emptyStats} itemStats={summarize(items.filter((item) => item.region === region.name), region.name === currentRegion ? photos : [])} />
                  <div className="region-actions">
                    <button className="primary" onClick={() => chooseRegion(region.name)}>작업</button>
                    <button title="엑셀 내보내기" onClick={() => doExportExcel(region.name)}><Download size={16} />엑셀</button>
                    <button title="사진 ZIP" onClick={() => doExportZip(region.name)}><Download size={16} />사진</button>
                    <button title="백업 내려받기" onClick={() => doBackup(region.name)}><Download size={16} />백업</button>
                  </div>
                </article>
              );
            })}
          </div>
        </main>
      )}

      {view === "workspace" && currentRegion && (
        <main className="page">
          <div className="sticky-search workspace-search">
            <SearchBox value={query} onChange={setQuery} placeholder="마트명 / 업체명 / 주소 / 품목명 / 바코드" />
            <button className="tool-toggle" onClick={() => setWorkspaceToolsOpen((value) => !value)} aria-expanded={workspaceToolsOpen}>
              <SlidersHorizontal size={18} /> 필터 {workspaceToolsOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
            </button>
          </div>
          {workspaceToolsOpen && (
            <section className="tool-panel">
              <FilterBar filter={filter} setFilter={setFilter} />
              <label className="sort-control">정렬
                <select value={storeSort} onChange={(event) => setStoreSort(event.target.value as StoreSort)}>
                  <option>방문순서</option>
                  <option>주소순</option>
                  <option>미완료 많은 순</option>
                  <option>품목 많은 순</option>
                  <option>사진누락 많은 순</option>
                </select>
              </label>
              <button
                className={`order-edit-toggle ${orderEditing ? "active" : ""}`}
                onClick={() => {
                  setOrderEditing((value) => !value);
                  setStoreSort("방문순서");
                }}
              >
                순서 편집 {orderEditing ? "끄기" : "켜기"}
              </button>
              <a className="map-ref" target="_blank" href="https://www.google.com/maps/d/u/1/viewer?mid=1ej99Lo6WS4GROBCQPr0a66MhQR_vXuM&ll=37.49945198941339%2C127.04262669775987&z=14">조사대상 참고 지도 열기</a>
            </section>
          )}
          <div className="list">
            {visibleRegionStores.map((store) => {
              const ownItems = regionItems.filter((item) => item.storeId === store.id);
              const ownPhotos = photos.filter((photo) => photo.storeId === store.id);
              const ownStats = summarize(ownItems, ownPhotos);
              return (
                <StoreCard
                  key={store.id}
                  store={store}
                  stats={ownStats}
                  items={ownItems}
                  focused={selectedStoreId === store.id}
                  orderEditing={orderEditing}
                  dragging={dragStoreId === store.id}
                  onOpen={() => openStore(store)}
                  onContacts={() => setContactStoreId(store.id)}
                  onOrderChange={(value) => saveVisitOrder(store, value)}
                  onMoveUp={() => moveVisitOrder(store, -1)}
                  onMoveDown={() => moveVisitOrder(store, 1)}
                  onDragStart={() => setDragStoreId(store.id)}
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={() => reorderVisitOrder(dragStoreId, store.id)}
                  onDragEnd={() => setDragStoreId("")}
                />
              );
            })}
          </div>
        </main>
      )}

      {view === "store" && selectedStore && (
        <main className="page narrow">
          <section className="panel">
            <h1>{selectedStore.storeName}</h1>
            <p>{selectedStore.storeAddress}</p>
            <h2>업체 전경사진</h2>
            <PhotoInput label={selectedStore.frontPhotoId ? "업체사진" : "업체사진"} onFile={saveStorePhoto} />
            {selectedStore.frontPhotoId && <button className="danger full-button store-photo-delete" onClick={removeStorePhoto}>지우기</button>}
            <p className={selectedStore.frontPhotoId ? "ok" : "warn"}>촬영 상태: {selectedStore.frontPhotoId ? "촬영완료" : "미촬영"}</p>
          </section>
          <Contacts items={storeItems} />
          <section className="panel">
            <p>조사 품목: {storeItems.length.toLocaleString()}건</p>
            <label>방문 조사일<input type="date" value={selectedStore.surveyDate} onChange={async (event) => { await putStore({ ...selectedStore, surveyDate: event.target.value, updatedAt: now() }); await refresh(selectedStore.region); }} /></label>
            <button className="primary sticky-lite" onClick={() => selectedStore.frontPhotoId ? setView("items") : alert("업체 전경사진을 먼저 촬영/선택해 주세요.")}>조사 입력</button>
          </section>
        </main>
      )}

      {view === "items" && selectedStore && (
        <main className="page">
          <div className="sticky-search item-search">
            <SearchBox value={query} onChange={setQuery} placeholder="품목명 / 바코드 / 순번 검색" />
            <button className="tool-toggle icon-button" onClick={() => setSummaryOpen(true)} aria-label="현황 보기"><InfoIcon size={18} /></button>
            <button className="tool-toggle" onClick={() => setItemToolsOpen((value) => !value)} aria-expanded={itemToolsOpen}>
              <SlidersHorizontal size={18} /> 필터 {itemToolsOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
            </button>
          </div>
          {itemToolsOpen && (
            <section className="tool-panel">
              <Stats stats={summarize(storeItems, photos.filter((photo) => photo.storeId === selectedStore.id))} totalLabel="품목 전체" />
              <FilterBar filter={filter} setFilter={setFilter} />
            </section>
          )}
          <div className="list">
            {storeItems.filter((item) => `${item.itemNo} ${item.productName} ${item.barcode}`.includes(query)).filter((item) => filter === "전체" || (filter === "미완료" ? item.status !== "완료" : filter === "사진누락" ? requiredPhotoLabels(item, photos.filter((photo) => photo.storeId === item.storeId)).length > 0 : item.status === filter)).map((item) => (
              <article className={`card compact item-card ${selectedItemId === item.id ? "focused" : ""}`} key={item.id}>
                <div className="item-card-head"><h2 className="item-title"><span className="item-code">{item.itemNo}</span><span>{item.productName}</span></h2><Badge text={item.status} /></div>
                <p>바코드: {item.barcode || "-"} · 기준가격: {item.basePrice?.toLocaleString() ?? "-"}원</p>
                <p>조사가격: {item.normalPrice?.toLocaleString() ?? "-"}원</p>
                <button className="primary" onClick={() => { setSelectedItemId(item.id); setView("item"); }}>입력</button>
              </article>
            ))}
          </div>
        </main>
      )}

      {view === "item" && selectedItem && (
        <ItemEditor item={selectedItem} storeItems={storeItems} photos={photos.filter((photo) => photo.storeId === selectedItem.storeId)} onPhoto={saveItemPhoto} onDeletePhoto={removeItemPhoto} onSave={saveItem} onList={(focusId) => { if (focusId) setSelectedItemId(focusId); setView("items"); }} onMove={(id) => setSelectedItemId(id)} />
      )}

      {view === "validation" && (
        <main className="page">
          <Validation title="미완료 품목" items={regionItems.filter((item) => item.status !== "완료")} open={(id) => { setSelectedItemId(id); setView("item"); }} />
          <Validation title="사진누락 품목" items={regionItems.filter((item) => requiredPhotoLabels(item, photos.filter((photo) => photo.storeId === item.storeId)).length > 0)} open={(id) => { setSelectedItemId(id); setView("item"); }} />
          <Validation title="정상진열 X 품목" items={regionItems.filter((item) => item.normalDisplay === "X")} open={(id) => { setSelectedItemId(id); setView("item"); }} />
          <Validation title="특이사항 입력 품목" items={regionItems.filter((item) => item.memo)} open={(id) => { setSelectedItemId(id); setView("item"); }} />
        </main>
      )}

      {view === "backup" && (
        <main className="page narrow">
          <section className="backup-grid">
            <article className="panel">
              <h2>백업 내려받기</h2>
              <p className="muted">현재 브라우저에 저장된 조사 데이터와 사진을 JSON으로 저장합니다.</p>
              <button className="primary full-button" onClick={() => doBackup(undefined)}><Download size={17} />전체 백업 다운로드</button>
            </article>
            <article className="panel">
              <h2>백업 업로드</h2>
              <p className="muted">다른 폰이나 PC에서 만든 백업 JSON을 불러옵니다.</p>
              <label className="photo-button"><Upload size={18} />백업 JSON 업로드<input type="file" accept="application/json,.json" onChange={(event) => event.target.files?.[0] && restoreBackup(event.target.files[0])} /></label>
            </article>
            <article className="panel">
              <h2>초기화</h2>
              <p className="muted">기기 안의 IndexedDB 조사 데이터를 모두 삭제합니다.</p>
            <button className="danger" onClick={async () => { if (confirm("모든 IndexedDB 데이터를 삭제합니다. 계속할까요?")) { await clearAllData(); await refresh(undefined); setView("upload"); } }}>전체 데이터 초기화</button>
            </article>
          </section>
        </main>
      )}
      {contactStoreId && (
        <ContactModal
          store={stores.find((store) => store.id === contactStoreId)}
          items={items.filter((item) => item.storeId === contactStoreId)}
          onClose={() => setContactStoreId("")}
        />
      )}
      {summaryOpen && view === "workspace" && (
        <SummaryModal
          region={currentRegion}
          stats={stats}
          storeCount={regionStores.length}
          completedStoreCount={regionStores.filter((store) => regionItems.filter((item) => item.storeId === store.id).every((item) => item.status === "완료")).length}
          onClose={() => setSummaryOpen(false)}
        />
      )}
      {summaryOpen && view === "items" && selectedStore && (
        <SummaryModal
          region={selectedStore.storeName}
          stats={summarize(storeItems, photos.filter((photo) => photo.storeId === selectedStore.id))}
          storeCount={1}
          completedStoreCount={storeItems.every((item) => item.status === "완료") ? 1 : 0}
          onClose={() => setSummaryOpen(false)}
        />
      )}
    </div>
  );
}

function SearchBox({ value, onChange, placeholder }: { value: string; onChange: (value: string) => void; placeholder: string }) {
  return <label className="search"><Search size={18} /><input value={value} onChange={(event) => onChange(event.target.value)} placeholder={placeholder} /></label>;
}

function FilterBar({ filter, setFilter }: { filter: Filter; setFilter: (filter: Filter) => void }) {
  return <div className="segmented filter-chips">{(["전체", "미완료", "미조사", "조사중", "완료", "사진누락"] as Filter[]).map((value) => <button className={filter === value ? "active" : ""} key={value} onClick={() => setFilter(value)}>{value}</button>)}</div>;
}

function Stats({ stats, totalLabel = "전체" }: { stats: RegionStats; totalLabel?: string }) {
  return (
    <div className="stats">
      <div className="stats-line">
        <strong>{totalLabel} {stats.total.toLocaleString()}개</strong>
        <span>완료 {stats.completed.toLocaleString()}</span>
        <span>조사중 {stats.inProgress.toLocaleString()}</span>
        <span>미조사 {stats.notStarted.toLocaleString()}</span>
      </div>
      <div className="stats-progress"><span style={{ width: `${stats.total ? Math.round((stats.completed / stats.total) * 100) : 0}%` }} /></div>
      <div className={stats.photoMissing > 0 ? "stats-missing" : "stats-photo-ok"}>{stats.photoMissing > 0 ? `사진누락 ${stats.photoMissing.toLocaleString()}개` : "사진"}</div>
    </div>
  );
}

function RegionSummary({ stats, itemStats }: { stats: RegionStats; itemStats: RegionStats }) {
  const storePercent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;
  const itemPercent = itemStats.total ? Math.round((itemStats.completed / itemStats.total) * 100) : 0;
  return (
    <div className="region-summary">
      <div className="region-metric">
        <span>업체</span>
        <strong>{stats.completed.toLocaleString()}<small>/{stats.total.toLocaleString()}</small></strong>
        <div className="mini-progress"><i style={{ width: `${storePercent}%` }} /></div>
        <em>미조사 {stats.notStarted.toLocaleString()}</em>
      </div>
      <div className="region-metric">
        <span>품목</span>
        <strong>{itemStats.completed.toLocaleString()}<small>/{itemStats.total.toLocaleString()}</small></strong>
        <div className="mini-progress"><i style={{ width: `${itemPercent}%` }} /></div>
        <em>미조사 {itemStats.notStarted.toLocaleString()}</em>
      </div>
    </div>
  );
}

function Badge({ text }: { text: string }) {
  return <span className={`badge badge-${text}`}>{text}</span>;
}

function StoreCard({
  store,
  stats,
  items,
  focused,
  orderEditing,
  dragging,
  onOpen,
  onContacts,
  onOrderChange,
  onMoveUp,
  onMoveDown,
  onDragStart,
  onDragOver,
  onDrop,
  onDragEnd,
}: {
  store: SurveyStore;
  stats: RegionStats;
  items: SurveyItem[];
  focused: boolean;
  orderEditing: boolean;
  dragging: boolean;
  onOpen: () => void;
  onContacts: () => void;
  onOrderChange: (value: string) => void;
  onMoveUp: () => void;
  onMoveDown: () => void;
  onDragStart: () => void;
  onDragOver: (event: DragEvent<HTMLElement>) => void;
  onDrop: () => void;
  onDragEnd: () => void;
}) {
  const completed = items.filter((item) => item.status === "완료");
  const latestSurveyDate = completed.map((item) => item.surveyDate).filter(Boolean).sort().at(-1) ?? "-";
  const percent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;
  if (orderEditing) {
    return (
      <article className={`visit-order-row ${focused ? "focused" : ""} ${dragging ? "dragging" : ""}`} draggable onDragStart={onDragStart} onDragOver={onDragOver} onDrop={onDrop} onDragEnd={onDragEnd}>
        <span className="drag-handle" aria-hidden="true">☰</span>
        <input aria-label={`${store.storeName} 방문순서`} inputMode="numeric" value={store.visitOrder ?? ""} placeholder="-" onChange={(event) => onOrderChange(event.target.value)} />
        <div className="visit-order-name">
          <strong>{store.storeName}</strong>
          <span>{store.storeAddress || "주소 없음"}</span>
        </div>
        <div className="visit-order-actions">
          <button type="button" onClick={onMoveUp}>↑</button>
          <button type="button" onClick={onMoveDown}>↓</button>
        </div>
      </article>
    );
  }

  return (
    <article className={`card store-card ${focused ? "focused" : ""}`}>
      <div className="card-head">
        <div>
          <h2>{store.storeName}</h2>
          <p>{store.storeAddress || "주소 없음"}</p>
        </div>
        <details className="card-menu">
          <summary aria-label="업체 메뉴"><MoreVertical size={18} /></summary>
          <div className="menu-popover">
            {mapLinks(store.storeAddress).map(([name, href]) => <a key={name} href={href} target="_blank"><MapPin size={15} />{name} 지도</a>)}
          </div>
        </details>
      </div>
      <div className="store-progress">
        <div className="store-progress-head">
          <strong>업체: 완료 {stats.completed.toLocaleString()} / 전체 {stats.total.toLocaleString()}</strong>
          <span>미조사 {(stats.total - stats.completed).toLocaleString()}</span>
        </div>
        <div className="progress-line"><span style={{ width: `${percent}%` }} /></div>
      </div>
      <div className="store-meta">
        <span className={`store-photo-badge ${store.frontPhotoId ? "done" : ""}`}>업체사진</span>
        {stats.photoMissing > 0 && <span className="store-missing">품목사진 누락 {stats.photoMissing.toLocaleString()}건</span>}
        <span className="store-date">조사일: {latestSurveyDate}</span>
      </div>
      <div className="card-actions">
        <button onClick={onContacts}><Phone size={16} />담당자 정보</button>
        <button className="primary" onClick={onOpen}>입력</button>
      </div>
    </article>
  );
}

function ContactModal({ store, items, onClose }: { store?: SurveyStore; items: SurveyItem[]; onClose: () => void }) {
  if (!store) return null;
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal">
        <div className="modal-head">
          <div>
            <h2>{store.storeName}</h2>
            <p>{store.storeAddress || "주소 없음"}</p>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        <Contacts items={items} />
      </section>
    </div>
  );
}

function SummaryModal({ region, stats, storeCount, completedStoreCount, onClose }: { region?: string; stats: RegionStats; storeCount: number; completedStoreCount: number; onClose: () => void }) {
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal">
        <div className="modal-head">
          <div>
            <h2>{region ?? "현재 지역"} 현황</h2>
            <p>업체와 품목 기준 진행률입니다.</p>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        <Stats stats={stats} />
        <div className="summary-grid">
          <div><strong>{storeCount.toLocaleString()}</strong><span>방문지</span></div>
          <div><strong>{completedStoreCount.toLocaleString()}</strong><span>완료 업체</span></div>
          <div><strong>{stats.completed.toLocaleString()}</strong><span>완료 품목</span></div>
          <div><strong>{(stats.total - stats.completed).toLocaleString()}</strong><span>남은 품목</span></div>
        </div>
      </section>
    </div>
  );
}

function Contacts({ items }: { items: SurveyItem[] }) {
  const contacts = Array.from(new Map(items.map((item) => [`${item.companyName}|${item.companyManager}|${item.companyTel}`, item])).values());
  return (
    <section className="panel">
      <h2>업체 연락처</h2>
      {contacts.length === 0 && <p className="warn">확인 필요: 연락처 정보가 없습니다.</p>}
      {contacts.map((item) => {
        const count = items.filter((candidate) => candidate.companyName === item.companyName && candidate.companyManager === item.companyManager && candidate.companyTel === item.companyTel).length;
        return (
          <div className="contact" key={`${item.companyName}-${item.companyManager}-${item.companyTel}`}>
            <strong>{item.companyName || "업체명 확인 필요"}</strong>
            <span>담당자: {item.companyManager || "확인 필요"}</span>
            {item.companyTel ? <a href={`tel:${item.companyTel.replace(/[^\d+]/g, "")}`}><Phone size={15} />{item.companyTel}</a> : <span className="warn">전화: 확인 필요</span>}
            <span>해당 품목: {count.toLocaleString()}개</span>
          </div>
        );
      })}
    </section>
  );
}

function ItemContact({ item }: { item: SurveyItem }) {
  const hasAnyContact = Boolean(item.companyName || item.companyManager || item.companyTel);
  return (
    <section className={`item-contact ${hasAnyContact && item.companyTel ? "" : "needs-check"}`}>
      <div>
        <strong>{item.companyName || "업체명 확인 필요"}</strong>
        <span>담당자: {item.companyManager || "확인 필요"}</span>
      </div>
      {item.companyTel ? <a href={`tel:${item.companyTel.replace(/[^\d+]/g, "")}`}><Phone size={15} />{item.companyTel}</a> : <span className="warn">연락처 확인 필요</span>}
    </section>
  );
}

function PhotoInput({ id, label, onFile }: { id?: string; label: string; onFile: (file: File) => void | Promise<void> }) {
  const pickId = `${id ?? uid("photo_pick")}-pick`;
  const cameraId = `${id ?? uid("photo_camera")}-camera`;
  return (
    <div className="photo-picker">
      <span>{label}</span>
      <div>
        <label className="photo-button" htmlFor={cameraId}><Camera size={18} />촬영<input id={cameraId} type="file" accept="image/*" capture="environment" onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])} /></label>
        <label className="photo-button" htmlFor={pickId}><Upload size={18} />선택<input id={pickId} type="file" accept="image/*" onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])} /></label>
      </div>
    </div>
  );
}

function ItemEditor({ item, storeItems, photos, onPhoto, onDeletePhoto, onSave, onList, onMove }: { item: SurveyItem; storeItems: SurveyItem[]; photos: SurveyPhoto[]; onPhoto: (item: SurveyItem, type: PhotoType, file: File) => Promise<void>; onDeletePhoto: (photo: SurveyPhoto) => Promise<void>; onSave: (item: SurveyItem) => Promise<boolean>; onList: (focusId?: string) => void; onMove: (id: string) => void }) {
  const [draft, setDraft] = useState(item);
  const [photoMessage, setPhotoMessage] = useState("");
  const [saveMessage, setSaveMessage] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  useEffect(() => setDraft(item), [item.id]);
  useEffect(() => {
    if (!saveMessage) return;
    const timer = window.setTimeout(() => setSaveMessage(""), 5000);
    return () => window.clearTimeout(timer);
  }, [saveMessage]);
  const update = (patch: Partial<SurveyItem>) => setDraft((old) => ({ ...old, ...patch, status: old.status === "미조사" ? "조사중" : old.status }));
  const missing = requiredPhotoLabels(draft, photos);
  const itemPhotos = {
    display: photos.find((photo) => photo.itemId === draft.id && photo.type === "PRODUCT_DISPLAY"),
    info: photos.find((photo) => photo.itemId === draft.id && photo.type === "PRODUCT_INFO_BARCODE"),
    pos: photos.find((photo) => photo.itemId === draft.id && photo.type === "POS_RECEIPT"),
  };
  const upload = async (type: PhotoType, file: File, label: string) => {
    await onPhoto(draft, type, file);
    setPhotoMessage(`${label} 업로드 완료`);
  };
  const nextTodoId = () => {
    const currentIndex = storeItems.findIndex((candidate) => candidate.id === draft.id);
    const ordered = [...storeItems.slice(currentIndex + 1), ...storeItems.slice(0, Math.max(0, currentIndex))];
    return ordered.find((candidate) => candidate.status !== "완료")?.id;
  };
  const nextSequentialId = () => {
    const currentIndex = storeItems.findIndex((candidate) => candidate.id === draft.id);
    return storeItems[currentIndex + 1]?.id;
  };
  const handleSave = async () => {
    setIsSaving(true);
    setSaveMessage("저장 중...");
    try {
      const saved = await onSave(draft);
      if (saved) {
        setDraft((old) => ({ ...old, status: "완료" }));
        setSaveMessage(`저장 완료 · ${new Date().toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" })}`);
        const nextId = nextTodoId();
        if (nextId) {
          if (confirm("저장되었습니다. 다음 미등록 상품으로 이동할까요?")) onMove(nextId);
        } else if (confirm("전 품목 입력완료입니다. 목록으로 돌아갈까요?")) {
          onList();
        }
      } else {
        setSaveMessage("저장이 취소되었습니다.");
      }
    } catch (error) {
      console.error(error);
      setSaveMessage("저장 실패: 다시 눌러주세요.");
      alert("저장에 실패했습니다. 다시 시도해 주세요.");
    } finally {
      setIsSaving(false);
    }
  };
  const saveAndList = async () => {
    setIsSaving(true);
    setSaveMessage("저장 중...");
    try {
      const saved = await onSave(draft);
      if (saved) onList(nextTodoId());
      else setSaveMessage("저장이 취소되었습니다.");
    } catch (error) {
      console.error(error);
      setSaveMessage("저장 실패: 다시 눌러주세요.");
      alert("저장에 실패했습니다. 다시 시도해 주세요.");
    } finally {
      setIsSaving(false);
    }
  };
  const saveAndNext = async () => {
    const nextId = nextSequentialId();
    if (nextId) onMove(nextId);
    else alert("마지막 품목입니다.");
  };
  return <main className="page item-page"><section className="item-hero compact-hero"><div><h1 className="item-title"><span className="item-code">{draft.itemNo}</span><span>{draft.productName}</span></h1><Badge text={draft.status} /></div></section>
    <ItemContact item={draft} />
    <details className="panel" open><summary>① 업체 제시정보</summary><Info item={draft} /></details>
    <section className="panel"><h2>② 실물 확인</h2><Choice label="정상진열" value={draft.normalDisplay} values={["O", "X"]} onChange={(value) => update({ normalDisplay: value as SurveyItem["normalDisplay"], photoCase: value === "X" ? "POS_ONLY" : value === "O" ? "NORMAL" : "", specMatch: value === "X" ? "" : draft.specMatch, barcodeMatch: value === "X" ? "" : draft.barcodeMatch, barcodeRegistered: value === "O" ? "" : draft.barcodeRegistered, abnormalStatus: value === "O" ? "" : draft.abnormalStatus, posChecked: value === "O" ? "" : draft.posChecked, posPrice: value === "O" ? null : draft.posPrice })} /><Choice label="규격일치" disabled={draft.normalDisplay !== "O"} value={draft.normalDisplay === "O" ? draft.specMatch : ""} values={["O", "X", "-"]} onChange={(value) => update({ specMatch: value as SurveyItem["specMatch"] })} /><Choice label="바코드일치" disabled={draft.normalDisplay !== "O"} value={draft.normalDisplay === "O" ? draft.barcodeMatch : ""} values={["O", "X", "-"]} onChange={(value) => update({ barcodeMatch: value as SurveyItem["barcodeMatch"] })} /></section>
    <section className={`panel ${draft.normalDisplay === "X" ? "" : "disabled-block"}`}><h2>③ 진열상태</h2><Choice label="바코드 등록 여부" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.barcodeRegistered : ""} values={["O", "X"]} onChange={(value) => update({ barcodeRegistered: value as SurveyItem["barcodeRegistered"] })} /><Choice label="상태" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.abnormalStatus : ""} values={["미진열", "미판매"]} onChange={(value) => update({ abnormalStatus: value as SurveyItem["abnormalStatus"] })} /><Choice label="POS 조회 여부" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.posChecked : ""} values={["조회함", "조회불가", "미조회"]} onChange={(value) => update({ posChecked: value as SurveyItem["posChecked"] })} /><Money label="POS 확인 가격" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.posPrice : null} onChange={(value) => update({ posPrice: num(value) })} /></section>
    <section className="panel"><h2>④ 사진자료</h2><p className="photo-rule">업체사진: {photos.some((photo) => photo.type === "STORE_FRONT") ? "촬영완료" : "미촬영"}</p>{!draft.normalDisplay && <p className="notice">먼저 ② 실물 확인에서 정상진열 O/X를 선택하면 필요한 사진 입력칸이 표시됩니다.</p>}{photoMessage && <p className="ok upload-message">{photoMessage}</p>}{draft.normalDisplay === "O" && <><PhotoSlot id="photo-product-display" label="제품진열사진" photo={itemPhotos.display} onFile={(file) => upload("PRODUCT_DISPLAY", file, "제품진열사진")} onDelete={onDeletePhoto} /><PhotoSlot id="photo-product-info" label="제품정보/후면/바코드사진" photo={itemPhotos.info} onFile={(file) => upload("PRODUCT_INFO_BARCODE", file, "제품정보/후면/바코드사진")} onDelete={onDeletePhoto} /></>}{draft.normalDisplay === "X" && <PhotoSlot id="photo-pos-receipt" label="POS/영수증사진" photo={itemPhotos.pos} onFile={(file) => upload("POS_RECEIPT", file, "POS/영수증사진")} onDelete={onDeletePhoto} />}{missing.length > 0 && <p className="warn">사진누락: {missing.join(", ")}</p>}</section>
    <section className="panel"><h2>⑤ 가격</h2><Money label="정상가" value={draft.normalPrice} onChange={(value) => update({ normalPrice: num(value) })} /><Choice label="할인 여부" value={draft.hasDiscount === null ? "" : draft.hasDiscount ? "할인 있음" : "할인 없음"} values={["할인 없음", "할인 있음"]} onChange={(value) => update({ hasDiscount: value === "할인 있음" })} /><div className={draft.hasDiscount === false ? "disabled-block" : ""}><Money label="할인가" value={draft.discountPrice} disabled={draft.hasDiscount === false} onChange={(value) => update({ discountPrice: num(value) })} /><label>할인 시작일<input type="date" disabled={draft.hasDiscount === false} value={draft.discountStartDate} onChange={(event) => update({ discountStartDate: event.target.value })} /></label><label>할인 종료일<input type="date" disabled={draft.hasDiscount === false} value={draft.discountEndDate} onChange={(event) => update({ discountEndDate: event.target.value })} /></label><DiscountPeriod disabled={draft.hasDiscount === false} value={draft.discountType} oral={draft.discountOral ?? draft.discountType.includes("구두")} onChange={(discountType, discountOral) => update({ discountType, discountOral })} /></div>{draft.basePrice !== null && draft.normalPrice !== null && <p className="notice">참고: 정상가가 기준가격보다 {draft.normalPrice < draft.basePrice ? "낮습니다" : draft.normalPrice > draft.basePrice ? "높습니다" : "같습니다"}.</p>}</section>
    <section className="panel"><h2>⑥ 비정상 진열 / 비고</h2><div className="abnormal-block"><Choice label="비정상진열" value={draft.abnormalDisplay ?? ""} values={["O", "X"]} onChange={(value) => update({ abnormalDisplay: value as SurveyItem["abnormalDisplay"] })} />{draft.abnormalDisplay === "O" && <p className="small-help warn">비정상진열이면 어떤 위치에 어떻게 진열되어 있었는지 아래 비고에 적어주세요.</p>}</div><div className="memo-block"><h3>비고</h3><p className="small-help">자주 쓰는 문구를 누르면 비고에 추가됩니다.</p><div className="chips memo-chips">{["가격표 수기 작성", "POS 확인", "규격 불일치", "바코드 불일치", "장기 할인", "구두 확인", "비정상진열", "미진열", "미판매", "폐점", "품절", "기타"].map((text) => <button key={text} onClick={() => update({ memo: draft.memo ? `${draft.memo} / ${text}` : text })}>{text}</button>)}</div><textarea placeholder="예: 같은 카테고리 매대가 아닌 행사 매대에 단독 진열 / 사진 촬영 불가 / POS 확인" value={draft.memo} onChange={(event) => update({ memo: event.target.value })} /></div></section>
    {saveMessage && <div className={`save-toast ${saveMessage.includes("실패") ? "danger-toast" : ""}`}>{saveMessage}</div>}
    <div className="item-action-fab">
      <div className="item-progress-mini"><span style={{ width: `${storeItems.length ? Math.round((storeItems.filter((candidate) => candidate.status === "완료").length + (draft.status === "완료" && !storeItems.find((candidate) => candidate.id === draft.id && candidate.status === "완료") ? 1 : 0)) / storeItems.length * 100) : 0}%` }} /></div>
      <button type="button" onClick={saveAndList} disabled={isSaving}>목록</button>
      <button type="button" className="primary" onClick={handleSave} disabled={isSaving} aria-label="저장"><CheckCircle2 size={19} />{isSaving ? "저장 중" : "저장"}</button>
      <button type="button" onClick={saveAndNext} disabled={isSaving}>다음</button>
    </div>
  </main>;
}

function PhotoSlot({ id, label, photo, onFile, onDelete }: { id: string; label: string; photo?: SurveyPhoto; onFile: (file: File) => void | Promise<void>; onDelete: (photo: SurveyPhoto) => void | Promise<void> }) {
  return (
    <div id={`${id}-slot`} className={`photo-slot ${photo ? "uploaded" : ""}`}>
      <div>
        <strong>{label}</strong>
        <span>{photo ? "업로드됨" : "미업로드"}</span>
      </div>
      <div className="photo-actions">
        <PhotoInput id={id} label={photo ? "다시 올리기" : "촬영/첨부"} onFile={onFile} />
        {photo && <button className="danger" onClick={() => onDelete(photo)}>지우기</button>}
      </div>
    </div>
  );
}

function Info({ item }: { item: SurveyItem }) {
  return <dl className="info"><dt>업체명</dt><dd>{item.companyName}</dd><dt>업체연락처</dt><dd>{item.companyTel}</dd><dt>마트명</dt><dd>{item.martName}</dd><dt>바코드</dt><dd>{item.barcode}</dd><dt>물품명</dt><dd>{item.productName}</dd><dt>규격</dt><dd>{item.spec}</dd><dt>기준가격</dt><dd>{item.basePrice?.toLocaleString() ?? "-"}</dd></dl>;
}

function DiscountPeriod({ value, oral, disabled, onChange }: { value: string; oral: boolean; disabled?: boolean; onChange: (value: string, oral: boolean) => void }) {
  const normalized = value.replace("구두", "");
  return (
    <div className="field-row">
      <span>기간구분</span>
      <div className="period-control">
        <div className="segmented">
          <button disabled={disabled} className={normalized === "①" ? "active" : ""} onClick={() => onChange(normalized === "①" ? "" : "①", oral)}>① 31일 이내</button>
          <button disabled={disabled} className={normalized === "②" ? "active" : ""} onClick={() => onChange(normalized === "②" ? "" : "②", oral)}>② 32일 이상</button>
        </div>
        <label className="inline-check"><input type="checkbox" disabled={disabled} checked={oral} onChange={(event) => onChange(normalized, event.target.checked)} /> 구두 확인</label>
      </div>
    </div>
  );
}

function Choice({ label, value, values, disabled, onChange }: { label: string; value: string; values: string[]; disabled?: boolean; onChange: (value: string) => void }) {
  return <div className="field-row"><span>{label}</span><div className="segmented">{values.map((candidate) => <button disabled={disabled} className={value === candidate ? "active" : ""} key={candidate} onClick={() => onChange(value === candidate ? "" : candidate)}>{candidate}</button>)}</div></div>;
}

function Money({ label, value, disabled, onChange }: { label: string; value: number | null; disabled?: boolean; onChange: (value: string) => void }) {
  return <label>{label}<input inputMode="numeric" enterKeyHint="done" pattern="[0-9]*" disabled={disabled} value={typeof value === "number" && Number.isFinite(value) ? value : ""} onChange={(event) => onChange(event.target.value.replace(/\D/g, ""))} onKeyDown={(event) => { if (event.key === "Enter") event.currentTarget.blur(); }} placeholder="원" /></label>;
}

function Validation({ title, items, open }: { title: string; items: SurveyItem[]; open: (id: string) => void }) {
  return <section className="panel"><h2>{title} ({items.length.toLocaleString()}개)</h2>{items.slice(0, 80).map((item) => <button className="row-button" key={item.id} onClick={() => open(item.id)}>{item.itemNo} · {item.productName} · {item.storeName}</button>)}</section>;
}

export default App;
