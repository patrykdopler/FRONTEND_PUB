

  // extract systems from item.name
  function extractSystem(name){
    if(!name) return "";

    let s = String(name);

    // obetnij wszystko po pierwszym nawiasie z wymiarami, np. (B=..., H=...)
    const parenIdx = s.indexOf("(");
    if(parenIdx !== -1){
      s = s.substring(0, parenIdx);
    }

    // usuń 'Nazwa:' jeśli by się gdzieś wkradło
    s = s.replace(/Nazwa:/i, "").trim();

    // usuń wstępne 'Poz.'
    s = s.replace(/^Poz\./i, "").trim();

    // usuń pierwszy token typu OBZ/OZ/itp., ale NIE jeśli zaczyna się od MB lub AS
    s = s.replace(/^(?!MB\b)(?!AS\b)[A-Z0-9]+\s+/i, "");

    // usuń numer pozycji na początku
    s = s.replace(/^\d+\s+/, "");

    const parts = s.split(/\s+/);

    // pierwsze AS albo MB (MB-79N, MB60EI, AS 75P itd.)
    const idx = parts.findIndex(p => /^AS/i.test(p) || /^MB/i.test(p));
    if(idx === -1) return "";

    s = parts.slice(idx).join(" ");

    // jeżeli jest ' - ' to rozdziel na system + opis, w przeciwnym razie pierwszy token to system
    const dashIndex = s.indexOf(" - ");
    let left, right;
    if(dashIndex !== -1){
      left  = s.substring(0, dashIndex).trim();
      right = s.substring(dashIndex + 3).trim();
    } else {
      const toks = s.split(/\s+/);
      left  = toks[0];
      right = toks.slice(1).join(" ").trim();
    }

    // sprzątanie opisu
    left  = left.replace(/\(.*?\)/g, "").trim();
    right = right.replace(/\(.*?\)/g, "").trim();
    right = right
      .replace(/B=\d+/gi, "")
      .replace(/H=\d+/gi, "")
      .replace(/L=\d+/gi, "")
      .replace(/\s{2,}/g, " ")
      .replace(/,+$/, "")
      .trim();

    if(!left || !right) return "";

    return left + " – " + right;
  }




let items = [];
let originalImageUrls = [];
let dragSrcIndex = null;

const fileInput   = document.getElementById('fileInput');
const dropzone    = document.getElementById('dropzone');
const positionsEl = document.getElementById('positions');
const statusEl    = document.getElementById('status');
const errorBox    = document.getElementById('errorBox');
const downloadBtn = document.getElementById('downloadBtn');

function setStatus(msg){
  if(statusEl) statusEl.textContent = msg || "";
}
function setError(msg, err){
  if(errorBox) errorBox.textContent = msg || "";
  if(msg) console.error(msg, err || "");
}

// wyciąga wypełnienie z tekstu "Opis:" w G kolumnie
function extractFillFromOpis(text){
  if(!text) return "";
  const lines = String(text).split(/\r?\n/);
  for(const line of lines){
    if(line.toLowerCase().includes("wypełn")){
      const parts = line.split(":");
      if(parts.length >= 2){
        return parts.slice(1).join(":").replace(/_x000D_/g,"").trim();
      }
    }
  }
  return "";
}

// parser pozycji pod Twój Excel: kolumna F (index 5) – etykiety, G (6) – wartości
function parsePositions(rows){
  const result = [];
  let current = null;

  for(let i=0;i<rows.length;i++){
    const row = rows[i] || [];
    const label = row[5] != null ? String(row[5]).trim() : "";
    const value = row[6] != null ? String(row[6]).trim() : "";

    if(!label) continue;

    if(label.startsWith("Pozycja")){
      if(current) result.push(current);
      current = { number:value || "", name:"", qty:"", fill:"" };
      continue;
    }
    if(!current) continue;

    if(label.startsWith("Nazwa")){
      current.name = value;
    }else if(label.startsWith("Ilość")){
      current.qty = value;
    }else if(label.startsWith("Opis")){
      current.fill = extractFillFromOpis(value);
    }
  }
  if(current) result.push(current);
  return result;
}

async function handleFile(arrayBuffer, fileName){
  setError("");
  setStatus("Przetwarzam plik: " + fileName + " …");
  items = [];
  positionsEl.innerHTML = "";

  let wb;
  try{
    wb = XLSX.read(arrayBuffer, { type:"array" });
  }catch(e){
    setError("Błąd odczytu arkusza XLSX.", e);
    setStatus("Błąd.");
    return;
  }

  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:true });

  if(!rows || !rows.length){
    positionsEl.innerHTML = "<i>Arkusz jest pusty.</i>";
    setStatus("Brak danych w arkuszu.");
    return;
  }

  const positions = parsePositions(rows);
  if(!positions.length){
    positionsEl.innerHTML = "<i>Nie znaleziono żadnych pozycji w pliku.</i>";
    setStatus("Brak pozycji.");
    return;
  }

  
  
// wydobycie obrazków z xl/media – tylko po kolei
let imageUrls = [];
try{
  const zip = await JSZip.loadAsync(arrayBuffer);
  const mediaEntries = {};
  zip.forEach((path, entry)=>{
    const lower = path.toLowerCase();
    if(lower.startsWith("xl/media/") && !entry.dir){
      if(/\.(png|jpg|jpeg|gif|webp|bmp)$/i.test(lower)){
        mediaEntries[lower]=entry;
      }
    }
  });
  const keys=Object.keys(mediaEntries).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  for(const k of keys){
    const blob=await mediaEntries[k].async("blob");
    const url=URL.createObjectURL(blob);
    imageUrls.push(url);
  }
}catch(e){
  setError("Błąd odczytu obrazów",e);
}
items = positions.map((p, idx)=>({
    number: p.number,
    name: p.name,
    qty: p.qty,
    fill: p.fill,
    imageUrl: imageUrls[idx] || null
  }));

  let html = "";
  items.forEach(item=>{
    html += `
      <div class="position-box">
        <h3>${item.number ? ("Pozycja " + item.number) : ""} ${item.name || ""}</h3>
        <div class="qty">Ilość: <strong>${item.qty || ""}</strong></div>
        ${item.imageUrl ? `<img src="${item.imageUrl}" alt="Zdjęcie pozycji">`
                        : `<div class="no-image">Brak zdjęcia</div>`}
        <div class="fill">Wypełnienie: <strong>${item.fill || ""}</strong></div>
      </div>
    `;
  });
  positionsEl.innerHTML = html || "<i>Brak pozycji do wyświetlenia.</i>";

  const boxes = positionsEl.querySelectorAll(".position-box");
  boxes.forEach((box, idx)=>{
    box.dataset.index = idx;
    const img = box.querySelector("img");
    if(img){
      img.draggable = true;
    }
  });

  originalImageUrls = items.map(it => it.imageUrl);


  // render unique systems above positions
  const systems = [...new Set(items.map(i => extractSystem(i.name)))].filter(s=>s);
  const sysHTML = "<div class='systems-box'><strong>Systemy w ofercie:</strong> " + systems.join(", ") + "<span class=\"sys-date\"></span></div>";
  const posContainer = document.getElementById('positions');
  posContainer.insertAdjacentHTML('beforebegin', sysHTML);
  const sysDateEl = document.querySelector('.systems-box .sys-date');
  if(sysDateEl){ sysDateEl.textContent = new Date().toLocaleDateString('pl-PL'); sysDateEl.style.float='right'; sysDateEl.style.opacity='0.7'; }

  setStatus(`Wczytano plik: ${fileName}. Pozycji: ${items.length}, obrazków: ${imageUrls.length}.`);
}

// obsługa kliknięcia i drag&drop
if(dropzone && fileInput){
  // klik – wybór pliku
  dropzone.addEventListener("click", ()=>{
    fileInput.click();
  });

  // zmiana pliku
  fileInput.addEventListener("change", e=>{
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = ev=> handleFile(ev.target.result, file.name);
    reader.readAsArrayBuffer(file);
  });

  ["dragenter","dragover","dragleave","drop"].forEach(ev=>{
    dropzone.addEventListener(ev, e=>{
      e.preventDefault();
      e.stopPropagation();
    });
  });

  dropzone.addEventListener("dragover", ()=>{
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", ()=>{
    dropzone.classList.remove("dragover");
  });
  dropzone.addEventListener("drop", e=>{
    dropzone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = ev=> handleFile(ev.target.result, file.name);
    reader.readAsArrayBuffer(file);
  });
}

// przycisk GENERUJ OFERTĘ – do rozbudowy
if(downloadBtn){
  downloadBtn.addEventListener("click", ()=>{
    if(!items || !items.length){
      alert("Najpierw wczytaj pozycje z pliku XLSX.");
      return;
    }
    generujOferte();
  });
}

// ================= PREVIEW, RESET GLOBALNY, DRAG&DROP =================
(function(){
  const modal       = document.getElementById("imgModal");
  const modalImg    = document.getElementById("modalImg");
  const modalClose  = document.getElementById("modalClose");
  const resetAllBtn = document.getElementById("resetAllBtn");

  // podgląd zdjęcia po kliknięciu
  if(modal && modalImg && modalClose){
    document.addEventListener("click", e=>{
      const img = e.target.closest(".position-box img");
      if(img){
        if(!img.src) return;
        modalImg.src = img.src;
        modal.style.display = "flex";
        return;
      }
    });

    modalClose.addEventListener("click", ()=>{
      modal.style.display = "none";
    });

    modal.addEventListener("click", e=>{
      if(e.target === modal){
        modal.style.display = "none";
      }
    });
  }

  // globalny reset rysunków
  if(resetAllBtn){
    resetAllBtn.addEventListener("click", ()=>{
      if(!items || !items.length || !originalImageUrls.length) return;
      items.forEach((it, i)=>{
        it.imageUrl = originalImageUrls[i] || null;
      });
      if(positionsEl){
        const boxes = positionsEl.querySelectorAll(".position-box");
        boxes.forEach((box, i)=>{
          const img = box.querySelector("img");
          if(!img) return;
          const url = items[i] ? items[i].imageUrl : null;
          if(url){
            img.src = url;
          }else{
            img.removeAttribute("src");
          }
        });
      }
    });
  }

  // drag & drop obrazków
  if(positionsEl){
    positionsEl.addEventListener("dragstart", e=>{
      const img = e.target.closest(".position-box img");
      if(!img) return;
      const box = img.closest(".position-box");
      if(!box || box.dataset.index === undefined) return;
      dragSrcIndex = parseInt(box.dataset.index, 10);
      img.classList.add("dragging");
      if(e.dataTransfer){
        e.dataTransfer.effectAllowed = "move";
      }
    });

    positionsEl.addEventListener("dragend", e=>{
      const img = e.target.closest(".position-box img");
      if(img) img.classList.remove("dragging");
      positionsEl.querySelectorAll(".drop-target").forEach(el=>el.classList.remove("drop-target"));
      dragSrcIndex = null;
    });

    positionsEl.addEventListener("dragover", e=>{
      const box = e.target.closest(".position-box");
      if(!box) return;
      e.preventDefault();
      const img = box.querySelector("img");
      if(img) img.classList.add("drop-target");
      if(e.dataTransfer){
        e.dataTransfer.dropEffect = "move";
      }
    });

    positionsEl.addEventListener("dragleave", e=>{
      const box = e.target.closest(".position-box");
      if(!box) return;
      const img = box.querySelector("img");
      if(img) img.classList.remove("drop-target");
    });

    positionsEl.addEventListener("drop", e=>{
      const targetBox = e.target.closest(".position-box");
      if(!targetBox || dragSrcIndex === null) return;
      e.preventDefault();
      const dstIndex = parseInt(targetBox.dataset.index, 10);
      if(isNaN(dstIndex) || dstIndex === dragSrcIndex) return;

      const srcBox = positionsEl.querySelector(`.position-box[data-index="${dragSrcIndex}"]`);
      const srcImg = srcBox ? srcBox.querySelector("img") : null;
      const dstImg = targetBox.querySelector("img");
      if(!srcImg || !dstImg) return;

      const tmpSrc = srcImg.src;
      srcImg.src   = dstImg.src;
      dstImg.src   = tmpSrc;

      if(Array.isArray(items) && items[dragSrcIndex] && items[dstIndex]){
        const tmpUrl = items[dragSrcIndex].imageUrl;
        items[dragSrcIndex].imageUrl = items[dstIndex].imageUrl;
        items[dstIndex].imageUrl = tmpUrl;
      }

      positionsEl.querySelectorAll(".drop-target").forEach(el=>el.classList.remove("drop-target"));
      const dragging = positionsEl.querySelector(".dragging");
      if(dragging) dragging.classList.remove("dragging");
      dragSrcIndex = null;
    });
  }
})();


// ================= BACKEND DOCX GENERATION (MINIMAL, NIE RUSZA IMPORTU XLSX) ===============

function imgToBase64(imgEl){
  return new Promise((resolve)=>{
    if(!imgEl){
      resolve("");
      return;
    }
    const c = document.createElement("canvas");
    const w = imgEl.naturalWidth || imgEl.width;
    const h = imgEl.naturalHeight || imgEl.height;
    if(!w || !h){
      resolve("");
      return;
    }
    c.width = w;
    c.height = h;
    const ctx = c.getContext("2d");
    ctx.drawImage(imgEl, 0, 0, w, h);
    const dataUrl = c.toDataURL("image/png");
    const prefix = "data:image/png;base64,";
    if(dataUrl.startsWith(prefix)){
      resolve(dataUrl.substring(prefix.length));
    }else{
      resolve(dataUrl);
    }
  });
}

const BACKEND_URL = "https://ofertus-backend2.onrender.com";

async function generujOferte(){
  try{
    if(!items || !items.length){
      alert("Najpierw wczytaj pozycje z pliku XLSX.");
      return;
    }

    const today = new Date();

    let systemsSet = new Set();
    if(typeof extractSystem === "function"){
      items.forEach(it=>{
        const s = extractSystem(it.name);
        if(s) systemsSet.add(s);
      });
    }
    const systems = Array.from(systemsSet);

    const itemsPayload = [];
    const boxes = document.querySelectorAll(".position-box");
    for(let i=0; i<items.length; i++){
      const it = items[i];
      let imgB64 = "";
      const box = boxes[i];
      if(box){
        const imgEl = box.querySelector("img");
        imgB64 = await imgToBase64(imgEl);
      }
      itemsPayload.push({
        lp: it.number || "",
        nazwa_rysunek: it.name || "",
        ilosc: it.qty ? ("X" + it.qty) : "",
        opis: it.fill || "",
        image_base64: imgB64
      });
    }

    const payload = {
      data: today.toLocaleDateString("pl-PL"),
      numer_oferty: "",
      klient_imie: "",
      klient_email: "",
      klient_tel: "",
      lokalizacja_obiektu: "",
      systemy: systems,
      kolor: "",
      kwota_netto: "",
      handlowiec_imie: "",
      handlowiec_tel: "",
      handlowiec_mail: "",
      items: itemsPayload
    };

    const fd = new FormData();
    fd.append("payload", JSON.stringify(payload));

    const res = await fetch(BACKEND_URL + "/generate-docx", {
      method: "POST",
      body: fd
    });

    if(!res.ok){
      alert("Błąd generowania oferty (backend).");
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Oferta_" + (payload.numer_oferty || "DOPLER") + ".docx";
    a.click();
  }catch(e){
    console.error(e);
    alert("Nie udało się wygenerować oferty.");
  }
}

