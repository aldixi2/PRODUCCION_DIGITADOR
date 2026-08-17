const $ = (id) => document.getElementById(id);

function escapeHtml(s){
  return String(s ?? "").replace(/[&<>"']/g, (m)=>({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[m]));
}
function normalizarCodigo(v){
  const m = String(v||"").match(/\d+/);
  return m ? m[0].padStart(3,"0") : "";
}
function pickCode(value){
  return String(value||"").split("—")[0].trim();
}

let REGLAS = {};        // { "018": { cpms_allowed:[...], dx_hint:[...] }, ... }
let CPMS_CATALOGO = {}; // { "99402.04": "Orientación...", ... }

async function cargarData(){
  try{
    const r1 = await fetch("data/reglas.json?ts="+Date.now());
    REGLAS = await r1.json();

    const r2 = await fetch("data/cpms_catalog.json?ts="+Date.now());
    CPMS_CATALOGO = await r2.json();
  }catch(e){
    console.error("No se pudo cargar data JSON:", e);
  }
}

function cpmsLabel(code){
  const name = CPMS_CATALOGO[code];
  return name ? `${code} — ${name}` : `${code}`;
}

function buildDatalist(id, codes){
  const el = document.getElementById(id);
  if(!el) return;
  el.innerHTML = (codes||[]).map(c => `<option value="${escapeHtml(cpmsLabel(c))}"></option>`).join("");
}

const btnVer = $("btnVer");
const btnLimpiar = $("btnLimpiar");
const codServicio = $("codServicio");
const descServicio = $("descServicio");
const panel = $("panelResultado");
const resTitulo = $("resTitulo");
const resSub = $("resSub");
const resBody = $("resBody");
const btnImprimir = $("btnImprimir");
const btnCerrarPanel = $("btnCerrarPanel");

if(btnCerrarPanel) btnCerrarPanel.onclick = ()=> panel.style.display="none";
if(btnImprimir) btnImprimir.onclick = ()=> window.print();

document.querySelectorAll("#chipsRapidos .chip").forEach(ch=>{
  ch.addEventListener("click", ()=>{
    codServicio.value = ch.dataset.code || "";
    mostrarEjemplo();
  });
});

if(btnVer) btnVer.onclick = mostrarEjemplo;
if(btnLimpiar){
  btnLimpiar.onclick = ()=>{
    codServicio.value="";
    descServicio.value="";
    panel.style.display="none";
    codServicio.focus();
  };
}

// ====== Render FUA genérica REAL (con CPMS controlado por Excel) ======
function renderFUA(code){
  const reglas = REGLAS[code] || null;
  const cpmsAllowed = reglas?.cpms_allowed || [];

  // datalist CPMS
  setTimeout(()=> buildDatalist("dl_proc", cpmsAllowed), 0);

  const pdfPath = `pdf/${code}.pdf`;

  return `
  <div class="fuaSheet">

    <div class="fuaHeader">
      <div>
        <div class="fuaTitle">GUÍA FUA — PRESTACIÓN ${escapeHtml(code)}</div>
        <div class="fuaMeta">
          ${descServicio.value.trim() ? escapeHtml(descServicio.value.trim()) : "Seleccione códigos válidos (sin inventar)."}
        </div>
        <div class="fuaMeta">
          ${cpmsAllowed.length
            ? `<span class="lock">✅ CPMS permitidos: ${cpmsAllowed.length}</span>`
            : `<span class="lock">⚠️ Sin reglas CPMS en Excel (se mostrará PDF si existe)</span>`
          }
        </div>
      </div>

      <div class="boxRight">
        <div>
          <div class="cap">PDF oficial</div>
          <div class="val"><a href="${pdfPath}" target="_blank">Abrir ${code}.pdf</a></div>
        </div>
        <div>
          <div class="cap">Impresión</div>
          <div class="val">Usa “Imprimir”</div>
        </div>
      </div>
    </div>

    <div class="bloque" style="margin-top:8px">
      <div class="bloqueT">DIAGNÓSTICO (CIE10)</div>
      <div class="bloqueB">
        <div class="note">
          (Siguiente paso PRO+): aquí también cargaremos tu catálogo CIE10 para que elijan “código + nombre”.
        </div>
      </div>
    </div>

    <div class="bloque" style="margin-top:8px">
      <div class="bloqueT">PROCEDIMIENTOS (CPMS) — SOLO CÓDIGOS PERMITIDOS</div>
      <div class="bloqueB">

        ${cpmsAllowed.length ? `
          <div class="row" style="gap:8px; align-items:center; margin-top:0">
            <input class="inlineInput" id="procInput" list="dl_proc" placeholder="Elige CPMS: código — nombre">
            <datalist id="dl_proc"></datalist>
            <button class="smallBtn" id="btnAddProc" type="button">+ Agregar</button>
          </div>

          <div class="tableBox" style="margin-top:10px">
            <table id="tblProc">
              <thead>
                <tr><th style="width:160px">CPMS</th><th>Nombre</th><th style="width:90px">Quitar</th></tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>

          <div class="note">✅ Si no está en la lista, NO se registra.</div>
        ` : `
          <div class="note">
            Este servicio aún no tiene reglas CPMS en el Excel. Usa el PDF oficial arriba.
          </div>
        `}
      </div>
    </div>

    <div class="bloque" style="margin-top:8px">
      <div class="bloqueT">CHECKLIST</div>
      <div class="bloqueB">
        <ul class="list">
          <li>Usar únicamente CPMS permitidos por la prestación.</li>
          <li>Verificar en el PDF oficial si aplica diagnóstico / medicamentos / insumos.</li>
        </ul>
      </div>
    </div>

  </div>
  `;
}

function bindCPMS(code){
  const reglas = REGLAS[code];
  if(!reglas?.cpms_allowed?.length) return;

  const tbody = document.querySelector("#tblProc tbody");
  const allowedSet = new Set(reglas.cpms_allowed);

  function addProc(codePicked){
    const cp = pickCode(codePicked);
    if(!cp) return;
    if(!allowedSet.has(cp)){
      alert("Ese CPMS no está permitido para esta prestación.");
      return;
    }
    const exists = [...tbody.querySelectorAll("tr")].some(tr => tr.dataset.code === cp);
    if(exists) return;

    const tr = document.createElement("tr");
    tr.dataset.code = cp;
    tr.innerHTML = `
      <td><b>${escapeHtml(cp)}</b></td>
      <td>${escapeHtml(CPMS_CATALOGO[cp] || "")}</td>
      <td><button class="smallBtn" type="button">X</button></td>
    `;
    tr.querySelector("button").onclick = ()=> tr.remove();
    tbody.appendChild(tr);
  }

  $("btnAddProc").onclick = ()=>{
    addProc($("procInput").value.trim());
    $("procInput").value = "";
  };
}

function mostrarEjemplo(){
  const code = normalizarCodigo(codServicio.value);
  if(!code){
    alert("Escribe un código. Ej: 018");
    codServicio.focus();
    return;
  }

  resTitulo.textContent = `Guía por prestación: ${code}`;
  resSub.textContent = "Reglas cargadas desde Excel (Nivel PRO).";

  resBody.innerHTML = renderFUA(code);
  panel.style.display = "";
  panel.scrollIntoView({behavior:"smooth", block:"start"});

  bindCPMS(code);
}

// Cargar data al inicio
cargarData();
