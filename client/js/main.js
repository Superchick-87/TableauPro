/* client/js/main.js - VERSION V106 : INTELLIGENT HYBRID MENUS */

// --- 1. VARIABLES ---
var fullRawData = [];
var loadedData = [];
var headers = [];
var columnOrder = [];
var hiddenOriginalIndices = [];

var activeRules = [];
var activeSorts = {};   
var activeFilters = {}; 

var selectedUiIndices = [];
var lastClickedIndex = -1;
var isDragging = false;
var globalColAligns = {};

var isUserInteracting = false;
var isInit = true;
var interactTimer;
var savedHeaderHeight = 15; 

// VARIABLES SYNC
var lastSelectionData = "";
var isScriptUpdating = false;
var isFetching = false;
var updateTimeout = null;

function userAction() {
    if (isInit || isScriptUpdating) return;
    isUserInteracting = true;
    clearTimeout(interactTimer);
    interactTimer = setTimeout(function() { isUserInteracting = false; saveSettings(); }, 2000);
}

// --- 2. DEMARRAGE ---
window.onload = function() {
    console.log("Init V106 Hybrid...");
    setupContainerDragEvents();
    setupCriticalUI();
    try { initTabs(); initAllColors(); initInputs(); loadFonts(); } catch(e) { console.error("Erreur UI:", e); }
    try { loadSettings(); } catch(e) { localStorage.removeItem("TableauPro_Settings"); }
    startSelectionWatcher();
    setTimeout(function(){ isInit = false; }, 500);
};

// --- 3. UI CRITIQUE ---
function setupCriticalUI() {
    var btnImport = document.getElementById('btnImport');
    var fInput = document.getElementById('fileInput');
    var btnApply = document.getElementById('btnApply');
    var btnRaz = document.getElementById('btnRaz');

    // ETAT INITIAL
    var cbTitle = document.getElementById('cbShowTitle');
    var cbLegend = document.getElementById('cbShowLegend');
    
    if (cbTitle) { cbTitle.checked = false; cbTitle.disabled = true; }
    if (cbLegend) { cbLegend.checked = false; cbLegend.disabled = true; }
    if (btnApply) { btnApply.disabled = true; }

    if(btnApply) {
        btnApply.onclick = applyToIllustrator;
    }

    if(btnImport && fInput) {
        btnImport.onclick = function() { fInput.click(); };
        fInput.onchange = function(e) {
            var f = e.target.files[0]; if(!f) return;
            document.getElementById('fileNameDisplay').innerText = f.name;
            Papa.parse(f, { 
                skipEmptyLines: true, 
                complete: function(r) {
                    if(r.errors.length > 0) { alert("Erreur lecture CSV"); return; }
                    fullRawData = r.data; 
                    headers = fullRawData[0];
                    columnOrder = [];
                    if(fullRawData.length > 0) { fullRawData[0].forEach((_,i) => columnOrder.push(i)); }
                    hiddenOriginalIndices = []; selectedUiIndices = [];
                    
                    processDataRange(); updateFilterUI(); updateRuleBuilderUI();
                    
                    if(btnApply) btnApply.disabled = false;
                    
                    if (cbTitle) { 
                        cbTitle.disabled = false; 
                        cbTitle.checked = true; 
                        cbTitle.dispatchEvent(new Event('change')); 
                    }
                    
                    var txtLegend = document.getElementById('legendText');
                    if (txtLegend && cbLegend) {
                        var hasText = txtLegend.value.trim() !== "";
                        cbLegend.disabled = !hasText; 
                        cbLegend.checked = hasText;   
                    }
                }
            });
        };
    }
    var btnAddRule = document.getElementById('btnAddRule');
    if(btnAddRule) btnAddRule.onclick = addRule;
    if(btnRaz) btnRaz.onclick = resetDatas;
}

// --- 4. GESTION DRAG ---
function setupContainerDragEvents() {
    var lst = document.getElementById('colList');
    if(!lst) return;
    lst.addEventListener('dragover', function(e) {
        e.preventDefault(); e.dataTransfer.dropEffect = 'move';
        var draggingItems = lst.querySelectorAll('.col-item.ui-selected');
        if(draggingItems.length === 0) return;
        var siblings = [...lst.querySelectorAll('.col-item:not(.ui-selected)')];
        var nextSibling = siblings.find(sibling => { return e.clientY <= sibling.getBoundingClientRect().top + sibling.offsetHeight / 2; });
        draggingItems.forEach(function(item) { if(nextSibling) lst.insertBefore(item, nextSibling); else lst.appendChild(item); });
    });
    lst.addEventListener('drop', function(e) {
        e.preventDefault(); reorderColumnsData(); isDragging = false;
        document.querySelectorAll('.dragging').forEach(el => el.classList.remove('dragging'));
    });
}

// --- 5. LOGIQUE UI COLONNES ---
function generateColUI(preserveState) {
    var lst = document.getElementById('colList'); if(!lst) return; 
    lst.innerHTML = ""; 
    function isUiSelected(idx) { return selectedUiIndices.includes(idx); }

    columnOrder.forEach(function(realIndex, visualIndex){
        var h = headers[realIndex]; 
        var safe = String(h).replace(/^"|"$/g,'').trim();
        var isChecked = !hiddenOriginalIndices.includes(realIndex); 
        if (!globalColAligns[safe]) globalColAligns[safe] = "left"; 
        var alignState = globalColAligns[safe];

        var el = document.createElement('div'); el.className = 'col-item';
        if (isUiSelected(visualIndex)) el.classList.add('ui-selected'); 
        
        el.setAttribute('draggable', 'true'); 
        el.setAttribute('data-vidx', visualIndex); 
        el.setAttribute('data-realidx', realIndex); 
        el.setAttribute('data-colname', safe);

        var checkAttr = isChecked ? 'checked' : '';
        var clL = (alignState === 'left') ? 'selected' : ''; 
        var clC = (alignState === 'center') ? 'selected' : ''; 
        var clR = (alignState === 'right') ? 'selected' : '';
        
        el.innerHTML = `
            <div class="drag-handle" style="cursor:grab;">☰</div>
            <input type="checkbox" ${checkAttr} style="cursor:pointer;">
            <span class="col-name" title="${safe}">${safe}</span>
            <span class="edit-icon" title="Renommer">✎</span>
            <div class="align-grp">
                <div class="align-btn ${clL}" data-al="left">L</div>
                <div class="align-btn ${clC}" data-al="center">C</div>
                <div class="align-btn ${clR}" data-al="right">R</div>
            </div>`;
        
        el.addEventListener('click', function(e) {
            if (isDragging) return;
            if (e.target.tagName === 'INPUT' || e.target.classList.contains('align-btn') || e.target.classList.contains('edit-icon') || e.target.classList.contains('drag-handle') || e.target.tagName === 'TEXTAREA') return;
            var currentIdx = parseInt(el.getAttribute('data-vidx'));
            if (e.shiftKey && lastClickedIndex !== -1) {
                var start = Math.min(lastClickedIndex, currentIdx); var end = Math.max(lastClickedIndex, currentIdx);
                selectedUiIndices = []; for(var i=start; i<=end; i++) selectedUiIndices.push(i);
            } else if (e.ctrlKey || e.metaKey) {
                if (selectedUiIndices.includes(currentIdx)) selectedUiIndices = selectedUiIndices.filter(id => id !== currentIdx);
                else selectedUiIndices.push(currentIdx);
                lastClickedIndex = currentIdx;
            } else {
                selectedUiIndices = [currentIdx]; lastClickedIndex = currentIdx;
            }
            generateColUI(true);
        });

        el.addEventListener('dragstart', function(e) {
            isDragging = true;
            var currentIdx = parseInt(el.getAttribute('data-vidx'));
            if (!selectedUiIndices.includes(currentIdx)) {
                selectedUiIndices = [currentIdx];
                lst.querySelectorAll('.col-item').forEach(it => it.classList.remove('ui-selected'));
                el.classList.add('ui-selected');
            }
            lst.querySelectorAll('.col-item.ui-selected').forEach(function(item) { item.classList.add('dragging'); });
            e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', safe);
            userAction();
        });
        el.addEventListener('dragend', function() { document.querySelectorAll('.dragging').forEach(el => el.classList.remove('dragging')); isDragging = false; });

        var cb = el.querySelector('input[type="checkbox"]');
        cb.onclick = function(e) {
            var targetState = this.checked;
            var clickedVIdx = parseInt(el.getAttribute('data-vidx'));
            var itemsToProcess = isUiSelected(clickedVIdx) ? selectedUiIndices : [clickedVIdx];
            itemsToProcess.forEach(function(vIdx) {
                var domItem = lst.querySelector(`.col-item[data-vidx="${vIdx}"]`);
                if (domItem) {
                    domItem.querySelector('input[type="checkbox"]').checked = targetState;
                    var rIdx = parseInt(domItem.getAttribute('data-realidx'));
                    if (targetState) { hiddenOriginalIndices = hiddenOriginalIndices.filter(id => id !== rIdx); } 
                    else { if (!hiddenOriginalIndices.includes(rIdx)) hiddenOriginalIndices.push(rIdx); }
                }
            });
            processDataRange(); userAction();
        };

        var editIcon = el.querySelector('.edit-icon');
        editIcon.onclick = function(e) {
            e.stopPropagation();
            var spanName = el.querySelector('.col-name');
            var currentName = spanName.innerText;
            var input = document.createElement('input');
            input.type = 'text'; input.value = currentName; input.className = 'col-name-input';
            spanName.replaceWith(input); input.focus();
            function validateRename() {
                var newName = input.value.trim();
                if(newName && newName !== currentName) {
                    headers[realIndex] = newName;
                    if(globalColAligns[currentName]) { globalColAligns[newName] = globalColAligns[currentName]; }
                    processDataRange(); updateFilterUI(); updateRuleBuilderUI(); userAction();
                }
                generateColUI(true); 
            }
            input.onblur = validateRename;
            input.onkeydown = function(ev) { if(ev.key === 'Enter') { input.blur(); } };
            input.onclick = function(ev) { ev.stopPropagation(); };
        };

        el.querySelectorAll('.align-btn').forEach(function(b){
            b.onclick = function(e) {
                e.stopPropagation();
                var newAlign = this.getAttribute('data-al');
                var clickedVIdx = parseInt(el.getAttribute('data-vidx'));
                var itemsToProcess = isUiSelected(clickedVIdx) ? selectedUiIndices : [clickedVIdx];
                itemsToProcess.forEach(function(vIdx) {
                    var domItem = lst.querySelector(`.col-item[data-vidx="${vIdx}"]`);
                    if(domItem) {
                        var cName = domItem.getAttribute('data-colname');
                        globalColAligns[cName] = newAlign;
                        domItem.querySelectorAll('.align-btn').forEach(btn => btn.classList.remove('selected'));
                        domItem.querySelector(`.align-btn[data-al="${newAlign}"]`).classList.add('selected');
                    }
                });
                userAction();
            };
        });

        lst.appendChild(el);
    });
}

function reorderColumnsData() {
    var lst = document.getElementById('colList');
    var newVisualOrder = []; var newHiddenList = [];
    Array.from(lst.children).forEach(function(item) {
        var rIdx = parseInt(item.getAttribute('data-realidx'));
        if (!isNaN(rIdx)) {
            newVisualOrder.push(rIdx);
            if (!item.querySelector('input[type="checkbox"]').checked) { newHiddenList.push(rIdx); }
        }
    });
    columnOrder = newVisualOrder; hiddenOriginalIndices = newHiddenList;
    selectedUiIndices = []; 
    processDataRange(); updateFilterUI(); userAction(); generateColUI(true);
}

// --- 6. DATA PROCESSING ---
function processDataRange() {
    if (fullRawData.length === 0) return;
    
    var allRows = fullRawData.slice(1);

    // 1. PRE-CALCUL MIN/MAX
    var colLimits = {};
    for (var colName in activeFilters) {
        var f = activeFilters[colName];
        if (f.op === 'min' || f.op === 'max') {
            var tIdx = headers.indexOf(colName);
            if (tIdx > -1) {
                var valList = allRows.map(r => {
                    return parseFloat(String(r[tIdx]||"").replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
                }).filter(n => !isNaN(n));
                if (valList.length > 0) {
                    if (f.op === 'min') colLimits[colName] = Math.min(...valList);
                    else colLimits[colName] = Math.max(...valList);
                }
            }
        }
    }

    // 2. FILTRAGE GLOBAL
    var filteredRows = allRows.filter(function(row) {
        for (var colName in activeFilters) {
            var filter = activeFilters[colName];
            if (!filter || filter.op === 'none') continue;

            var targetIdx = headers.indexOf(colName);
            if (targetIdx === -1) continue;

            var cellVal = row[targetIdx] || "";
            var numCell = parseFloat(String(cellVal).replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
            var numA = parseFloat(String(filter.val).replace(',', '.'));
            
            var match = true;

            if (filter.op === 'not_empty') {
                if (String(cellVal).trim() === "") match = false;
            }
            else if (filter.op === 'min' || filter.op === 'max') {
                if (colLimits[colName] !== undefined) {
                    if (numCell !== colLimits[colName]) match = false;
                } else { match = false; }
            }
            else if (filter.op === '=') { if (String(cellVal).toLowerCase() !== String(filter.val).toLowerCase()) match = false; }
            else if (filter.op === 'contains') { if (String(cellVal).toLowerCase().indexOf(String(filter.val).toLowerCase()) === -1) match = false; }
            else if (['>', '>=', '<', '<=', 'between'].includes(filter.op)) {
                if (isNaN(numCell)) match = false;
                else {
                    if (filter.op === '>') { if (!(numCell > numA)) match = false; }
                    else if (filter.op === '>=') { if (!(numCell >= numA)) match = false; }
                    else if (filter.op === '<') { if (!(numCell < numA)) match = false; }
                    else if (filter.op === '<=') { if (!(numCell <= numA)) match = false; }
                    else if (filter.op === 'between') {
                        var numB = parseFloat(String(filter.valB).replace(',', '.'));
                        if (!(numCell >= Math.min(numA, numB) && numCell <= Math.max(numA, numB))) match = false;
                    }
                }
            }
            if (!match) return false;
        } 
        return true;
    });

    // 3. PAGINATION
    var inputStart = document.getElementById('rowStart'); 
    var valNbLines = document.getElementById('valNbLines');
    var cbNbLines = document.getElementById('cbNbLines');
    var start = parseInt(inputStart.value) || 0;
    var nbLines = parseInt(valNbLines.value) || 0;
    
    var visibleRows = filteredRows;
    if (start > 0) visibleRows = visibleRows.slice(Math.max(0, start - 1));
    if (cbNbLines && cbNbLines.checked) { if (nbLines > 0) visibleRows = visibleRows.slice(0, nbLines); } 
    else {
        var inputEnd = document.getElementById('rowEnd'); var end = parseInt(inputEnd.value) || 0;
        if (end > 0 && end >= start) visibleRows = visibleRows.slice(0, (end - start) + 1);
    }

    // 4. TRI LOCAL
    for (var colName in activeSorts) {
        var dir = activeSorts[colName];
        var c = headers.indexOf(colName);
        if (dir && c > -1) {
            visibleRows.sort(function(a, b) {
                var valA = a[c] || ""; var valB = b[c] || "";
                var numA = parseFloat(valA.replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, '')); 
                var numB = parseFloat(valB.replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
                if (!isNaN(numA) && !isNaN(numB)) { return (dir === 'asc') ? numA - numB : numB - numA; }
                return (dir === 'asc') ? valA.localeCompare(valB) : valB.localeCompare(valA);
            });
        }
    }

    // 5. OUTPUT
    var finalVisualCols = []; var finalColAligns = [];
    columnOrder.forEach(function(realIndex) {
        if (!hiddenOriginalIndices.includes(realIndex)) {
            finalVisualCols.push(realIndex);
            var colName = headers[realIndex]; 
            finalColAligns.push(globalColAligns[colName] || "left");
        }
    });

    var finalData = [];
    var headerRow = finalVisualCols.map(idx => headers[idx]); finalData.push(headerRow);
    visibleRows.forEach(function(row) { 
        var newRow = finalVisualCols.map(idx => row[idx]); finalData.push(newRow); 
    });

    loadedData = finalData; 
    if(!isDragging) generateColUI(true);
}

// --- 7. UI FILTRES (LOGIQUE HYBRIDE V103) ---
function updateFilterUI() {
    var container = document.getElementById('filterContainer'); 
    if(!container) return; 
    
    container.innerHTML = "";
    container.style.maxHeight = "450px"; 
    container.style.overflowY = "auto";
    container.style.overflowX = "hidden";
    container.style.paddingRight = "5px"; 
    
    var allSourceRows = fullRawData.slice(1);
    var visibleSourceRows = loadedData.slice(1);

    columnOrder.forEach(function(realIndex, visualIndex) {
        var colName = headers[realIndex]; 
        
        // Si masqué, on ignore
        if (hiddenOriginalIndices.includes(realIndex)) return;

        // Map CSV index -> Visual index
        var loadedDataColIndex = -1;
        var currentLoadedIdx = 0;
        columnOrder.forEach(function(rIdx) {
            if (!hiddenOriginalIndices.includes(rIdx)) {
                if (rIdx === realIndex) loadedDataColIndex = currentLoadedIdx;
                currentLoadedIdx++;
            }
        });

        if (loadedDataColIndex === -1) return;

        var item = document.createElement('div'); 
        item.className = 'filter-item';
        item.style.borderBottom = "1px solid #444";
        item.style.paddingBottom = "5px";
        item.style.marginBottom = "5px";

        var label = document.createElement('div'); 
        label.className = 'filter-header'; 
        label.innerText = colName; 
        item.appendChild(label);

        // --- CASCADING LOGIC HYBRIDE ---
        var uniqueVals = new Set();
        var isFiltered = (activeFilters[colName] !== undefined && activeFilters[colName].op !== 'none');

        if (isFiltered) {
            // CAS 1: CETTE COLONNE EST DÉJÀ FILTRÉE (ex: Département = 16)
            // On veut permettre de CHANGER de département, donc on affiche TOUS les départements dispos.
            // On prend TOUT le fichier, on applique les filtres des AUTRES colonnes, mais on s'ignore.
            
            allSourceRows.forEach(function(row) {
                var isVisible = true;
                
                for (var otherCol in activeFilters) {
                    if (otherCol === colName) continue; // IGNORE MON PROPRE FILTRE
                    
                    var filter = activeFilters[otherCol];
                    if (!filter || filter.op === 'none') continue;
    
                    var targetIdx = headers.indexOf(otherCol);
                    if (targetIdx === -1) continue;
    
                    var cellVal = row[targetIdx] || "";
                    var numCell = parseFloat(String(cellVal).replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
                    var numA = parseFloat(String(filter.val).replace(',', '.'));
                    var match = true;
    
                    if (filter.op === 'not_empty') { if (String(cellVal).trim() === "") match = false; }
                    else if (filter.op === '=') { if (String(cellVal).toLowerCase() !== String(filter.val).toLowerCase()) match = false; }
                    else if (filter.op === 'contains') { if (String(cellVal).toLowerCase().indexOf(String(filter.val).toLowerCase()) === -1) match = false; }
                    else if (['>', '>=', '<', '<=', 'between'].includes(filter.op)) {
                        if (isNaN(numCell)) match = false;
                        else {
                            if (filter.op === '>') { if (!(numCell > numA)) match = false; }
                            else if (filter.op === '>=') { if (!(numCell >= numA)) match = false; }
                            else if (filter.op === '<') { if (!(numCell < numA)) match = false; }
                            else if (filter.op === '<=') { if (!(numCell <= numA)) match = false; }
                            else if (filter.op === 'between') {
                                var numB = parseFloat(String(filter.valB).replace(',', '.'));
                                if (!(numCell >= Math.min(numA, numB) && numCell <= Math.max(numA, numB))) match = false;
                            }
                        }
                    }
                    if (!match) isVisible = false;
                }
    
                if (isVisible) {
                    var val = row[realIndex];
                    if (val !== undefined && val !== "") uniqueVals.add(val);
                }
            });

        } else {
            // CAS 2: CETTE COLONNE N'EST PAS FILTRÉE (ex: Ville, alors qu'on a déjà filtré Dept=16)
            // On ne veut voir QUE ce qui est affiché à l'écran (les 10 villes du 16).
            // On prend loadedData (données visibles).
            
            visibleSourceRows.forEach(function(row) {
                var val = row[loadedDataColIndex];
                if (val !== undefined && val !== "") uniqueVals.add(val);
            });
        }

        var sortedVals = Array.from(uniqueVals).sort(function(a, b) {
            return String(a).localeCompare(String(b), undefined, {numeric: true, sensitivity: 'base'});
        });

        var controls = document.createElement('div');
        controls.style.display = 'flex';
        controls.style.flexDirection = 'column';
        controls.style.gap = '4px';

        // A. TRI
        var sortSelect = document.createElement('select');
        sortSelect.className = 'filter-select';
        sortSelect.style.marginBottom = "2px";
        sortSelect.innerHTML = `
            <option value="none">-- Ordre (Aucun) --</option>
            <option value="asc">Trier A -> Z (Visible)</option>
            <option value="desc">Trier Z -> A (Visible)</option>
        `;
        if (activeSorts[colName]) { sortSelect.value = activeSorts[colName]; }
        sortSelect.onchange = function() {
            var dir = this.value;
            if (dir === 'none') delete activeSorts[colName];
            else { activeSorts = {}; activeSorts[colName] = dir; } 
            processDataRange(); updateFilterUI(); userAction();
        };

        // B. FILTRE
        var filterSelect = document.createElement('select');
        filterSelect.className = 'filter-select';
        
        var ops = [
            {v:'none', t:'-- Tout (Annuler filtre) --'},
            {v:'not_empty', t:'Différent de vide'},
            {v:'min', t:'Minimum (Auto)'}, 
            {v:'max', t:'Maximum (Auto)'}, 
            {v:'contains', t:'Contient...'},
            {v:'=', t:'Est égal à (=) ...'},
            {v:'>', t:'Supérieur à (>)'},
            {v:'>=', t:'Sup. ou égal (>=)'},
            {v:'<', t:'Inférieur à (<)'},
            {v:'<=', t:'Inf. ou égal (<=)'},
            {v:'between', t:'Entre (min/max)'}
        ];
        ops.forEach(o => { var opt = document.createElement('option'); opt.value = o.v; opt.text = o.t; filterSelect.add(opt); });

        var sep = document.createElement('option'); sep.disabled = true; sep.text = "───────"; filterSelect.add(sep);

        var grpVal = document.createElement('optgroup'); 
        grpVal.label = isFiltered ? "Changer de valeur (Global)" : "Valeurs Affichées";
        
        sortedVals.forEach(function(v) { 
            var opt = document.createElement('option'); 
            opt.value = "u_" + v; 
            opt.text = v; 
            grpVal.appendChild(opt); 
        });
        filterSelect.add(grpVal);

        var inputA = document.createElement('input');
        inputA.type = "text"; inputA.placeholder = "Valeur..."; inputA.style.width = "100%"; inputA.style.display = "none";
        
        var inputB = document.createElement('input');
        inputB.type = "text"; inputB.placeholder = "Max..."; inputB.style.width = "100%"; inputB.style.display = "none";

        var f = activeFilters[colName];
        if (f) {
            if (!f.isUniqueSelect) {
                filterSelect.value = f.op;
                inputA.value = f.val || "";
                if (f.valB) inputB.value = f.valB;
                if (['none', 'min', 'max', 'not_empty'].includes(f.op)) {
                    inputA.style.display = 'none';
                } else {
                    inputA.style.display = 'block';
                    if (f.op === 'between') inputB.style.display = 'block';
                }
            } else {
                filterSelect.value = "u_" + f.val;
                inputA.style.display = 'none'; 
            }
        }

        function updateFilterState() {
            var rawVal = filterSelect.value;
            
            if (['none', 'min', 'max', 'not_empty'].includes(rawVal) || rawVal.startsWith('u_')) {
                inputA.style.display = 'none'; inputB.style.display = 'none';
            } else {
                inputA.style.display = 'block';
                if (rawVal === 'between') {
                    inputA.placeholder = "Min..."; inputB.style.display = 'block';
                } else {
                    inputA.placeholder = "Valeur..."; inputB.style.display = 'none';
                }
            }

            if (rawVal === 'none') {
                delete activeFilters[colName];
            }
            else if (['min', 'max', 'not_empty'].includes(rawVal)) {
                activeFilters[colName] = { op: rawVal, isUniqueSelect: false };
            } 
            else if (rawVal.startsWith('u_')) {
                var realVal = rawVal.substring(2); 
                activeFilters[colName] = { op: '=', val: realVal, isUniqueSelect: true };
            } 
            else {
                var filterObj = { op: rawVal, val: inputA.value, isUniqueSelect: false };
                if (rawVal === 'between') filterObj.valB = inputB.value;
                activeFilters[colName] = filterObj;
            }
            processDataRange(); userAction();
            setTimeout(updateFilterUI, 10); 
        }

        filterSelect.onchange = updateFilterState;
        inputA.oninput = updateFilterState;
        inputB.oninput = updateFilterState;

        controls.appendChild(sortSelect);
        controls.appendChild(filterSelect);
        controls.appendChild(inputA);
        controls.appendChild(inputB);
        item.appendChild(controls);
        container.appendChild(item);
    });
}

// --- 8. REGLES & EDIT ---
function updateRuleBuilderUI() {
    var sel = document.getElementById('ruleCol'); if(!sel) return; sel.innerHTML = "";
    headers.forEach(function(h) { var opt = document.createElement('option'); opt.value = h; opt.text = h; sel.add(opt); });
    var opSel = document.getElementById('ruleOp'); var valCont = document.getElementById('ruleValContainer'); var valBet = document.getElementById('ruleValBetween');
    opSel.onchange = function() { if(this.value === 'between') { valCont.style.display='none'; valBet.style.display='flex'; } else { valCont.style.display='flex'; valBet.style.display='none'; } };
}

function addRule() {
    var colName = document.getElementById('ruleCol').value; var op = document.getElementById('ruleOp').value;
    var scope = document.querySelector('input[name="ruleScope"]:checked').value;
    var bg = getCMYK('ui-ruleBg'); var txt = getCMYK('ui-ruleTxt');
    var val = document.getElementById('ruleVal').value; var valA = document.getElementById('ruleValA').value; var valB = document.getElementById('ruleValB').value;
    var desc = colName + " " + op + " "; var ruleData = { colName: colName, op: op, scope: scope, bg: bg, txt: txt };
    if(op === 'between') { if(valA === "" || valB === "") return; ruleData.valA = parseFloat(valA); ruleData.valB = parseFloat(valB); desc += valA + " et " + valB; } else { if(val === "") return; ruleData.val = val; desc += val; }
    ruleData.desc = desc; activeRules.push(ruleData); renderRules(); userAction();
}

function renderRules() {
    var list = document.getElementById('rulesList'); list.innerHTML = "";
    activeRules.forEach(function(r, i) {
        var div = document.createElement('div'); div.className = 'rule-item';
        div.innerHTML = `<span class="rule-desc">${r.desc} (${r.scope==='row'?'Ligne':'Cel.'})</span>
                         <div class="rule-actions">
                             <span class="rule-btn move-up" title="Monter">▲</span>
                             <span class="rule-btn move-down" title="Descendre">▼</span>
                             <span class="rule-btn edit" title="Modifier">✎</span>
                             <span class="rule-btn delete" title="Supprimer">✕</span>
                         </div>`;
        
        div.querySelector('.delete').onclick = function() { activeRules.splice(i, 1); renderRules(); userAction(); };
        div.querySelector('.edit').onclick = function() {
            var rule = activeRules[i];
            document.getElementById('ruleCol').value = rule.colName;
            document.getElementById('ruleOp').value = rule.op;
            document.getElementById('ruleOp').dispatchEvent(new Event('change')); 
            if(rule.op === 'between') { document.getElementById('ruleValA').value = rule.valA; document.getElementById('ruleValB').value = rule.valB; } 
            else { document.getElementById('ruleVal').value = rule.val; }
            document.querySelectorAll('input[name="ruleScope"]').forEach(rad => { if(rad.value === rule.scope) rad.checked=true; });
            setCMYK('ui-ruleBg', rule.bg); setCMYK('ui-ruleTxt', rule.txt);
            activeRules.splice(i, 1); renderRules(); userAction();
        };
        div.querySelector('.move-up').onclick = function() { if(i > 0) { var tmp = activeRules[i]; activeRules[i] = activeRules[i-1]; activeRules[i-1] = tmp; renderRules(); userAction(); } };
        div.querySelector('.move-down').onclick = function() { if(i < activeRules.length - 1) { var tmp = activeRules[i]; activeRules[i] = activeRules[i+1]; activeRules[i+1] = tmp; renderRules(); userAction(); } };
        list.appendChild(div);
    });
}

function resetDatas() { activeSorts = {}; activeFilters = {}; activeRules = []; document.getElementById('rulesList').innerHTML = ""; processDataRange(); updateFilterUI(); userAction(); }

// --- 9. APPLY TO ILLUSTRATOR (SYNC V103) ---
function applyToIllustrator() {
    userAction(); 
    if(loadedData.length === 0) { alert("Aucune donnée."); return; }

    var fullState = gatherFullState();
    var u = fullState.uiState;

    var act = []; 
    if(loadedData.length > 0) { for(var i=0; i<loadedData[0].length; i++) act.push(i); }

    var finalAligns = [];
    columnOrder.forEach(function(realIndex) {
        if (!hiddenOriginalIndices.includes(realIndex)) {
            var colName = headers[realIndex];
            finalAligns.push(globalColAligns[colName] || "left");
        }
    });

    var allRows = fullRawData.slice(1);

    // PRE-CALCUL MIN/MAX
    var colLimits = {};
    for (var colName in activeFilters) {
        var f = activeFilters[colName];
        if (f.op === 'min' || f.op === 'max') {
            var tIdx = headers.indexOf(colName);
            if (tIdx > -1) {
                var valList = allRows.map(r => {
                    return parseFloat(String(r[tIdx]||"").replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
                }).filter(n => !isNaN(n));
                if (valList.length > 0) {
                    if (f.op === 'min') colLimits[colName] = Math.min(...valList);
                    else colLimits[colName] = Math.max(...valList);
                }
            }
        }
    }
    
    // A. FILTRAGE GLOBAL
    var filteredRows = allRows.filter(function(row) {
        for (var colName in activeFilters) {
            var filter = activeFilters[colName];
            if (!filter || filter.op === 'none') continue;
            
            var targetIdx = headers.indexOf(colName);
            if (targetIdx === -1) continue;

            var cellVal = row[targetIdx] || "";
            var numCell = parseFloat(String(cellVal).replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
            var numA = parseFloat(String(filter.val).replace(',', '.'));
            
            var match = true;

            if (filter.op === 'not_empty') {
                if (String(cellVal).trim() === "") match = false;
            }
            else if (filter.op === 'min' || filter.op === 'max') {
                if (colLimits[colName] !== undefined) {
                    if (numCell !== colLimits[colName]) match = false;
                } else { match = false; }
            }
            else if (filter.op === '=') { if (String(cellVal).toLowerCase() !== String(filter.val).toLowerCase()) match = false; }
            else if (filter.op === 'contains') { if (String(cellVal).toLowerCase().indexOf(String(filter.val).toLowerCase()) === -1) match = false; }
            else if (['>', '>=', '<', '<=', 'between'].includes(filter.op)) {
                if (isNaN(numCell)) match = false;
                else {
                    if (filter.op === '>') { if (!(numCell > numA)) match = false; }
                    else if (filter.op === '>=') { if (!(numCell >= numA)) match = false; }
                    else if (filter.op === '<') { if (!(numCell < numA)) match = false; }
                    else if (filter.op === '<=') { if (!(numCell <= numA)) match = false; }
                    else if (filter.op === 'between') {
                        var numB = parseFloat(String(filter.valB).replace(',', '.'));
                        if (!(numCell >= Math.min(numA, numB) && numCell <= Math.max(numA, numB))) match = false;
                    }
                }
            }
            if (!match) return false;
        } 
        return true;
    });

    // B. PAGINATION
    var inputStart = document.getElementById('rowStart'); 
    var valNbLines = document.getElementById('valNbLines');
    var cbNbLines = document.getElementById('cbNbLines');
    var start = parseInt(inputStart.value) || 0;
    var nbLines = parseInt(valNbLines.value) || 0;
    
    var visibleRows = filteredRows;
    if (start > 0) visibleRows = visibleRows.slice(Math.max(0, start - 1));
    if (cbNbLines && cbNbLines.checked) { if (nbLines > 0) visibleRows = visibleRows.slice(0, nbLines); } 
    else {
        var inputEnd = document.getElementById('rowEnd'); var end = parseInt(inputEnd.value) || 0;
        if (end > 0 && end >= start) visibleRows = visibleRows.slice(0, (end - start) + 1);
    }

    // C. TRI LOCAL
    for (var colName in activeSorts) {
        var dir = activeSorts[colName];
        var c = headers.indexOf(colName);
        if (dir && c > -1) {
            visibleRows.sort(function(a, b) {
                var valA = a[c] || ""; var valB = b[c] || "";
                var numA = parseFloat(valA.replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, '')); 
                var numB = parseFloat(valB.replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
                if (!isNaN(numA) && !isNaN(numB)) { return (dir === 'asc') ? numA - numB : numB - numA; }
                return (dir === 'asc') ? valA.localeCompare(valB) : valB.localeCompare(valA);
            });
        }
    }

    // 2. STYLES
    var customStyles = { rows: {}, cells: {} }; 
    var finalHeaders = loadedData[0];

    for (var r = 0; r < visibleRows.length; r++) { 
        var fullRow = visibleRows[r]; 
        var outputRowIdx = r + 1; 

        activeRules.forEach(function(rule) {
            var condColIdx = headers.indexOf(rule.colName);
            if (condColIdx > -1) {
                var cellVal = fullRow[condColIdx];
                var numC = parseFloat(String(cellVal).replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, ''));
                var match = false;
                if (rule.op === 'contains') { if(String(cellVal).indexOf(rule.val) > -1) match=true; }
                else if (rule.op === 'between') { if (!isNaN(numC) && numC >= rule.valA && numC <= rule.valB) match=true; }
                else if (!isNaN(numC)) { var v = parseFloat(rule.val); if (rule.op==='=') match = (numC == v); if (rule.op==='>') match = (numC > v); if (rule.op==='>=') match = (numC >= v); if (rule.op==='<') match = (numC < v); if (rule.op==='<=') match = (numC <= v); } 
                else { if (rule.op==='=') match = (cellVal == rule.val); }
                
                if (match) {
                    if (rule.scope === 'row') customStyles.rows[outputRowIdx] = { bg: rule.bg, txt: rule.txt }; 
                    else { var finalVisualIndex = finalHeaders.indexOf(rule.colName); if (finalVisualIndex > -1) customStyles.cells[outputRowIdx + "_" + finalVisualIndex] = { bg: rule.bg, txt: rule.txt }; }
                }
            }
        });
    }

    var turboData = []; 

    var safePad = Math.max(0, parseFloat(u.pad)||0);
    var p = {
        data: loadedData, turboData: turboData,
        activeCols: act, colAligns: finalAligns, customStyles: customStyles, savedState: fullState, 
        csvOptions: { legend: u.legendText },
        geo: {
            fName: u.fName, fSize: parseFloat(u.fSize)||10, leading: parseFloat(u.fLeading)||12, isBold: u.isBold,
            hHead: parseFloat(u.hHead)||15, hRow: parseFloat(u.hRow)||10, pad: safePad, wrapHead: u.wrapHead, maxRows: 0, totalW: 0,
            radius: { tl: parseFloat(u.radius.tl)||0, tr: parseFloat(u.radius.tr)||0, bl: parseFloat(u.radius.bl)||0, br: parseFloat(u.radius.br)||0 },
            legRadius: { tl: parseFloat(u.legRadius.tl)||0, tr: parseFloat(u.legRadius.tr)||0, bl: parseFloat(u.legRadius.bl)||0, br: parseFloat(u.legRadius.br)||0 }
        },
        colors: { 
            bgHead: getCMYK('ui-cBgHead'), txtHead: getCMYK('ui-cTxtHead'), 
            txtCont: getCMYK('ui-cTxtCont'), 
            bgRow1: getCMYK('ui-cBgRow1'), bgRow2: getCMYK('ui-cBgRow2'), 
            bgLeg: getCMYK('ui-cBgLeg'), txtLeg: getCMYK('ui-cTxtLeg'), 
            strokeHead: getCMYK('ui-cStrokeHead') 
        },
        borders: { vert: document.getElementById('bVert') ? document.getElementById('bVert').checked : true, horz: document.getElementById('bHorz') ? document.getElementById('bHorz').checked : true, head: document.getElementById('bHead') ? document.getElementById('bHead').checked : true, outer: document.getElementById('bOuter') ? document.getElementById('bOuter').checked : true, weight: parseFloat(document.getElementById('bWeight').value)||0.5 }
    };
    
    // GESTION BOUTONS SLIDE
    if (!u.cbShowTitle) { p.geo.hHead = 0; p.borders.head = false; }
    if (!u.cbShowLegend) { p.csvOptions.legend = ""; } 

    runJSX('creerTableau("' + toHex(JSON.stringify(p)) + '")', function(){});
}

// --- 10. HELPERS ---
function gatherFullState() {
    var dataState = { fullRawData: fullRawData, headers: headers, columnOrder: columnOrder, hiddenOriginalIndices: hiddenOriginalIndices, activeRules: activeRules, activeSorts: activeSorts, activeFilters: activeFilters, globalColAligns: globalColAligns, selectedUiIndices: selectedUiIndices };
    
    var uiState = {
        rowStart: document.getElementById('rowStart').value, 
        rowEnd: document.getElementById('rowEnd').value, 
        valNbLines: document.getElementById('valNbLines').value, 
        cbNbLines: document.getElementById('cbNbLines').checked,
        pad: document.getElementById('pad').value, 
        legendText: document.getElementById('legendText').value,

        fName: (document.getElementById('ui-font-full') ? document.getElementById('ui-font-full').value : document.getElementById('fName').value),
        fSize: document.getElementById('fSize').value, 
        fLeading: document.getElementById('fLeading').value, 
        isBold: document.getElementById('isBold') ? document.getElementById('isBold').checked : false,
        
        hHead: document.getElementById('hHead').value, 
        hRow: document.getElementById('hRow').value, 
        wrapHead: document.getElementById('wrapHead') ? document.getElementById('wrapHead').checked : false, 
        
        cbShowTitle: document.getElementById('cbShowTitle') ? document.getElementById('cbShowTitle').checked : true,
        cbShowLegend: document.getElementById('cbShowLegend') ? document.getElementById('cbShowLegend').checked : true,
        
        radius: { tl: document.getElementById('radTL').value, tr: document.getElementById('radTR').value, bl: document.getElementById('radBL').value, br: document.getElementById('radBR').value },
        legRadius: { tl: document.getElementById('legRadTL').value, tr: document.getElementById('legRadTR').value, bl: document.getElementById('legRadBL').value, br: document.getElementById('legRadBR').value },
        borders: { vert: document.getElementById('bVert').checked, horz: document.getElementById('bHorz').checked, head: document.getElementById('bHead').checked, outer: document.getElementById('bOuter').checked, weight: document.getElementById('bWeight').value },
        colors: { bgHead: getCMYK('ui-cBgHead'), txtHead: getCMYK('ui-cTxtHead'), txtCont: getCMYK('ui-cTxtCont'), bgRow1: getCMYK('ui-cBgRow1'), bgRow2: getCMYK('ui-cBgRow2'), bgLeg: getCMYK('ui-cBgLeg'), txtLeg: getCMYK('ui-cTxtLeg'), strokeHead: getCMYK('ui-cStrokeHead') }
    };
    return { dataState: dataState, uiState: uiState };
}

function setupShiftIncrement(input) {
    if (!input || input.type !== 'number') return;
    if (input.dataset.shiftAttached === "true") return;
    input.dataset.shiftAttached = "true";
    input.addEventListener('keydown', function(e) {
        if (e.shiftKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
            e.preventDefault(); e.stopImmediatePropagation();
            var currentVal = parseFloat(this.value) || 0;
            var step = (e.key === "ArrowDown") ? -10 : 10;
            var newVal = currentVal + step;
            if (this.min !== "" && newVal < parseFloat(this.min)) newVal = parseFloat(this.min);
            if (this.max !== "" && newVal > parseFloat(this.max)) newVal = parseFloat(this.max);
            this.value = newVal; this.dispatchEvent(new Event('input'));
        }
    });
}
function runJSX(script, callback) { if (window.__adobe_cep__ && window.__adobe_cep__.evalScript) window.__adobe_cep__.evalScript(script, callback); else if (window.cep && window.cep.process && window.cep.process.evalScript) window.cep.process.evalScript(script, callback); }
function toHex(str) { try { var utf8 = unescape(encodeURIComponent(str)); var r = ''; for (var i = 0; i < utf8.length; i++) { var h = utf8.charCodeAt(i).toString(16); r += ("0" + h).slice(-2); } return r.toUpperCase(); } catch (e) { return ""; } }
function cmykToHex(c, m, y, k) { var r = 255 * (1 - c / 100) * (1 - k / 100); var g = 255 * (1 - m / 100) * (1 - k / 100); var b = 255 * (1 - y / 100) * (1 - k / 100); function compToHex(c) { var hex = Math.round(c).toString(16); return hex.length == 1 ? "0" + hex : hex; } return "#" + compToHex(r) + compToHex(g) + compToHex(b); }
function hexToCmyk(hex) { var r = parseInt(hex.substring(1, 3), 16) / 255; var g = parseInt(hex.substring(3, 5), 16) / 255; var b = parseInt(hex.substring(5, 7), 16) / 255; var k = 1 - Math.max(r, g, b); var c = (1 - r - k) / (1 - k) || 0; var m = (1 - g - k) / (1 - k) || 0; var y = (1 - b - k) / (1 - k) || 0; return [Math.round(c * 100), Math.round(m * 100), Math.round(y * 100), Math.round(k * 100)]; }
function getCMYK(id) { var d = document.getElementById(id); if (!d) return [0, 0, 0, 100]; var v = []; for (var i = 0; i < 4; i++) { var val = parseInt(d.querySelector(`.inp-${i}`).value); if (isNaN(val)) val = 0; v.push(val); } return v; }
function setCMYK(id, values) { var d = document.getElementById(id); if (!d) return; if (!values || !Array.isArray(values) || values.length < 4) values = [0, 0, 0, 100]; for (var i = 0; i < 4; i++) { var val = values[i] || 0; var sl = d.querySelector(`.sl-${i}`); var inp = d.querySelector(`.inp-${i}`); if(sl) sl.value = val; if(inp) inp.value = val; } try { var hex = cmykToHex(values[0] || 0, values[1] || 0, values[2] || 0, values[3] || 0); var pk = d.querySelector('.color-hidden'); var sw = d.querySelector('.swatch-visual'); if (pk) pk.value = hex; if (sw) sw.style.backgroundColor = hex; } catch(e) {} }
function createCMYKControl(id, t, v) { var d = document.getElementById(id); if (!d) return; var h = `<div class="cmyk-title">${t}</div><div class="cmyk-wrapper"><div class="cmyk-body"><div style="display:flex; flex-direction:column; gap:2px;">`; var l = ['C', 'M', 'J', 'N']; for (var i = 0; i < 4; i++) { h += `<div class="cmyk-row"><span class="cmyk-label" style="width:15px; font-size:10px; color:#aaa;">${l[i]}</span><input type="range" min="0" max="100" value="${v[i]}" class="sl-${i}"><div class="number-wrapper"><input type="number" min="0" max="100" value="${v[i]}" class="inp-${i}"><span class="unit-symbol">%</span></div></div>`; } h += `</div><div><div class="swatch-visual" style="width:25px; height:25px; border:1px solid #555; background-color:#000; cursor:pointer;" title="Changer"></div><input type="color" class="color-hidden" style="display:none;" value="#000000"></div></div></div>`; d.innerHTML = h; var inputs = d.querySelectorAll('input[type="number"]'); inputs.forEach(setupShiftIncrement); var swatch = d.querySelector('.swatch-visual'); var hiddenPicker = d.querySelector('.color-hidden'); swatch.onclick = function() { hiddenPicker.click(); }; function updatePreview() { var vals = getCMYK(id); var hex = cmykToHex(vals[0], vals[1], vals[2], vals[3]); hiddenPicker.value = hex; swatch.style.backgroundColor = hex; } function updateFromPicker() { var hex = hiddenPicker.value; swatch.style.backgroundColor = hex; setCMYK(id, hexToCmyk(hex)); userAction(); } hiddenPicker.oninput = function() { updateFromPicker(); userAction(); }; hiddenPicker.onchange = function() { updateFromPicker(); userAction(); }; var ranges = d.querySelectorAll('input[type="range"]'); ranges.forEach(function(rng) { rng.onmousedown = function() { userAction(); }; rng.oninput = function() { this.nextElementSibling.querySelector('input').value = this.value; updatePreview(); userAction(); }; }); inputs.forEach(function(num) { num.onchange = function() { var val = parseInt(this.value); if (isNaN(val)) val = 0; if (val > 100) val = 100; if (val < 0) val = 0; this.value = val; this.parentElement.previousElementSibling.value = val; updatePreview(); userAction(); }; num.oninput = function() { if (this.value === "") { updatePreview(); return; } var val = parseInt(this.value); if (val > 100) { val = 100; this.value = 100; } this.parentElement.previousElementSibling.value = val; updatePreview(); userAction(); }; }); updatePreview(); }
function initAllColors() { createCMYKControl('ui-cBgHead', 'Fond', [0, 0, 0, 80]); createCMYKControl('ui-cTxtHead', 'Texte', [0, 0, 0, 0]); createCMYKControl('ui-cBgRow1', 'Fond ligne 1', [0, 0, 0, 0]); createCMYKControl('ui-cBgRow2', 'Fond ligne 2', [0, 0, 0, 10]); createCMYKControl('ui-cTxtCont', 'Texte', [0, 0, 0, 100]); createCMYKControl('ui-cBgLeg', 'Fond Légende', [0, 0, 0, 0]); createCMYKControl('ui-cTxtLeg', 'Texte Légende', [0, 0, 0, 100]); createCMYKControl('ui-cStrokeHead', 'Filets Titre', [0, 0, 0, 0]); createCMYKControl('ui-ruleBg', 'Fond Conditionnel', [0, 50, 0, 0]); createCMYKControl('ui-ruleTxt', 'Texte Conditionnel', [0, 0, 0, 100]); }
function initInputs() { document.querySelectorAll('.container input[type="number"]').forEach(setupShiftIncrement); var cbNb = document.getElementById('cbNbLines'); if (cbNb) setTimeout(function() { cbNb.dispatchEvent(new Event('change')); }, 50); var bAll = document.getElementById('bAll'); if (bAll) { bAll.onchange = function() { var st = this.checked; ['bVert', 'bHorz', 'bHead', 'bOuter'].forEach(id => { var el = document.getElementById(id); if (el) el.checked = st; }); userAction(); }; } ['bVert', 'bHorz', 'bHead', 'bOuter'].forEach(id => { var el = document.getElementById(id); if (el) el.addEventListener('change', function() { userAction(); }); }); var radInputs = ['radTL', 'radTR', 'radBL', 'radBR', 'radGlobal', 'legRadTL', 'legRadTR', 'legRadBL', 'legRadBR', 'legRadGlobal']; radInputs.forEach(id => { var el = document.getElementById(id); if (el) el.addEventListener('input', function() { userAction(); }); }); var inputStart = document.getElementById('rowStart'); var inputEnd = document.getElementById('rowEnd'); var valNbLines = document.getElementById('valNbLines'); var cbNbLines = document.getElementById('cbNbLines'); function onRangeChange() { userAction(); processDataRange(); } if (cbNbLines) { cbNbLines.onchange = function() { var w1 = document.getElementById('wrapRowEnd'); var w2 = document.getElementById('wrapNbLines'); if (this.checked) { w1.classList.add('input-disabled'); w2.classList.remove('input-disabled'); inputEnd.disabled = true; valNbLines.disabled = false; } else { w1.classList.remove('input-disabled'); w2.classList.add('input-disabled'); inputEnd.disabled = false; valNbLines.disabled = true; } onRangeChange(); }; } if (inputStart) inputStart.oninput = onRangeChange; if (inputEnd) inputEnd.oninput = onRangeChange; if (valNbLines) valNbLines.oninput = onRangeChange; var cbTitle = document.getElementById('cbShowTitle'); var inputHead = document.getElementById('hHead'); if (cbTitle && inputHead) { inputHead.addEventListener('input', function() { var v = parseFloat(this.value); if (v > 0) savedHeaderHeight = v; userAction(); }); cbTitle.addEventListener('change', function() { if (this.checked) { if (inputHead.value == 0 || inputHead.value == "0") inputHead.value = savedHeaderHeight; } else { var current = parseFloat(inputHead.value); if (current > 0) savedHeaderHeight = current; } userAction(); }); } }
function toggleSection(id, btn) { var el = document.getElementById(id); if (el) { if (el.classList.contains('collapsed')) { el.classList.remove('collapsed'); btn.classList.remove('collapsed'); } else { el.classList.add('collapsed'); btn.classList.add('collapsed'); } } }
function initTabs() { var mainTabs = document.querySelectorAll('.tab-btn'); mainTabs.forEach(function(btn) { btn.onclick = function() { var group = this.getAttribute('data-group'); document.querySelectorAll(`.tab-btn[data-group="${group}"]`).forEach(b => b.classList.remove('active')); this.classList.add('active'); document.querySelectorAll(`.tab-content[data-group="${group}"]`).forEach(c => c.classList.remove('active')); document.getElementById(this.getAttribute('data-tab')).classList.add('active'); }; }); var subTabs = document.querySelectorAll('.sub-tab-btn'); subTabs.forEach(function(btn) { btn.onclick = function() { var p = this.parentElement; p.querySelectorAll('.sub-tab-btn').forEach(b => b.classList.remove('active')); this.classList.add('active'); var tid = this.getAttribute('data-subtab'); var tc = document.getElementById(tid); if (tc) { Array.from(tc.parentElement.children).forEach(c => { if (c.classList.contains('sub-tab-content')) c.classList.remove('active'); }); tc.classList.add('active'); } }; }); }
function loadFonts() { var originalSelect = document.getElementById('fName'); var container = document.getElementById('font-ui-container'); if (!container && originalSelect) { originalSelect.style.display = 'none'; container = document.createElement('div'); container.id = 'font-ui-container'; container.style.display = 'flex'; container.style.flexDirection = 'column'; container.style.gap = '5px'; var selFull = document.createElement('select'); selFull.id = 'ui-font-full'; selFull.style.width = '240px'; var filterInput = document.getElementById("fontFilter"); if (filterInput) { filterInput.oninput = function() { var q = this.value.toLowerCase(); for (var i = 0; i < selFull.options.length; i++) { var opt = selFull.options[i]; opt.style.display = opt.text.toLowerCase().includes(q) ? "block" : "none"; } }; } originalSelect.parentNode.insertBefore(container, originalSelect); container.appendChild(selFull); selFull.onchange = function() { if (isScriptUpdating) return; userAction(); originalSelect.value = this.value; }; } var selFull = document.getElementById('ui-font-full'); runJSX('getSystemFonts()', function(res) { if (!res) return; var fonts = res.split("|").filter(f => f && f.trim().length > 0).sort(); if (selFull) selFull.innerHTML = ""; if (originalSelect) originalSelect.innerHTML = ""; fonts.forEach(function(f) { var o = document.createElement('option'); o.value = f; o.innerText = f; if (selFull) selFull.appendChild(o); var oh = document.createElement('option'); oh.value = f; oh.innerText = f; if (originalSelect) originalSelect.appendChild(oh); }); if (fonts.length > 0) { if (selFull) selFull.value = fonts[0]; if (originalSelect) originalSelect.value = fonts[0]; } }); }
function saveSettings() { if (isInit || isScriptUpdating) return; var s = { activeRules: activeRules, activeSorts: activeSorts, activeFilters: activeFilters, globalColAligns: globalColAligns, hiddenOriginalIndices: hiddenOriginalIndices, rowStart: document.getElementById('rowStart').value, rowEnd: document.getElementById('rowEnd').value, valNbLines: document.getElementById('valNbLines').value, cbNbLines: document.getElementById('cbNbLines').checked, pad: document.getElementById('pad').value, fSize: document.getElementById('fSize').value, colors: { bgHead: getCMYK('ui-cBgHead'), txtHead: getCMYK('ui-cTxtHead'), bgRow1: getCMYK('ui-cBgRow1'), bgRow2: getCMYK('ui-cBgRow2') } }; localStorage.setItem("TableauPro_Settings", JSON.stringify(s)); }
function loadSettings() { var raw = localStorage.getItem("TableauPro_Settings"); if (!raw) return; try { var s = JSON.parse(raw); if (s.activeRules) activeRules = s.activeRules; if (s.activeSorts) activeSorts = s.activeSorts; if (s.activeFilters) activeFilters = s.activeFilters; if (s.globalColAligns) globalColAligns = s.globalColAligns; if (s.hiddenOriginalIndices) hiddenOriginalIndices = s.hiddenOriginalIndices; else if (s.uncheckedColIndices) hiddenOriginalIndices = s.uncheckedColIndices; if (s.rowStart) document.getElementById('rowStart').value = s.rowStart; if (s.valNbLines) document.getElementById('valNbLines').value = s.valNbLines; if (s.pad) document.getElementById('pad').value = s.pad; if (s.fSize) document.getElementById('fSize').value = s.fSize; if (s.colors) { if (s.colors.bgHead) setCMYK('ui-cBgHead', s.colors.bgHead); if (s.colors.txtHead) setCMYK('ui-cTxtHead', s.colors.txtHead); if (s.colors.bgRow1) setCMYK('ui-cBgRow1', s.colors.bgRow1); if (s.colors.bgRow2) setCMYK('ui-cBgRow2', s.colors.bgRow2); } renderRules(); } catch (e) {} }

// --- 11. WATCHER ---
function startSelectionWatcher() { try { var csInterface = new CSInterface(); csInterface.addEventListener("documentAfterSelectionChanged", function() { clearTimeout(updateTimeout); updateTimeout = setTimeout(checkForUpdates, 50); }); } catch (e) {} setInterval(checkForUpdates, 500); }
function checkForUpdates() { if (isFetching || isUserInteracting) return; isFetching = true; runJSX('getSelectionData()', function(res) { isFetching = false; if (!res || res === "" || res === lastSelectionData) { if (res === "") lastSelectionData = ""; return; } lastSelectionData = res; updateUIFromData(res); }); }
function updateUIFromData(hexData) { try { isScriptUpdating = true; var hexArr = []; for (var i = 0; i < hexData.length; i += 2) { hexArr.push(String.fromCharCode(parseInt(hexData.substr(i, 2), 16))); } var p = JSON.parse(decodeURIComponent(escape(hexArr.join('')))); var d, u; if (p.savedState) { d = p.savedState.dataState; u = p.savedState.uiState; fullRawData = d.fullRawData || []; headers = d.headers || []; columnOrder = d.columnOrder || []; hiddenOriginalIndices = d.hiddenOriginalIndices || []; activeRules = d.activeRules || []; activeSorts = d.activeSorts || {}; activeFilters = d.activeFilters || {}; globalColAligns = d.globalColAligns || {}; } else { d = {}; u = { geo: p.geo || {}, colors: p.colors || {}, borders: p.borders || {}, csvOptions: p.csvOptions || {} }; if (p.data) loadedData = p.data; } document.getElementById('fileNameDisplay').innerText = "Données Illustrator récupérées"; document.getElementById('btnApply').disabled = false; if(u.rowStart !== undefined) setVal('rowStart', u.rowStart); if(u.valNbLines !== undefined) setVal('valNbLines', u.valNbLines); if(u.cbNbLines !== undefined) setCheck('cbNbLines', u.cbNbLines); var srcGeo = u.geo || u; setVal('pad', (u.pad !== undefined) ? u.pad : srcGeo.pad); setVal('fSize', (u.fSize !== undefined) ? u.fSize : srcGeo.fSize); setVal('fLeading', (u.fLeading !== undefined) ? u.fLeading : srcGeo.leading); if (srcGeo.hHead !== undefined) setVal('hHead', srcGeo.hHead); if (srcGeo.hRow !== undefined) setVal('hRow', srcGeo.hRow); var hHeadVal = parseFloat(document.getElementById('hHead').value); setCheck('cbShowTitle', hHeadVal > 0); var cols = u.colors || p.colors || {}; setCMYK('ui-cBgHead', cols.bgHead); setCMYK('ui-cTxtHead', cols.txtHead); setCMYK('ui-cTxtCont', cols.txtCont); setCMYK('ui-cBgRow1', cols.bgRow1); setCMYK('ui-cBgRow2', cols.bgRow2); setCMYK('ui-cBgLeg', cols.bgLeg); setCMYK('ui-cTxtLeg', cols.txtLeg); setCMYK('ui-cStrokeHead', cols.strokeHead); var b = u.borders || p.borders || {}; setCheck('bVert', b.vert); setCheck('bHorz', b.horz); setCheck('bHead', b.head); setCheck('bOuter', b.outer); if (b.weight !== undefined) setVal('bWeight', b.weight); 
// NOUVEAU FIX UPDATE : On respecte l'état explicite s'il existe
if (u.cbShowTitle !== undefined) setCheck('cbShowTitle', u.cbShowTitle); 
if (u.cbShowLegend !== undefined) { 
    setCheck('cbShowLegend', u.cbShowLegend); 
    // On met à jour l'état visuel (disabled) en fonction du texte récupéré
    var txt = document.getElementById('legendText');
    var cb = document.getElementById('cbShowLegend');
    if(txt && cb) cb.disabled = (txt.value.trim() === "");
}
if (p.savedState) { generateColUI(false); updateFilterUI(); renderRules(); } isScriptUpdating = false; } catch (e) { console.error("Erreur UpdateUI:", e); isScriptUpdating = false; } }
function setVal(id, v) { var el = document.getElementById(id); if(el && v !== undefined) el.value = v; }
function setCheck(id, v) { var el = document.getElementById(id); if(el && v !== undefined) { el.checked = v; el.dispatchEvent(new Event('change')); } }