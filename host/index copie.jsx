/* host/index.jsx - VERSION V68 : SUPPORTS FULL STATE SAVE/RESTORE */

$.global.creerTableau = function(hexData) {
    illustratorEngine(hexData);
    return "Fait";
};

// --- FONCTION DE RECUPERATION DES DONNEES ---
$.global.getSelectionData = function() {
    if (app.documents.length === 0) return "";
    var doc = app.activeDocument;
    
    // On ne traite que si un seul objet (le groupe tableau) est sélectionné
    if (doc.selection.length !== 1) return ""; 
    var item = doc.selection[0];
    
    try {
        // On cherche notre étiquette spéciale contenant tout l'état (Data + UI)
        return item.tags.getByName("TableauPro_Data").value;
    } catch(e) {
        // Pas de tag = ce n'est pas un de nos tableaux
        return "";
    }
};

$.global.getSystemFonts = function() {
    $.gc(); 
    var arr = [];
    var fonts = app.textFonts;
    for (var i = 0; i < fonts.length; i++) {
        try { arr.push(fonts[i].name); } catch(e) {}
    }
    return arr.join("|");
};

var illustratorEngine = function(hexData) {
    $.gc(); 

    function fromHex(hex) {
        var str = "";
        for (var i = 0; i < hex.length; i += 2) str += String.fromCharCode(parseInt(hex.substr(i, 2), 16));
        try { return decodeURIComponent(escape(str)); } catch(e) { return str; }
    }

    function makeCMYK(arr){
        var c=new CMYKColor(); c.cyan=arr[0]; c.magenta=arr[1]; c.yellow=arr[2]; c.black=arr[3]; return c;
    }

    function getFontByName(name) {
        try { return app.textFonts.getByName(name); } catch(e) {
            for (var i = 0; i < app.textFonts.length; i++) { if (app.textFonts[i].name == name) return app.textFonts[i]; }
            if (app.textFonts.length > 0) return app.textFonts[0];
            return null;
        }
    }

    // --- GEOMETRIE ---
    function drawRoundedPath(container, x, y, w, h, r) {
        var p = container.pathItems.add();
        var left = x; var right = x + w; var top = y; var bottom = y - h;
        var maxR = Math.min(w/2, h/2);
        var rtl = Math.min(r.tl, maxR); var rtr = Math.min(r.tr, maxR);
        var rbr = Math.min(r.br, maxR); var rbl = Math.min(r.bl, maxR);
        var cornerTL = [left, top]; var cornerTR = [right, top];
        var cornerBR = [right, bottom]; var cornerBL = [left, bottom];

        function addPt(ax, ay, ix, iy, ox, oy) {
            var pp = p.pathPoints.add();
            pp.anchor = [ax, ay]; pp.leftDirection = [ix, iy]; pp.rightDirection = [ox, oy];
            pp.pointType = PointType.SMOOTH;
        }

        addPt(left + rtl, top, (rtl > 0) ? cornerTL[0] : left, (rtl > 0) ? cornerTL[1] : top, left + rtl, top);
        addPt(right - rtr, top, right - rtr, top, (rtr > 0) ? cornerTR[0] : right, (rtr > 0) ? cornerTR[1] : top);
        addPt(right, top - rtr, (rtr > 0) ? cornerTR[0] : right, (rtr > 0) ? cornerTR[1] : top, right, top - rtr);
        addPt(right, bottom + rbr, right, bottom + rbr, (rbr > 0) ? cornerBR[0] : right, (rbr > 0) ? cornerBR[1] : bottom);
        addPt(right - rbr, bottom, (rbr > 0) ? cornerBR[0] : right, (rbr > 0) ? cornerBR[1] : bottom, right - rbr, bottom);
        addPt(left + rbl, bottom, left + rbl, bottom, (rbl > 0) ? cornerBL[0] : left, (rbl > 0) ? cornerBL[1] : bottom);
        addPt(left, bottom + rbl, (rbl > 0) ? cornerBL[0] : left, (rbl > 0) ? cornerBL[1] : bottom, left, bottom + rbl);
        addPt(left, top - rtl, left, top - rtl, (rtl > 0) ? cornerTL[0] : left, (rtl > 0) ? cornerTL[1] : top);

        p.closed = true;
        return p;
    }

    // --- NETTOYAGE TEXTE ---
    function cleanText(str) {
        if (!str) return "";
        var s = String(str);
        s = s.replace(/^"|"$/g,'');
        var CR = String.fromCharCode(13);
        s = s.replace(/\\n/g, CR).replace(/\n/g, CR).replace(/\r/g, CR);
        s = s.replace(/[\u00A0\u202F\u2007\u2060]/g, " "); 
        s = s.replace(/^\s+|\s+$/g, ''); 
        return s;
    }

    try {
        var params = eval('(' + fromHex(hexData) + ')'); 
        if (app.documents.length === 0) return;
        var doc = app.documents[0]; 

        var layer; try{layer=doc.layers.getByName("Calque_Tableau");}catch(e){layer=doc.layers.add();layer.name="Calque_Tableau";}
        layer.locked=false; layer.visible=true;

        var startX = 0; var startY = 0; var isUpdate = false;
        
        // --- DETECTION DE MISE A JOUR ---
        // On vérifie si l'utilisateur a sélectionné un ancien tableau pour le remplacer
        if (doc.selection.length === 1) {
            var sel = doc.selection[0];
            try {
                // On cherche le tag "TableauPro_Data" pour confirmer que c'est bien notre tableau
                // (On pourrait aussi chercher "TableData" pour la compatibilité avec tes versions précédentes)
                sel.tags.getByName("TableauPro_Data");
                
                // Si trouvé, on garde sa position et on le supprime pour redessiner le nouveau à la place
                startX = sel.left; startY = sel.top;
                sel.remove(); isUpdate = true;
            } catch(e) {
                // Si pas de tag Pro, on check l'ancien tag par sécurité
                try {
                    sel.tags.getByName("TableData");
                    startX = sel.left; startY = sel.top;
                    sel.remove(); isUpdate = true;
                } catch(ex) {}
            }
        }

        var data = params.data;
        var cols = params.activeCols; 
        var aligns = params.colAligns;
        
        var customStyles = params.customStyles || { rows:{}, cells:{} }; 

        if (!cols || cols.length === 0) return;

        var rowLimit = data.length;
        if (params.geo.maxRows > 0 && (params.geo.maxRows+1) < data.length) rowLimit = params.geo.maxRows + 1;

        var cHeadT=makeCMYK(params.colors.txtHead); var cHeadB=makeCMYK(params.colors.bgHead);
        var cContT=makeCMYK(params.colors.txtCont); var cRow1=makeCMYK(params.colors.bgRow1);
        var cRow2=makeCMYK(params.colors.bgRow2); 
        var cBlack=new CMYKColor(); cBlack.cyan=0; cBlack.magenta=0; cBlack.yellow=0; cBlack.black=100;
        
        var cBgLeg = (params.colors.bgLeg) ? makeCMYK(params.colors.bgLeg) : makeCMYK([0,0,0,0]);
        var cTxtLeg = (params.colors.txtLeg) ? makeCMYK(params.colors.txtLeg) : cBlack;
        var cStrokeHead = (params.colors.strokeHead) ? makeCMYK(params.colors.strokeHead) : makeCMYK([0,0,0,0]);

        var geo = params.geo;
        if (geo.pad <= 0) geo.pad = 0.1;

        var borders = params.borders; 
        var strokeW = (typeof borders.weight === 'number') ? borders.weight : 0.5;
        var targetFont = getFontByName(geo.fName);

        var rads = { tl:0, tr:0, bl:0, br:0 };
        if (geo.radius) {
            rads.tl = geo.radius.tl * 2.834645; rads.tr = geo.radius.tr * 2.834645;
            rads.bl = geo.radius.bl * 2.834645; rads.br = geo.radius.br * 2.834645;
        }
        var lRads = { tl:0, tr:0, bl:0, br:0 };
        if (geo.legRadius) {
            lRads.tl = geo.legRadius.tl * 2.834645; lRads.tr = geo.legRadius.tr * 2.834645;
            lRads.bl = geo.legRadius.bl * 2.834645; lRads.br = geo.legRadius.br * 2.834645;
        }

        var tempText = layer.textFrames.add();
        var tRange = tempText.textRange;
        try { if(targetFont) tRange.characterAttributes.textFont = targetFont; } catch(e){}
        tRange.characterAttributes.size = geo.fSize;

        var finalColWidths = []; 
        var colDataWidths = []; 
        var totalW = 0;
        var scanLimit = rowLimit; 
        var startRowCalc = 1; if (data.length <= 1) startRowCalc = 0; 

        // --- CALCUL LARGEURS ---
        for (var k=0; k<cols.length; k++) {
            var colIdx = cols[k]; 
            var maxW = 0;     
            var maxDataW = 0; 

            // A. Données
            for (var r=startRowCalc; r<scanLimit; r++) {
                var txt = (data[r] && data[r][colIdx]) ? cleanText(data[r][colIdx]) : "";
                if(txt!==""){
                    try { if(targetFont) tempText.textRange.characterAttributes.textFont = targetFont; } catch(e){}
                    tempText.textRange.characterAttributes.size = geo.fSize;
                    tempText.textRange.characterAttributes.strokeWeight = 0;
                    tempText.contents = txt;
                    
                    var thisW = tempText.width;
                    if (thisW > maxW) maxW = thisW;
                    if (thisW > maxDataW) maxDataW = thisW; 
                }
            }
            
            // B. Header
            if (geo.wrapHead && data[0] && data[0][colIdx]) {
                var hTxt = cleanText(data[0][colIdx]);
                try { if(targetFont) tempText.textRange.characterAttributes.textFont = targetFont; } catch(e){}
                tempText.textRange.characterAttributes.size = geo.fSize;
                
                if(geo.isBold) { try{ tempText.textRange.characterAttributes.strokeWeight = 0.25; }catch(e){} } 
                else { try{ tempText.textRange.characterAttributes.strokeWeight = 0; }catch(e){} }
                
                var hWords = hTxt.split(" ");
                var calculatedMaxW = 0;
                for(var z=0; z<hWords.length; z++) {
                    var word = hWords[z];
                    if (z > 0 && word.length <= 3) {
                        var pair = hWords[z-1] + " " + word;
                        tempText.contents = pair;
                        var wPair = tempText.width;
                        if (wPair > calculatedMaxW) calculatedMaxW = wPair;
                    } else {
                        tempText.contents = word;
                        var wWord = tempText.width;
                        if (wWord > calculatedMaxW) calculatedMaxW = wWord;
                    }
                }
                var safeHeaderW = calculatedMaxW * 1.10; 
                if (maxW < safeHeaderW) maxW = safeHeaderW;
            }

            var cw = Math.max(maxW + (geo.pad*2), 2);
            finalColWidths.push(cw);
            colDataWidths.push(maxDataW); 
            totalW += cw;
        }
        
        // --- WRAP FUNCTION ---
        function applyWordWrap(textString, maxWidth, isBoldText) {
            if (!textString || textString.indexOf(String.fromCharCode(13)) > -1) return textString; 
            try { if(targetFont) tempText.textRange.characterAttributes.textFont = targetFont; } catch(e){}
            tempText.textRange.characterAttributes.size = geo.fSize;
            if (isBoldText) { try { tempText.textRange.characterAttributes.strokeWeight = 0.25; } catch(e) {} } 
            else { try { tempText.textRange.characterAttributes.strokeWeight = 0; } catch(e) {} }
            var words = textString.split(" ");
            var resultLines = [];
            var currentLine = words[0];
            for (var i = 1; i < words.length; i++) {
                var word = words[i];
                var testLine = currentLine + " " + word;
                tempText.contents = testLine; 
                if (tempText.width > maxWidth) { resultLines.push(currentLine); currentLine = word; } 
                else { currentLine = testLine; }
            }
            resultLines.push(currentLine);
            return resultLines.join(String.fromCharCode(13));
        }

        if (!isUpdate) {
            try { 
                var abRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
                startX = abRect[0] + ((abRect[2]-abRect[0])/2) - (totalW/2); 
                startY = abRect[1] - 50;
            } catch(e) { }
        }

        var grp = layer.groupItems.add();
        grp.name = "Tableau_GENERE";
        
        // --- SAUVEGARDE DES DONNEES (TAG) ---
        // C'est ici qu'on stocke tout le JSON (savedState inclus) dans le groupe
        var tag = grp.tags.add(); 
        tag.name = "TableauPro_Data"; // Nom officiel pour la V68+
        tag.value = hexData;

        var curY = startY;
        var grpTable = grp.groupItems.add();
        grpTable.name = "Contenu_Tableau";

        // --- DESSIN TABLEAU ---
        for (var r=0; r<rowLimit; r++) {
            var isH = (r===0);
            if (isH && geo.hHead === 0) { continue; }

            var lH = isH ? geo.hHead : geo.hRow;
            // 1. Couleur de base de la ligne
            var baseRowBg = isH ? cHeadB : (r%2!==0 ? cRow1 : cRow2);
            var baseRowTxt = isH ? cHeadT : cContT;

            // 2. Surcharge LIGNE
            if (customStyles.rows && customStyles.rows[r]) {
                if(customStyles.rows[r].bg) baseRowBg = makeCMYK(customStyles.rows[r].bg);
                if(customStyles.rows[r].txt) baseRowTxt = makeCMYK(customStyles.rows[r].txt);
            }
            
            var curX = startX;

            for (var k=0; k<cols.length; k++) {
                var colIdx = cols[k]; var w = finalColWidths[k];
                var refDataW = colDataWidths[k]; 
                var txt = (data[r] && data[r][colIdx]) ? cleanText(data[r][colIdx]) : "";
                var align = aligns[k] || "left"; 

                // 3. Init Cellule
                var cBg = baseRowBg;
                var cTxt = baseRowTxt;

                // 4. Surcharge CELLULE
                var cellKey = r + "_" + k; 
                if (customStyles.cells && customStyles.cells[cellKey]) {
                    if (customStyles.cells[cellKey].bg) cBg = makeCMYK(customStyles.cells[cellKey].bg);
                    if (customStyles.cells[cellKey].txt) cTxt = makeCMYK(customStyles.cells[cellKey].txt);
                }

                var rect = grpTable.pathItems.rectangle(curY, curX, w, lH);
                rect.stroked = false; rect.filled = true; rect.fillColor = cBg;

                if (txt !== "") {
                    try {
                        if (isH && geo.wrapHead) {
                            var availW = w - (geo.pad * 2);
                            try { if(targetFont) tempText.textRange.characterAttributes.textFont = targetFont; } catch(e){}
                            tempText.textRange.characterAttributes.size = geo.fSize;
                            if (geo.isBold) availW = availW * 0.95; 
                            txt = applyWordWrap(txt, availW, geo.isBold);
                        }

                        var tf = grpTable.textFrames.add(); 
                        tf.contents = txt;
                        var range = tf.textRange;
                        var ca = range.characterAttributes;
                        var pa = range.paragraphAttributes;

                        // --- JUSTIFICATION ---
                        if (isH) {
                            pa.justification = Justification.CENTER;
                        } else {
                            if (align === "center") pa.justification = Justification.CENTER;
                            else if (align === "right") pa.justification = Justification.RIGHT;
                            else pa.justification = Justification.LEFT;
                        }

                        if (targetFont) { try { ca.textFont = targetFont; } catch(e){} }
                        ca.size = geo.fSize; ca.fillColor = cTxt; 

                        if (isH) {
                            if (geo.isBold) { try { ca.strokeWeight=0.25; ca.strokeColor=cTxt; } catch(e){} }
                            if (geo.leading > 0) { ca.autoLeading = false; ca.leading = geo.leading; }
                        }
                        
                        var doScale = true; 
                        var availW = w - geo.pad;
                        if (doScale && availW > 5 && tf.width > availW) {
                            var sc = (availW / tf.width) * 100;
                            if (sc < 10) sc = 10; 
                            if (sc > 100) sc = 100;
                            ca.horizontalScale = sc;
                        }

                        var tW = tf.width; var tH = tf.height;
                        
                        // --- POSITIONNEMENT ---
                        var pX = curX + geo.pad;
                        var colCenter = curX + (w/2);
                        
                        if (isH) {
                            pX = colCenter - (tW/2);
                        } 
                        else {
                            var boxLeft = colCenter - (refDataW / 2);
                            var boxRight = colCenter + (refDataW / 2);

                            if(align === "right") { pX = boxRight - tW; } 
                            else if(align === "left") { pX = boxLeft; }
                            else { pX = colCenter - (tW/2); }
                        }

                        var pY = curY - (lH/2) + (tH/2);
                        pY -= (geo.fSize * 0.11); 

                        tf.left = pX; tf.top = pY;
                        
                    } catch(et){}
                }
                curX += w;
            }
            curY -= lH;
        }

        var totalH = Math.abs(curY - startY);
        var rawTableH = totalH;

        if (rads.tl > 0 || rads.tr > 0 || rads.bl > 0 || rads.br > 0) {
            var clipPath = drawRoundedPath(grpTable, startX, startY, totalW, rawTableH, rads);
            clipPath.stroked = false; clipPath.filled = false;
            grpTable.clipped = true;
        }

        // --- LÉGENDE ---
        if (params.csvOptions && params.csvOptions.legend && params.csvOptions.legend !== "") {
            var rawLegend = cleanText(params.csvOptions.legend); 
            var mm = 2.834645;
            var gapLegend = 0.8 * mm; 
            var padLegend = 0.6 * mm; 
            var tableBottomY = startY - rawTableH;
            var bgTop = tableBottomY - gapLegend;

            try {
                tempText.contents = ""; 
                tempText.textRange.characterAttributes.size = 7;
                tempText.textRange.characterAttributes.strokeWeight = 0;
                try { tempText.textRange.characterAttributes.textFont = app.textFonts.getByName("Arial-Italic"); } catch(e){}

                var availW = totalW - (padLegend * 2);
                if (availW < 10) availW = 10; 
                
                var lWords = rawLegend.split(" ");
                var lLines = [];
                var lCurLine = lWords[0];
                for(var i=1; i<lWords.length; i++) {
                    var lWord = lWords[i];
                    var lTest = lCurLine + " " + lWord;
                    tempText.contents = lTest;
                    if(tempText.width > availW) {
                        lLines.push(lCurLine);
                        lCurLine = lWord;
                    } else {
                        lCurLine = lTest;
                    }
                }
                lLines.push(lCurLine);
                var wrappedLegend = lLines.join(String.fromCharCode(13));

                var legTf = grp.textFrames.add();
                legTf.contents = wrappedLegend;
                var legRange = legTf.textRange;
                var legCa = legRange.characterAttributes;
                
                var italicFont = null;
                try { italicFont = app.textFonts.getByName("Arial-Italic"); } catch(e){}
                if (!italicFont) { try { italicFont = app.textFonts.getByName("Helvetica-Oblique"); } catch(e){} }
                
                if (italicFont) { legCa.textFont = italicFont; } 
                else { try { if(targetFont) legCa.textFont = targetFont; } catch(e){} }

                legCa.size = 7; 
                legCa.autoLeading = false; 
                legCa.leading = 7; 
                legCa.fillColor = cTxtLeg;

                legTf.top = bgTop - padLegend; legTf.left = startX + padLegend;
                
                var bgH = legTf.height + (padLegend * 2);

                var legBg = drawRoundedPath(grp, startX, bgTop, totalW, bgH, lRads);
                legBg.stroked = false; legBg.filled = true; legBg.fillColor = cBgLeg; 
                legBg.move(legTf, ElementPlacement.PLACEAFTER); 
            } catch(e_leg) {}
        }

        tempText.remove(); 

        function drawLine(x1, y1, x2, y2, container, color) {
            var l = container.pathItems.add(); l.setEntirePath([[x1, y1], [x2, y2]]);
            l.stroked = true; l.strokeColor = (color) ? color : cBlack; l.strokeWidth = strokeW; l.filled = false;
        }
        
        var yHeaderBottom = startY - geo.hHead;

        if (borders.vert) {
            var vx = startX;
            for(var k=0; k<cols.length-1; k++) { 
                vx += finalColWidths[k]; 
                if (geo.hHead > 0) { drawLine(vx, startY, vx, yHeaderBottom, grpTable, cStrokeHead); }
                var startBody = (geo.hHead > 0) ? yHeaderBottom : startY;
                drawLine(vx, startBody, vx, startY - rawTableH, grpTable, cBlack); 
            }
        }
        if (borders.horz) {
            var startLine = (geo.hHead === 0) ? 0 : 1; 
            var hy = startY - geo.hHead;
            for(var r=startLine; r<rowLimit-1; r++) { hy -= geo.hRow; drawLine(startX, hy, startX+totalW, hy, grpTable, cBlack); }
        }
        if (borders.head && geo.hHead > 0) { drawLine(startX, yHeaderBottom, startX+totalW, yHeaderBottom, grpTable, cStrokeHead); }
        
        if (borders.outer) {
            var out = drawRoundedPath(grp, startX, startY, totalW, rawTableH, rads);
            out.filled = false; out.stroked = true; out.strokeColor = cBlack; out.strokeWidth = strokeW;
            out.zOrder(ZOrderMethod.BRINGTOFRONT);
        }
        
        grp.selected = true;

    } catch(err) { alert("Erreur JSX: " + err.message); }
};