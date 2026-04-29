function processFile() {
  const fileInput = document.getElementById("upload");
  const file = fileInput.files[0];
  if (!file) { alert("اختار ملف الأول"); return; }

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      
      const rawRows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        blankrows: false,
      });

      if (rawRows.length === 0) { alert("الملف فاضي!"); return; }

      // ────────────────────────────────────────────────────────────
      // 1) Helpers – أدوات التطبيع والتنظيف
      // ────────────────────────────────────────────────────────────
      
      /** يحول النص لحالة موحّدة ويشيل الرموز والتشكيل */
      function normalizeText(str) {
        return String(str)
          .trim()
          .toLowerCase()
          .replace(/\s+/g, " ")
          .replace(/[\u064B-\u065F\u0670]/g, "")       // تشكيل عربي
          .replace(/[أإآٱ]/g, "ا")                     // ألفات
          .replace(/[ة]/g, "ه")                         // تاء مربوطة
          .replace(/[يى]̨/g, "ي")                       // ياء
          .replace(/[ؤئ]/g, "و")                        // واو مهموزة
          .replace(/[٠-٩]/g, d => "0123456789"["٠١٢٣٤٥٦٧٨٩".indexOf(d)]) // أرقام عربية > لاتينية
          .replace(/٫/g, ".")                           // فاصلة عشرية عربية
          .replace(/٬/g, "")                            // فاصل آلاف عربي
          .replace(/[\.#\-_,;:’'′‘’“”"`@~|\\\/\+\=\(\)\[\]\{\}!?&]/g, "") // رموز
          .trim();
      }

      function cleanValue(val) {
        if (val === null || val === undefined) return "";
        return String(val).trim();
      }

      /** يحول أي قيمة لعدد حقيقي، يدعم الأرقام العربية */
      function normalizeNumber(val) {
        if (val === "" || val === null || val === undefined) return 0;
        if (typeof val === "number") return val;
        let s = String(val).trim()
          .replace(/[٠-٩]/g, d => "0123456789"["٠١٢٣٤٥٦٧٨٩".indexOf(d)])
          .replace(/٫/g, ".")
          .replace(/[^\d.,\-]/g, "");
        if (!s) return 0;
        if (s.includes(",") && s.includes(".")) {
          s = s.lastIndexOf(".") > s.lastIndexOf(",")
            ? s.replace(/,/g, "")
            : s.replace(/\./g, "").replace(",", ".");
        } else {
          s = s.replace(",", ".");
        }
        return Number(s) || 0;
      }

      // ────────────────────────────────────────────────────────────
      // 2) قاموس المسميات المحتملة (أكثر من 100 مرادف)
      // ────────────────────────────────────────────────────────────
      const templateMap = {
        "الكود": [
          // English / Intl
          "item no", "item no.", "item#", "item number", "model", "model no", "model no.", 
          "model number", "model#", "sku", "part no", "part no.", "part number",
          "ref", "reference", "product code", "product id", "product no",
          "art. no.", "art no", "artikel", "codice", "codigo",
          // Arabic
          "الكود", "كود", "code", "رقم", "id", "رقم المنتج", "كود المنتج",
          "رقم الصنف", "رقم السلعة", "رقم السلعة"
        ],
        "البند": [
          "type", "series", "line", "product line", "product type", "brand",
          "brand name", "brand series", "collection",
          "البند", "بند", "category", "نوع", "group", "قسم", "فئة", "المجموعة",
          "المجموعه", "القسم", "section", "department"
        ],
        "التصنيف": [
          "key data", "classification", "class", "segment", "sub type",
          "sub category", "categoria", "categorie",
          "التصنيف", "تصنيف", "مجموعة", "مجموعه", "الفئة الفرعية"
        ],
        "اسم المنتج": [
          "product name", "product description", "item name", "item description",
          "description", "name", "product", "item", "title",
          "description & features", "description and features",
          "goods name", "goods description", "article", "descrizione",
          // Arabic
          "اسم المنتج", "المنتج", "اسم", "الصنف", "اسم الصنف",
          "وصف المنتج", "بيان", "المادة", "الاسم", "البيان", "الوصف",
          "اسم السلعة", "السلعة", "شرح", "تعريف"
        ],
        "الوحدة": [
          "unit", "uom", "u/m", "unit of measure", "unit of measurement",
          "pack unit", "selling unit", "packing unit", "pcs/unit",
          "الوحدة", "وحدة", "وحدة القياس", "وحدة البيع", "الوحده",
          "معيار", "التعبئة", "packaging"
        ],
        "الكمية": [
          "qty", "quantity", "order qty", "order quantity",
          "total qty", "requested qty", "pcs", "pieces",
          "الكمية", "كمية", "كم", "عدد", "الكميه", "قطعة", "قطع",
          "العدد", "الكمية المطلوبة"
        ],
        "سعرالشراء": [
          // شراء / تكلفة
          "cost", "unit cost", "buy price", "purchase price",
          "fob", "fob price", "ex-factory price", "factory price",
          "net price", "net cost", "cost price", "price (exw)",
          "سعرالشراء", "سعر الشراء", "سعر شراء", "الشراء", "تكلفة", "تكلفه",
          "سعر التكلفة", "سعر المصنع", "سعر المصنعة", "سعر الشراء (ج.م.ع)",
          "تكلفة الوحدة", "سعر الجملة شراء", "prezzo di acquisto",
          "prezzo di costo"
        ],
        "سعر البيع": [
          "sell price", "selling price", "retail price", "customer price",
          "market price", "msrp", "rrp", "price", "list price",
          "البيع", "سعر البيع", "سعر بيع", "سعر البيع (ج.م.ع)",
          "بيع", "سعر العميل", "سعر القطاعي", "سعر السوق",
          "سعر البيع للمستهلك", "السعر", "prezzo di vendita"
        ],
      };

      // ────────────────────────────────────────────────────────────
      // 3) كشف صف العنوان (أول 30 صف) مع تجاهل شبه الفارغ
      // ────────────────────────────────────────────────────────────
      function findFieldForHeader(text) {
        const norm = normalizeText(text);
        if (!norm) return null;
        // exact match first
        for (const [field, aliases] of Object.entries(templateMap)) {
          if (aliases.some(a => normalizeText(a) === norm)) return field;
        }
        // partial match
        for (const [field, aliases] of Object.entries(templateMap)) {
          if (aliases.some(a => {
            const na = normalizeText(a);
            return na.length > 2 && (norm.includes(na) || na.includes(norm));
          })) return field;
        }
        return null;
      }

      function scoreRow(row) {
        const seen = new Set();
        for (const cell of row) {
          const f = findFieldForHeader(cell);
          if (f && !seen.has(f)) seen.add(f);
        }
        return seen.size;
      }

      const maxHeaderScan = Math.min(30, rawRows.length);
      let headerRowIndex = 0, bestScore = 0;
      for (let i = 0; i < maxHeaderScan; i++) {
        // تجاهل الصف إذا كان 80% منه فاضي
        const filledCount = rawRows[i].filter(c => cleanValue(c) !== "").length;
        if (filledCount === 0) continue;
        const s = scoreRow(rawRows[i]);
        if (s > bestScore) {
          bestScore = s;
          headerRowIndex = i;
        }
      }

      const headerRow = rawRows[headerRowIndex];

      // ────────────────────────────────────────────────────────────
      // 4) ربط الأعمدة بالحقول
      // ────────────────────────────────────────────────────────────
      const colIndexToField = {};
      const usedFields = new Set();

      // قائمة محظورة محسّنة للأعمدة اللي نادراً نحتاجها
      const SKIP_HEADERS = [
        "picture", "image", "photo", "pic", "صورة", "صوره",
        "key data", "packed by", "packing", "carton", "cbm",
        "gw", "nw", "gross weight", "net weight", "weight", "wgt",
        "barcode", "ean", "upc", "barcode/pcs", "barcode/ctn",
        "qty received", "qty on the way", "plan to ship",
        "total comments", "customer comments", "feedback",
        "dimensions", "volume", "length", "width", "height",
        "hs code", "customs code", "country", "origin"
      ];

      for (let col = 0; col < headerRow.length; col++) {
        const text = cleanValue(headerRow[col]);
        if (!text) continue;
        const normText = normalizeText(text);
        if (SKIP_HEADERS.some(s => normText.includes(normalizeText(s)))) continue;

        const field = findFieldForHeader(text);
        if (field && !usedFields.has(field)) {
          colIndexToField[col] = field;
          usedFields.add(field);
        }
      }

      // ────────────────────────────────────────────────────────────
      // 5) الفال باك الذكي للأعمدة الناقصة
      // ────────────────────────────────────────────────────────────
      const missingFields = Object.keys(templateMap).filter(f => !usedFields.has(f));

      if (missingFields.length > 0) {
        const sampleRows = rawRows
          .slice(headerRowIndex + 1, headerRowIndex + 21)
          .filter(r => r.some(c => cleanValue(c) !== ""));

        const unmappedCols = [];
        for (let col = 0; col < headerRow.length; col++) {
          if (colIndexToField[col]) continue;
          const vals = sampleRows.map(r => cleanValue(r[col])).filter(v => v !== "");
          if (vals.length === 0) continue;

          const nums = vals.map(v => normalizeNumber(v)).filter(n => n !== 0);
          const isNumeric = nums.length / vals.length >= 0.6;
          const avg = isNumeric ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
          const allIntegers = nums.length > 0 && nums.every(n => Number.isInteger(n));
          const hasCodePattern = vals.filter(v => /^[A-Za-z]{2,}[A-Za-z0-9]{2,}$/.test(normalizeText(v))).length / vals.length >= 0.4;

          unmappedCols.push({ col, vals, nums, isNumeric, avg, allIntegers, hasCodePattern });
        }

        // 1- كود (نمط حروف+أرقام)
        if (!usedFields.has("الكود")) {
          const codeCol = unmappedCols.find(c => c.hasCodePattern);
          if (codeCol) {
            colIndexToField[codeCol.col] = "الكود";
            usedFields.add("الكود");
          }
        }

        // 2- كمية (رقمية صحيحة، قيمتها معقولة)
        if (!usedFields.has("الكمية")) {
          const qtyCol = unmappedCols
            .filter(c => !colIndexToField[c.col] && c.isNumeric && c.allIntegers && c.avg < 50000)
            .sort((a, b) => a.avg - b.avg)[0];
          if (qtyCol) {
            colIndexToField[qtyCol.col] = "الكمية";
            usedFields.add("الكمية");
          }
        }

        // 3- سعر الشراء (الأصغر) وسعر البيع (الأكبر)
        const priceCandidates = unmappedCols
          .filter(c => !colIndexToField[c.col] && c.isNumeric && c.avg > 0.1)
          .sort((a, b) => a.avg - b.avg);

        if (!usedFields.has("سعرالشراء") && priceCandidates[0]) {
          colIndexToField[priceCandidates[0].col] = "سعرالشراء";
          usedFields.add("سعرالشراء");
        }
        if (!usedFields.has("سعر البيع") && priceCandidates[1]) {
          colIndexToField[priceCandidates[1].col] = "سعر البيع";
          usedFields.add("سعر البيع");
        }

        // 4- المحاولة الأخيرة: استخدام العناوين الفرعية كـ "اسم المنتج" إذا كان العمود يحتوي نصوصاً طويلة
        if (!usedFields.has("اسم المنتج")) {
          const textCol = unmappedCols.find(c => {
            if (colIndexToField[c.col]) return false;
            return !c.isNumeric && c.vals.some(v => v.length > 20);
          });
          if (textCol) {
            colIndexToField[textCol.col] = "اسم المنتج";
            usedFields.add("اسم المنتج");
          }
        }
      }

      // ────────────────────────────────────────────────────────────
      // 6) معالجة الصفوف (تجاهل صفوف المجاميع الذكية)
      // ────────────────────────────────────────────────────────────
      function isSummaryRow(row) {
        const sumKw = ["total", "subtotal", "مجموع", "الاجمالي", "إجمالي", "grand total", "sum", "total:"];
        // نتأكد إن الكلمة مش موجودة إلا في عمود كمي أو سعر وليس اسم منتج
        const fields = Object.keys(colIndexToField).map(Number);
        const suspectColumns = fields.filter(f => ["الكمية","سعرالشراء","سعر البيع","الكود"].includes(colIndexToField[f]));
        return row.some((cell, idx) => {
          if (!suspectColumns.includes(idx)) return false;
          const n = normalizeText(cell);
          return n && sumKw.some(k => n.includes(k));
        });
      }

      const NUMERIC_FIELDS = new Set(["الكمية", "سعرالشراء", "سعر البيع"]);
      const finalData = [];

      for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (!row.some(c => cleanValue(c) !== "")) continue;
        if (isSummaryRow(row)) continue;

        const out = {
          "الكود": "", "البند": "", "التصنيف": "", "اسم المنتج": "",
          "الوحدة": "", "الكمية": 0, "سعرالشراء": 0, "سعر البيع": 0
        };

        for (const [colIdx, field] of Object.entries(colIndexToField)) {
          const raw = cleanValue(row[parseInt(colIdx)]);
          out[field] = NUMERIC_FIELDS.has(field) ? normalizeNumber(raw) : raw;
        }

        // سطر بدون اسم منتج ولا كود يعتبر غير كافي ونستبعده
        if (!out["اسم المنتج"] && !out["الكود"]) continue;

        finalData.push(out);
      }

      // لو طلع مافيش بيانات بعد كل ده
      if (finalData.length === 0) {
        const details = Object.keys(templateMap)
          .filter(f => !usedFields.has(f))
          .join(", ");
        alert(`⚠️ مفيش بيانات اتحولت!\nأعمدة مش لاقياها: ${details || "—"}\nراجع الـ Console للتفاصيل.`);
        console.warn("التعريفات المستخدمة:", colIndexToField);
        return;
      }

      // ────────────────────────────────────────────────────────────
      // 7) تصدير الملف المرتب
      // ────────────────────────────────────────────────────────────
      const ORDERED_HEADERS = [
        "الكود", "البند", "التصنيف", "اسم المنتج",
        "الوحدة", "الكمية", "سعرالشراء", "سعر البيع"
      ];

      const ordered = finalData.map(row => {
        const obj = {};
        for (const h of ORDERED_HEADERS) obj[h] = row[h] ?? "";
        return obj;
      });

      const newSheet = XLSX.utils.json_to_sheet(ordered, { header: ORDERED_HEADERS });
      newSheet["!cols"] = [
        { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 35 },
        { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 14 }
      ];

      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Products");
      XLSX.writeFile(newWorkbook, "جاهز-للرفع.xlsx");

      const missing = Object.keys(templateMap).filter(f => !usedFields.has(f));
      const msg = missing.length
        ? `✅ تم! ${finalData.length} منتج\n⚠️ أعمدة مش لاقتهاش: ${missing.join(", ")}`
        : `✅ تم! ${finalData.length} منتج — كل الأعمدة اتحولت بنجاح`;
      alert(msg);

    } catch (err) {
      console.error("processFile error:", err);
      alert("❌ حصل خطأ: " + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}
