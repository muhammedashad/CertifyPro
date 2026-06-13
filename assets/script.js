/**
 * CertifyPro - Enterprise Certificate Generator Core Application Engine
 * SaaS Architecture, high performance, print-quality rendering & bulk processing pipelines.
 */

// Global state holding classes
let themeManager, historyManager, canvasEditor, templateLibrary, bulkProcessor, exportManager;

// Initial project configuration variables
const DEFAULT_WIDTH = 1920;
const DEFAULT_HEIGHT = 1080;

// --- Text Wrapping Helper Function ---
function wrapText(ctx, text, maxWidth) {
  const paragraphs = text.split("\n");
  const lines = [];
  
  paragraphs.forEach((p) => {
    const words = p.split(" ");
    let currentLine = "";
    
    words.forEach((word) => {
      const testLine = currentLine ? currentLine + " " + word : word;
      const metrics = ctx.measureText(testLine);
      
      if (metrics.width > maxWidth && currentLine) {
        lines.push(currentLine);
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    });
    if (currentLine) {
      lines.push(currentLine);
    }
    if (p === "") {
      lines.push("");
    }
  });
  
  return lines;
}

// --- Excel & Date Helper Functions ---
function excelDateToJSDate(serial) {
  return new Date((serial - 25569) * 86400 * 1000);
}

function tryFormatDateString(str) {
  if (str === null || str === undefined) return "";
  const trimmed = str.toString().trim();
  if (!trimmed) return "";

  // Date formats
  const ymdPattern = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/;
  const slashPattern = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/;
  const formattedPattern = /^\d{1,2}\s+[A-Za-z]+\s+\d{4}$/; // e.g. 18 March 2026

  if (formattedPattern.test(trimmed)) {
    return trimmed;
  }

  // Check if string is a raw Excel date serial number
  const num = Number(trimmed);
  if (!isNaN(num) && num > 30000 && num < 60000 && Number.isInteger(num)) {
    const d = excelDateToJSDate(num);
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      });
    }
  }

  // Check YYYY-MM-DD format
  if (ymdPattern.test(trimmed)) {
    const match = trimmed.match(ymdPattern);
    const d = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      });
    }
  }

  // Check DD/MM/YYYY or MM/DD/YYYY format
  if (slashPattern.test(trimmed)) {
    const match = trimmed.match(slashPattern);
    const p1 = parseInt(match[1]);
    const p2 = parseInt(match[2]);
    const y = parseInt(match[3]);
    
    let d = null;
    if (p1 > 12) { // DD/MM/YYYY
      d = new Date(y, p2 - 1, p1);
    } else if (p2 > 12) { // MM/DD/YYYY
      d = new Date(y, p1 - 1, p2);
    } else {
      // Ambiguous date - assume DD/MM/YYYY standard en-GB format
      d = new Date(y, p2 - 1, p1);
    }
    if (d && !isNaN(d.getTime())) {
      return d.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      });
    }
  }

  // Try standard parsing for generic strings (avoid parsing zip codes / simple IDs)
  const ts = Date.parse(trimmed);
  if (!isNaN(ts)) {
    const d = new Date(ts);
    if (isNaN(Number(trimmed))) {
      return d.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      });
    }
  }

  return trimmed;
}

function parseSheetToRows(worksheet) {
  if (!worksheet || !worksheet['!ref']) return [];
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const rows = [];
  
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
      const cell = worksheet[cellRef];
      let val = "";
      
      if (cell) {
        val = cell.v;
        
        let isDate = false;
        
        // 1. Check if SheetJS explicitly parsed it as a date
        if (cell.t === 'd' || val instanceof Date) {
          isDate = true;
        } 
        // 2. Or if cell number format contains date tokens (like y, m, d)
        else if (cell.t === 'n' && cell.z) {
          const fmt = cell.z.toLowerCase();
          if (fmt.includes('y') || fmt.includes('m') || fmt.includes('d')) {
            isDate = true;
          }
        }
        
        if (isDate) {
          if (typeof val === 'number') {
            val = excelDateToJSDate(val);
          }
          if (val instanceof Date && !isNaN(val.getTime())) {
            val = val.toLocaleDateString('en-GB', {
              day: 'numeric',
              month: 'long',
              year: 'numeric'
            });
          }
        } else if (val !== null && val !== undefined) {
          // Fallback string parser checking for CSV date strings
          val = tryFormatDateString(val.toString());
        }
        
        if (val !== undefined && val !== null && val !== "") {
          hasData = true;
        }
      }
      row.push(val === undefined || val === null ? "" : val);
    }
    
    // Keep header row and rows that have any actual content
    if (hasData || r === range.s.r) {
      rows.push(row);
    }
  }
  return rows;
}

document.addEventListener("DOMContentLoaded", () => {
  // Initialize Architecture modules
  themeManager = new ThemeManager();
  historyManager = new HistoryManager();
  canvasEditor = new CanvasEditor();
  templateLibrary = new TemplateLibrary();
  bulkProcessor = new BulkProcessor();
  exportManager = new ExportManager();

  // Load default template (Black A4 Page)
  templateLibrary.loadTemplate("black_a4");
  
  // Setup global event bindings
  setupGlobalBindings();
});

/* --- Theme Manager Module --- */
class ThemeManager {
  constructor() {
    this.themeToggleBtn = document.getElementById("theme-toggle");
    this.currentTheme = localStorage.getItem("certify-theme") || "dark";
    
    // Apply initial theme
    document.documentElement.setAttribute("data-theme", this.currentTheme);
    this.bindEvents();
  }

  bindEvents() {
    this.themeToggleBtn.addEventListener("click", () => {
      this.currentTheme = this.currentTheme === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", this.currentTheme);
      localStorage.setItem("certify-theme", this.currentTheme);
      this.pushNotification(
        `Theme switched to ${this.currentTheme === 'dark' ? 'Dark Mode' : 'Light Mode'}`
      );
    });
  }

  pushNotification(message, type = "info") {
    const list = document.getElementById("notifications-list");
    const countBadge = document.getElementById("notification-count");
    
    // Remove empty state if present
    const noNotif = list.querySelector(".no-notifications");
    if (noNotif) noNotif.remove();

    const item = document.createElement("div");
    item.className = `notification-item ${type}`;
    item.innerHTML = `
      <i class="fa-solid ${type === 'success' ? 'fa-circle-check' : (type === 'error' ? 'fa-circle-exclamation' : 'fa-bell')}"></i>
      <div class="notification-body">
        <p>${message}</p>
        <span style="font-size:9px;color:var(--text-muted);">${new Date().toLocaleTimeString()}</span>
      </div>
    `;
    list.insertBefore(item, list.firstChild);

    // Update count
    const count = parseInt(countBadge.innerText) + 1;
    countBadge.innerText = count;
    countBadge.style.display = "flex";
  }
}

/* --- History Stack Manager (Undo/Redo) --- */
class HistoryManager {
  constructor() {
    this.stack = [];
    this.index = -1;
    this.undoBtn = document.getElementById("btn-undo");
    this.redoBtn = document.getElementById("btn-redo");
    this.bindEvents();
  }

  pushState(layers) {
    // Slice stack if we performed action after undoing
    if (this.index < this.stack.length - 1) {
      this.stack = this.stack.slice(0, this.index + 1);
    }
    
    // Push deep clone of layers
    const clonedLayers = JSON.parse(JSON.stringify(layers));
    this.stack.push(clonedLayers);
    
    // Bound stack to max 50 steps
    if (this.stack.length > 50) {
      this.stack.shift();
    } else {
      this.index++;
    }
    this.updateUI();
    
    // Trigger auto-save to localStorage
    this.autoSave();

    if (exportManager) {
      exportManager.updateEstimates();
    }
  }

  undo() {
    if (this.index > 0) {
      this.index--;
      canvasEditor.layers = JSON.parse(JSON.stringify(this.stack[this.index]));
      canvasEditor.selectedLayer = null;
      canvasEditor.drawTextfromInputs();
      canvasEditor.updateLayersSidebar();
      canvasEditor.updateInspectorPanel();
      this.updateUI();
      themeManager.pushNotification("Undo action performed");

      if (exportManager) {
        exportManager.updateEstimates();
      }
    }
  }

  redo() {
    if (this.index < this.stack.length - 1) {
      this.index++;
      canvasEditor.layers = JSON.parse(JSON.stringify(this.stack[this.index]));
      canvasEditor.selectedLayer = null;
      canvasEditor.drawTextfromInputs();
      canvasEditor.updateLayersSidebar();
      canvasEditor.updateInspectorPanel();
      this.updateUI();
      themeManager.pushNotification("Redo action performed");

      if (exportManager) {
        exportManager.updateEstimates();
      }
    }
  }

  updateUI() {
    this.undoBtn.disabled = this.index <= 0;
    this.redoBtn.disabled = this.index >= this.stack.length - 1;
  }

  bindEvents() {
    this.undoBtn.addEventListener("click", () => this.undo());
    this.redoBtn.addEventListener("click", () => this.redo());
  }

  autoSave() {
    if (canvasEditor) {
      const dataToSave = {
        projectTitle: document.getElementById("project-title").value,
        layers: canvasEditor.layers,
        bgTemplateType: canvasEditor.bgTemplateType,
        customBgImageSrc: canvasEditor.customBgImageSrc
      };
      localStorage.setItem("certify-autosave", JSON.stringify(dataToSave));
    }
  }
}

/* --- Interactive Canvas Editor --- */
class CanvasEditor {
  constructor() {
    this.canvas = document.getElementById("certificatecanvas");
    this.ctx = this.canvas.getContext("2d");
    
    this.layers = [];
    this.selectedLayer = null;
    this.bgTemplateType = "classic"; // classic, corporate, creative, luxury, custom
    this.customBgImage = new Image();
    this.customBgImageSrc = null;
    
    // Zoom and grid coordinates
    this.zoomLevel = 1.0;
    this.gridSize = 20;
    this.snapToGrid = true;
    this.showGrid = false;
    
    // Drag parameters
    this.isDragging = false;
    this.dragStartX = 0;
    this.dragStartY = 0;
    
    // Drag-to-create parameters
    this.isCreatingTextBox = false;
    this.createStartX = 0;
    this.createStartY = 0;
    this.createEndX = 0;
    this.createEndY = 0;

    // Handle scaling parameters
    this.isScaling = false;
    this.scaleStartHandle = null;
    this.scaleStartFontSize = 2.0;
    this.scaleStartClickX = 0;
    this.scaleStartClickY = 0;
    this.scaleStartBoxX = 0;
    this.scaleStartBoxY = 0;
    this.scaleStartWidth = 0;
    this.scaleStartHeight = 0;
    
    // Dimensions
    this.canvas.width = DEFAULT_WIDTH;
    this.canvas.height = DEFAULT_HEIGHT;

    // Photoshop-style panning & zoom variables
    this.isPanning = false;
    this.isSpaceDown = false;
    this.panStartX = 0;
    this.panStartY = 0;
    this.scrollStartX = 0;
    this.scrollStartY = 0;
    this.scrollArea = document.getElementById("canvas-scroll-area");

    this.setupListeners();
    this.bindSettings();
  }

  setupListeners() {
    // Mouse canvas drag handlers
    this.canvas.addEventListener("mousedown", (e) => this.handleMouseDown(e));
    this.canvas.addEventListener("mousemove", (e) => this.handleMouseMove(e));
    window.addEventListener("mouseup", () => this.handleMouseUp());

    // Window resize scaling
    window.addEventListener("resize", () => this.drawTextfromInputs());

    // Photoshop zoom: Ctrl + Mouse Wheel (or trackpad pinch) zooms centered on mouse cursor
    this.scrollArea.addEventListener("wheel", (e) => this.handleCanvasWheel(e), { passive: false });

    // Photoshop pan: Spacebar & Middle click panning
    window.addEventListener("keydown", (e) => this.handleKeyDown(e));
    window.addEventListener("keyup", (e) => this.handleKeyUp(e));
    
    // Middle-click and spacebar dragging on the scroll container
    this.scrollArea.addEventListener("mousedown", (e) => this.handleScrollMouseDown(e));
    this.scrollArea.addEventListener("mousemove", (e) => this.handleScrollMouseMove(e));
    window.addEventListener("mouseup", () => this.handleScrollMouseUp());
  }

  handleCanvasWheel(e) {
    if (e.ctrlKey) {
      e.preventDefault();
      
      const rect = this.scrollArea.getBoundingClientRect();
      const mouseX = e.clientX - rect.left;
      const mouseY = e.clientY - rect.top;
      
      // Calculate cursor position relative to the scale space before zooming
      const contentX = (mouseX + this.scrollArea.scrollLeft) / this.zoomLevel;
      const contentY = (mouseY + this.scrollArea.scrollTop) / this.zoomLevel;
      
      // Zoom step (scaled relatively)
      const zoomStep = 0.05 * this.zoomLevel;
      const delta = -e.deltaY > 0 ? zoomStep : -zoomStep;
      
      this.zoomLevel = Math.max(0.15, Math.min(3.0, this.zoomLevel + delta));
      this.applyZoom();
      
      // Adjust scroll to keep cursor centered on the same coordinate
      this.scrollArea.scrollLeft = contentX * this.zoomLevel - mouseX;
      this.scrollArea.scrollTop = contentY * this.zoomLevel - mouseY;
      
      this.drawTextfromInputs();
    }
  }

  handleKeyDown(e) {
    // Ignore key binds when writing text in inputs
    if (document.activeElement.tagName === "INPUT" || document.activeElement.tagName === "TEXTAREA") {
      return;
    }

    if (e.code === "Space") {
      // Prevent standard browser scroll
      e.preventDefault();
      if (!this.isSpaceDown) {
        this.isSpaceDown = true;
        this.scrollArea.style.cursor = "grab";
      }
    }
  }

  handleKeyUp(e) {
    if (e.code === "Space") {
      this.isSpaceDown = false;
      this.scrollArea.style.cursor = "default";
    }
  }

  handleScrollMouseDown(e) {
    // Pan on middle click (button 1) or Left click + Space
    if (e.button === 1 || (e.button === 0 && this.isSpaceDown)) {
      e.preventDefault();
      this.isPanning = true;
      this.scrollArea.style.cursor = "grabbing";
      this.panStartX = e.clientX;
      this.panStartY = e.clientY;
      this.scrollStartX = this.scrollArea.scrollLeft;
      this.scrollStartY = this.scrollArea.scrollTop;
    }
  }

  handleScrollMouseMove(e) {
    if (!this.isPanning) return;
    e.preventDefault();
    const dx = e.clientX - this.panStartX;
    const dy = e.clientY - this.panStartY;
    
    this.scrollArea.scrollLeft = this.scrollStartX - dx;
    this.scrollArea.scrollTop = this.scrollStartY - dy;
  }

  handleScrollMouseUp() {
    if (this.isPanning) {
      this.isPanning = false;
      this.scrollArea.style.cursor = this.isSpaceDown ? "grab" : "default";
    }
  }

  bindSettings() {
    // Bind sidebar controls
    document.getElementById("setting-show-grid").addEventListener("change", (e) => {
      this.showGrid = e.target.checked;
      this.drawTextfromInputs();
    });

    document.getElementById("setting-snap-grid").addEventListener("change", (e) => {
      this.snapToGrid = e.target.checked;
    });

    document.getElementById("setting-grid-size").addEventListener("input", (e) => {
      this.gridSize = parseInt(e.target.value) || 20;
      this.drawTextfromInputs();
    });

    // Zoom Buttons
    document.getElementById("btn-zoom-in").addEventListener("click", () => this.adjustZoom(0.1));
    document.getElementById("btn-zoom-out").addEventListener("click", () => this.adjustZoom(-0.1));
    document.getElementById("btn-fit-screen").addEventListener("click", () => this.fitToScreen());

    // Drag-over custom backgrounds
    const bgZone = document.getElementById("bg-drag-zone");
    bgZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      bgZone.classList.add("dragover");
    });
    bgZone.addEventListener("dragleave", () => bgZone.classList.remove("dragover"));
    bgZone.addEventListener("drop", (e) => {
      e.preventDefault();
      bgZone.classList.remove("dragover");
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        this.loadCustomBgImage(e.dataTransfer.files[0]);
      }
    });

    document.getElementById("upload-bg-input").addEventListener("change", (e) => {
      if (e.target.files && e.target.files[0]) {
        this.loadCustomBgImage(e.target.files[0]);
      }
    });

    // Add Preset text actions
    document.getElementById("add-heading-btn").addEventListener("click", () => this.addTextLayer("Heading Accent", "heading"));
    document.getElementById("add-subheading-btn").addEventListener("click", () => this.addTextLayer("Subheading Text Description", "subheading"));
    document.getElementById("add-body-btn").addEventListener("click", () => this.addTextLayer("Body placeholder values", "body"));
  }

  loadCustomBgImage(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      this.bgTemplateType = "custom";
      this.customBgImageSrc = e.target.result;
      this.customBgImage.src = e.target.result;
      this.customBgImage.onload = () => {
        // Match project canvas dimensions to file aspect ratio
        this.canvas.width = this.customBgImage.width;
        this.canvas.height = this.customBgImage.height;
        this.drawTextfromInputs();
        themeManager.pushNotification("Custom certificate background uploaded", "success");
        historyManager.pushState(this.layers);
      };
    };
    reader.readAsDataURL(file);
  }

  adjustZoom(delta) {
    this.zoomLevel = Math.max(0.25, Math.min(2.5, this.zoomLevel + delta));
    this.applyZoom();
  }

  fitToScreen() {
    const scrollArea = document.getElementById("canvas-scroll-area");
    const ratio = Math.min(
      (scrollArea.clientWidth - 80) / this.canvas.width,
      (scrollArea.clientHeight - 80) / this.canvas.height
    );
    this.zoomLevel = ratio;
    this.applyZoom();
  }

  applyZoom() {
    const viewport = document.getElementById("canvas-viewport");
    viewport.style.transform = `scale(${this.zoomLevel})`;
    document.getElementById("zoom-percentage").innerText = `${Math.round(this.zoomLevel * 100)}%`;
  }

  drawTextfromInputs(isExporting = false) {
    // Clear canvas
    this.ctx.fillStyle = "white";
    this.ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);

    // Draw background
    if (this.bgTemplateType === "custom" && this.customBgImage.complete && this.customBgImage.naturalWidth > 0) {
      this.ctx.drawImage(this.customBgImage, 0, 0, this.canvas.width, this.canvas.height);
    } else {
      // Procedural Vector Background Design
      renderTemplateBackground(this.bgTemplateType, this.canvas, this.ctx);
    }

    // Render grid overlay (only on editor view, not in exports)
    if (this.showGrid && !isExporting) {
      this.drawGrid();
    }

    // Draw individual layers
    this.layers.forEach((layer) => {
      if (layer.visible) {
        this.drawLayer(layer);
      }
    });

    // Draw interactive bounding selector if selected (not in export)
    if (this.selectedLayer && !isExporting) {
      this.drawBoundingBox();
    }

    // Draw drag-to-create preview box (not in export)
    if (this.isCreatingTextBox && !isExporting) {
      this.ctx.save();
      this.ctx.strokeStyle = "rgba(37, 99, 235, 0.6)";
      this.ctx.lineWidth = 2;
      this.ctx.setLineDash([6, 4]);
      
      const x = this.createStartX;
      const y = this.createStartY;
      const w = this.createEndX - this.createStartX;
      const h = this.createEndY - this.createStartY;
      
      this.ctx.strokeRect(x, y, w, h);
      this.ctx.fillStyle = "rgba(37, 99, 235, 0.05)";
      this.ctx.fillRect(x, y, w, h);
      
      this.ctx.restore();
    }
  }

  drawLayer(layer) {
    this.ctx.save();

    // Map properties
    const defaultFontSizeVal = this.canvas.width / 100;
    const size = layer.fontSize * defaultFontSizeVal;
    
    // Typography properties
    this.ctx.font = `${layer.fontWeight === 'bold' ? 'bold' : ''} ${layer.fontStyle === 'italic' ? 'italic' : ''} ${size}px "${layer.fontFamily}", sans-serif`;
    this.ctx.globalAlpha = layer.opacity / 100;
    this.ctx.textAlign = layer.textAlign;
    this.ctx.textBaseline = "top";
    
    // Apply Letter spacing if supported
    if ('letterSpacing' in this.ctx) {
      this.ctx.letterSpacing = `${layer.letterSpacing || 0}px`;
    }

    const xPos = (layer.x / 100) * this.canvas.width;
    const yPos = (layer.y / 100) * this.canvas.height;
    
    // Text value (checking for temporary bulk overrides)
    const displayText = layer.tempValue !== undefined ? layer.tempValue : layer.text;
    
    let lines = [];
    if (layer.width !== undefined && layer.width !== null) {
      const maxWidth = (layer.width / 100) * this.canvas.width;
      lines = wrapText(this.ctx, displayText, maxWidth);
      layer.widthPx = maxWidth;
    } else {
      lines = displayText.split("\n");
      let maxLineWidth = 0;
      lines.forEach((line) => {
        const metrics = this.ctx.measureText(line);
        if (metrics.width > maxLineWidth) {
          maxLineWidth = metrics.width;
        }
      });
      layer.widthPx = maxLineWidth;
    }
    const lineHeights = size * (layer.lineHeight || 1.2);
    layer.heightPx = lineHeights * lines.length;

    // Effects: Text Shadow / Glow
    if (layer.shadowEnabled) {
      this.ctx.shadowOffsetX = layer.shadowX || 2;
      this.ctx.shadowOffsetY = layer.shadowY || 2;
      this.ctx.shadowBlur = layer.shadowBlur || 4;
      this.ctx.shadowColor = layer.shadowColor || "rgba(0,0,0,0.5)";
    } else if (layer.glowEnabled) {
      this.ctx.shadowOffsetX = 0;
      this.ctx.shadowOffsetY = 0;
      this.ctx.shadowBlur = layer.glowRadius || 10;
      this.ctx.shadowColor = layer.glowColor || "#06B6D4";
    }

    // Colors: Solid vs Gradient fills
    if (layer.colorType === "gradient") {
      const angleRad = ((layer.gradientAngle || 90) * Math.PI) / 180;
      let startX = xPos;
      if (layer.textAlign === "center") startX = xPos - layer.widthPx / 2;
      else if (layer.textAlign === "right") startX = xPos - layer.widthPx;

      const grad = this.ctx.createLinearGradient(
        startX, 
        yPos, 
        startX + layer.widthPx, 
        yPos + layer.heightPx
      );
      grad.addColorStop(0, layer.gradientStart || "#2563EB");
      grad.addColorStop(1, layer.gradientEnd || "#7C3AED");
      this.ctx.fillStyle = grad;
    } else {
      this.ctx.fillStyle = layer.color || "#000000";
    }

    // Draw each line
    lines.forEach((line, index) => {
      const lineY = yPos + index * lineHeights;

      // Apply Stroke Border if active
      if (layer.strokeEnabled) {
        this.ctx.strokeStyle = layer.strokeColor || "#000000";
        this.ctx.lineWidth = layer.strokeWidth || 1;
        this.ctx.strokeText(line, xPos, lineY);
      }

      // Draw main Text layout
      this.ctx.fillText(line, xPos, lineY);
    });

    this.ctx.restore();
  }

  drawBoundingBox() {
    const l = this.selectedLayer;
    const xPos = (l.x / 100) * this.canvas.width;
    const yPos = (l.y / 100) * this.canvas.height;
    
    let boxX = xPos;
    if (l.textAlign === "center") {
      boxX = xPos - l.widthPx / 2;
    } else if (l.textAlign === "right") {
      boxX = xPos - l.widthPx;
    }

    // Bounding Box Frame
    this.ctx.save();
    this.ctx.strokeStyle = "#2563EB";
    this.ctx.lineWidth = 2;
    this.ctx.setLineDash([4, 4]);
    this.ctx.strokeRect(boxX - 6, yPos - 6, l.widthPx + 12, l.heightPx + 12);
    
    // Draw solid corners
    this.ctx.fillStyle = "#FFFFFF";
    this.ctx.strokeStyle = "#2563EB";
    this.ctx.lineWidth = 2;
    this.ctx.setLineDash([]); // reset dash

    const corners = [
      [boxX - 6, yPos - 6], // Top Left
      [boxX + l.widthPx + 6, yPos - 6], // Top Right
      [boxX - 6, yPos + l.heightPx + 6], // Bottom Left
      [boxX + l.widthPx + 6, yPos + l.heightPx + 6] // Bottom Right
    ];

    corners.forEach(([cx, cy]) => {
      this.ctx.beginPath();
      this.ctx.arc(cx, cy, 5, 0, Math.PI * 2);
      this.ctx.fill();
      this.ctx.stroke();
    });

    this.ctx.restore();
  }

  drawGrid() {
    this.ctx.save();
    this.ctx.strokeStyle = "rgba(37, 99, 235, 0.08)";
    this.ctx.lineWidth = 1;
    
    // Horizontal lines
    for (let y = this.gridSize; y < this.canvas.height; y += this.gridSize) {
      this.ctx.beginPath();
      this.ctx.moveTo(0, y);
      this.ctx.lineTo(this.canvas.width, y);
      this.ctx.stroke();
    }
    
    // Vertical lines
    for (let x = this.gridSize; x < this.canvas.width; x += this.gridSize) {
      this.ctx.beginPath();
      this.ctx.moveTo(x, 0);
      this.ctx.lineTo(x, this.canvas.height);
      this.ctx.stroke();
    }
    this.ctx.restore();
  }

  textHittest(mx, my, layer) {
    if (!layer.visible) return false;
    const xPos = (layer.x / 100) * this.canvas.width;
    const yPos = (layer.y / 100) * this.canvas.height;
    
    let boxX = xPos;
    if (layer.textAlign === "center") {
      boxX = xPos - layer.widthPx / 2;
    } else if (layer.textAlign === "right") {
      boxX = xPos - layer.widthPx;
    }

    return (
      mx >= boxX &&
      mx <= boxX + layer.widthPx &&
      my >= yPos &&
      my <= yPos + layer.heightPx
    );
  }

  handleMouseDown(e) {
    // Bypass element highlights and dragging when spacebar-panning or middle clicking
    if (this.isSpaceDown || e.button === 1) {
      return;
    }
    e.preventDefault();
    const rect = this.canvas.getBoundingClientRect();
    
    // Calculate click coordinates mapped relative to actual resolution canvas size
    const clickX = ((e.clientX - rect.left) / rect.width) * this.canvas.width;
    const clickY = ((e.clientY - rect.top) / rect.height) * this.canvas.height;

    // 1. First check if clicked on handles of the currently selected layer
    if (this.selectedLayer) {
      const l = this.selectedLayer;
      const xPos = (l.x / 100) * this.canvas.width;
      const yPos = (l.y / 100) * this.canvas.height;
      let boxX = xPos;
      if (l.textAlign === "center") {
        boxX = xPos - l.widthPx / 2;
      } else if (l.textAlign === "right") {
        boxX = xPos - l.widthPx;
      }

      const handles = {
        topLeft: { x: boxX - 6, y: yPos - 6 },
        topRight: { x: boxX + l.widthPx + 6, y: yPos - 6 },
        bottomLeft: { x: boxX - 6, y: yPos + l.heightPx + 6 },
        bottomRight: { x: boxX + l.widthPx + 6, y: yPos + l.heightPx + 6 }
      };

      const hitRadius = 15; // click range
      let hitHandle = null;
      Object.entries(handles).forEach(([name, pos]) => {
        if (Math.hypot(clickX - pos.x, clickY - pos.y) < hitRadius) {
          hitHandle = name;
        }
      });

      if (hitHandle) {
        this.isScaling = true;
        this.scaleStartHandle = hitHandle;
        this.scaleStartFontSize = l.fontSize;
        this.scaleStartClickX = clickX;
        this.scaleStartClickY = clickY;
        this.scaleStartBoxX = boxX;
        this.scaleStartBoxY = yPos;
        this.scaleStartWidth = l.widthPx;
        this.scaleStartHeight = l.heightPx;
        this.isDragging = false;
        return;
      }
    }

    // Check hit tests in reverse order (to click topmost elements first)
    let found = null;
    for (let i = this.layers.length - 1; i >= 0; i--) {
      const layer = this.layers[i];
      if (this.textHittest(clickX, clickY, layer) && !layer.locked) {
        found = layer;
        break;
      }
    }

    if (found) {
      this.selectedLayer = found;
      this.isDragging = true;
      
      // Store start drag offsets
      this.dragStartX = clickX - (found.x / 100) * this.canvas.width;
      this.dragStartY = clickY - (found.y / 100) * this.canvas.height;

      this.updateInspectorPanel();
      this.updateLayersSidebar();
      this.drawTextfromInputs();
    } else {
      // Clear selection
      this.selectedLayer = null;
      this.updateInspectorPanel();
      this.updateLayersSidebar();
      this.drawTextfromInputs();

      // Start drag-to-create text box!
      this.isCreatingTextBox = true;
      this.createStartX = clickX;
      this.createStartY = clickY;
      this.createEndX = clickX;
      this.createEndY = clickY;
    }
  }

  handleMouseMove(e) {
    const rect = this.canvas.getBoundingClientRect();
    const clickX = ((e.clientX - rect.left) / rect.width) * this.canvas.width;
    const clickY = ((e.clientY - rect.top) / rect.height) * this.canvas.height;

    // A. Drag-to-create active
    if (this.isCreatingTextBox) {
      e.preventDefault();
      this.createEndX = clickX;
      this.createEndY = clickY;
      this.drawTextfromInputs();
      return;
    }

    // B. Interactive scaling active
    if (this.isScaling && this.selectedLayer) {
      e.preventDefault();
      const l = this.selectedLayer;
      
      // Opposite corner remains stationary as the scaling anchor
      let anchorX = 0;
      let anchorY = 0;
      let handleStartX = 0;
      let handleStartY = 0;

      if (this.scaleStartHandle === "bottomRight") {
        anchorX = this.scaleStartBoxX;
        anchorY = this.scaleStartBoxY;
        handleStartX = this.scaleStartBoxX + this.scaleStartWidth;
        handleStartY = this.scaleStartBoxY + this.scaleStartHeight;
      } else if (this.scaleStartHandle === "bottomLeft") {
        anchorX = this.scaleStartBoxX + this.scaleStartWidth;
        anchorY = this.scaleStartBoxY;
        handleStartX = this.scaleStartBoxX;
        handleStartY = this.scaleStartBoxY + this.scaleStartHeight;
      } else if (this.scaleStartHandle === "topRight") {
        anchorX = this.scaleStartBoxX;
        anchorY = this.scaleStartBoxY + this.scaleStartHeight;
        handleStartX = this.scaleStartBoxX + this.scaleStartWidth;
        handleStartY = this.scaleStartBoxY;
      } else if (this.scaleStartHandle === "topLeft") {
        anchorX = this.scaleStartBoxX + this.scaleStartWidth;
        anchorY = this.scaleStartBoxY + this.scaleStartHeight;
        handleStartX = this.scaleStartBoxX;
        handleStartY = this.scaleStartBoxY;
      }

      // Calculate distances to anchor
      const originalDistance = Math.hypot(handleStartX - anchorX, handleStartY - anchorY);
      const newDistance = Math.hypot(clickX - anchorX, clickY - anchorY);
      
      if (originalDistance > 5) {
        const scaleFactor = newDistance / originalDistance;
        const newFontSize = Math.max(0.5, Math.min(30.0, this.scaleStartFontSize * scaleFactor));
        
        // Calculate new bounding box sizes
        const newW = this.scaleStartWidth * scaleFactor;
        const newH = this.scaleStartHeight * scaleFactor;

        // Position adjustment to keep anchor stationary
        let newBoxX = this.scaleStartBoxX;
        let newBoxY = this.scaleStartBoxY;

        if (this.scaleStartHandle === "bottomRight") {
          newBoxX = anchorX;
          newBoxY = anchorY;
        } else if (this.scaleStartHandle === "bottomLeft") {
          newBoxX = anchorX - newW;
          newBoxY = anchorY;
        } else if (this.scaleStartHandle === "topRight") {
          newBoxX = anchorX;
          newBoxY = anchorY - newH;
        } else if (this.scaleStartHandle === "topLeft") {
          newBoxX = anchorX - newW;
          newBoxY = anchorY - newH;
        }

        // Translate back to aligned xPos, yPos
        let newXPos = newBoxX;
        if (l.textAlign === "center") {
          newXPos = newBoxX + newW / 2;
        } else if (l.textAlign === "right") {
          newXPos = newBoxX + newW;
        }
        let newYPos = newBoxY;

        // Apply back to model layer (percentages)
        l.x = parseFloat(((newXPos / this.canvas.width) * 100).toFixed(1));
        l.y = parseFloat(((newYPos / this.canvas.height) * 100).toFixed(1));
        l.fontSize = parseFloat(newFontSize.toFixed(1));

        // Update inspector inputs
        document.getElementById("prop-pos-x").value = l.x;
        document.getElementById("prop-pos-y").value = l.y;
        document.getElementById("prop-font-size").value = l.fontSize;

        this.drawTextfromInputs();
      }
      return;
    }

    // Change cursor on handle hover
    if (this.selectedLayer && !this.isDragging && !this.isScaling && !this.isCreatingTextBox && !this.isSpaceDown) {
      const l = this.selectedLayer;
      const xPos = (l.x / 100) * this.canvas.width;
      const yPos = (l.y / 100) * this.canvas.height;
      let boxX = xPos;
      if (l.textAlign === "center") {
        boxX = xPos - l.widthPx / 2;
      } else if (l.textAlign === "right") {
        boxX = xPos - l.widthPx;
      }

      const handles = {
        topLeft: { x: boxX - 6, y: yPos - 6, cursor: "nwse-resize" },
        topRight: { x: boxX + l.widthPx + 6, y: yPos - 6, cursor: "nesw-resize" },
        bottomLeft: { x: boxX - 6, y: yPos + l.heightPx + 6, cursor: "nesw-resize" },
        bottomRight: { x: boxX + l.widthPx + 6, y: yPos + l.heightPx + 6, cursor: "nwse-resize" }
      };

      let hoverCursor = "default";
      const hitRadius = 15;
      Object.values(handles).forEach((h) => {
        if (Math.hypot(clickX - h.x, clickY - h.y) < hitRadius) {
          hoverCursor = h.cursor;
        }
      });
      this.canvas.style.cursor = hoverCursor;
    }

    // C. Traditional dragging active
    if (!this.isDragging || !this.selectedLayer) return;
    e.preventDefault();

    let targetX = ((clickX - this.dragStartX) / this.canvas.width) * 100;
    let targetY = ((clickY - this.dragStartY) / this.canvas.height) * 100;

    // Snap to grid calculations
    if (this.snapToGrid) {
      const gridPctX = (this.gridSize / this.canvas.width) * 100;
      const gridPctY = (this.gridSize / this.canvas.height) * 100;
      targetX = Math.round(targetX / gridPctX) * gridPctX;
      targetY = Math.round(targetY / gridPctY) * gridPctY;
    }

    // Snap to horizontal/vertical centers (50%)
    const guideH = document.getElementById("guide-horizontal");
    const guideV = document.getElementById("guide-vertical");
    
    if (Math.abs(targetX - 50) < 1.0) {
      targetX = 50.0;
      guideV.style.display = "block";
      guideV.style.left = `${(50 / 100) * rect.width}px`;
    } else {
      guideV.style.display = "none";
    }

    if (Math.abs(targetY - 50) < 1.0) {
      targetY = 50.0;
      guideH.style.display = "block";
      guideH.style.top = `${(50 / 100) * rect.height}px`;
    } else {
      guideH.style.display = "none";
    }

    this.selectedLayer.x = parseFloat(targetX.toFixed(1));
    this.selectedLayer.y = parseFloat(targetY.toFixed(1));

    // Update coordinates panel
    document.getElementById("prop-pos-x").value = this.selectedLayer.x;
    document.getElementById("prop-pos-y").value = this.selectedLayer.y;

    this.drawTextfromInputs();
  }

  handleMouseUp() {
    if (this.isCreatingTextBox) {
      this.isCreatingTextBox = false;
      
      const x = Math.min(this.createStartX, this.createEndX);
      const y = Math.min(this.createStartY, this.createEndY);
      const w = Math.abs(this.createEndX - this.createStartX);
      const h = Math.abs(this.createEndY - this.createStartY);
      
      // If the dragged rectangle is large enough, create a text box
      if (w > 20 && h > 20) {
        const pctX = (x / this.canvas.width) * 100;
        const pctY = (y / this.canvas.height) * 100;
        
        // Keep standard default font size (2.0) on creation as requested
        this.addTextLayerAt("Enter text here...", pctX, pctY, 2.0);
      } else {
        this.drawTextfromInputs();
      }
    }

    if (this.isScaling) {
      this.isScaling = false;
      historyManager.pushState(this.layers);
    }

    if (this.isDragging) {
      this.isDragging = false;
      document.getElementById("guide-horizontal").style.display = "none";
      document.getElementById("guide-vertical").style.display = "none";
      historyManager.pushState(this.layers);
    }
  }

  addTextLayer(defaultText, stylePreset = "body") {
    let size = 2.0;
    let weight = "normal";
    let font = "Montserrat";
    
    if (stylePreset === "heading") {
      size = 5.0;
      weight = "bold";
      font = "Outfit";
    } else if (stylePreset === "subheading") {
      size = 3.2;
      weight = "normal";
      font = "Montserrat";
    }

    const layer = {
      id: "layer_" + Date.now(),
      text: defaultText,
      x: 50.0,
      y: 50.0 + (this.layers.length * 4), // slightly staggered coordinates
      fontFamily: font,
      fontSize: size,
      fontWeight: weight,
      fontStyle: "normal",
      colorType: "solid",
      color: "#0F172A",
      gradientStart: "#2563EB",
      gradientEnd: "#7C3AED",
      gradientAngle: 90,
      opacity: 100,
      letterSpacing: 0,
      lineHeight: 1.2,
      shadowEnabled: false,
      shadowX: 2,
      shadowY: 2,
      shadowBlur: 4,
      shadowColor: "#000000",
      strokeEnabled: false,
      strokeWidth: 1,
      strokeColor: "#000000",
      glowEnabled: false,
      glowRadius: 10,
      glowColor: "#06B6D4",
      locked: false,
      visible: true,
      editable: true
    };

    this.layers.push(layer);
    this.selectedLayer = layer;
    
    this.drawTextfromInputs();
    this.updateLayersSidebar();
    this.updateInspectorPanel();
    historyManager.pushState(this.layers);
    themeManager.pushNotification("Added new text layer");
  }

  addTextLayerAt(defaultText, pctX, pctY, size = 2.0) {
    const layer = {
      id: "layer_" + Date.now(),
      text: defaultText,
      x: parseFloat(pctX.toFixed(1)),
      y: parseFloat(pctY.toFixed(1)),
      fontFamily: "Montserrat",
      fontSize: size,
      fontWeight: "normal",
      fontStyle: "normal",
      colorType: "solid",
      color: canvasEditor.bgTemplateType.includes("black") ? "#FFFFFF" : "#0F172A",
      gradientStart: "#2563EB",
      gradientEnd: "#7C3AED",
      gradientAngle: 90,
      opacity: 100,
      letterSpacing: 0,
      lineHeight: 1.2,
      shadowEnabled: false,
      shadowX: 2,
      shadowY: 2,
      shadowBlur: 4,
      shadowColor: "#000000",
      strokeEnabled: false,
      strokeWidth: 1,
      strokeColor: "#000000",
      glowEnabled: false,
      glowRadius: 10,
      glowColor: "#06B6D4",
      locked: false,
      visible: true,
      editable: true
    };

    this.layers.push(layer);
    this.selectedLayer = layer;
    
    this.drawTextfromInputs();
    this.updateLayersSidebar();
    this.updateInspectorPanel();
    historyManager.pushState(this.layers);
    themeManager.pushNotification("Added text layer at position");
  }

  duplicateLayer() {
    if (!this.selectedLayer) return;
    const clone = JSON.parse(JSON.stringify(this.selectedLayer));
    clone.id = "layer_" + Date.now();
    clone.x = Math.min(100, clone.x + 3);
    clone.y = Math.min(100, clone.y + 3);
    
    this.layers.push(clone);
    this.selectedLayer = clone;
    
    this.drawTextfromInputs();
    this.updateLayersSidebar();
    this.updateInspectorPanel();
    historyManager.pushState(this.layers);
    themeManager.pushNotification("Layer duplicated");
  }

  deleteLayer() {
    if (!this.selectedLayer) return;
    this.layers = this.layers.filter(l => l.id !== this.selectedLayer.id);
    this.selectedLayer = null;
    
    this.drawTextfromInputs();
    this.updateLayersSidebar();
    this.updateInspectorPanel();
    historyManager.pushState(this.layers);
    themeManager.pushNotification("Layer deleted", "error");
  }

  updateLayersSidebar() {
    const list = document.getElementById("layers-list-container");
    list.innerHTML = "";
    
    if (this.layers.length === 0) {
      list.innerHTML = `<div class="empty-state">No active text layers. Add one above!</div>`;
      return;
    }

    // Render in reverse to show upper layers first
    [...this.layers].reverse().forEach((layer) => {
      const item = document.createElement("div");
      item.className = `layer-item ${this.selectedLayer && this.selectedLayer.id === layer.id ? 'active' : ''}`;
      
      item.innerHTML = `
        <div class="layer-item-left">
          <i class="fa-solid fa-grip-vertical layer-item-drag-handle"></i>
          <span class="layer-item-name" title="${layer.text}">${layer.text}</span>
        </div>
        <div class="layer-item-actions">
          <button class="layer-action-btn btn-lock ${layer.locked ? 'active' : ''}" title="Lock Layer">
            <i class="fa-solid ${layer.locked ? 'fa-lock' : 'fa-lock-open'}"></i>
          </button>
          <button class="layer-action-btn btn-vis ${!layer.visible ? 'active' : ''}" title="Toggle Visibility">
            <i class="fa-solid ${layer.visible ? 'fa-eye' : 'fa-eye-slash'}"></i>
          </button>
          <button class="layer-action-btn btn-del" title="Delete Layer">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;

      // Click to select layer
      item.addEventListener("mousedown", (e) => {
        // avoid selecting if clicking action icons
        if (e.target.closest(".layer-action-btn")) return;
        this.selectedLayer = layer;
        this.updateLayersSidebar();
        this.updateInspectorPanel();
        this.drawTextfromInputs();
      });

      // Actions event bindings
      item.querySelector(".btn-lock").addEventListener("click", () => {
        layer.locked = !layer.locked;
        if (this.selectedLayer && this.selectedLayer.id === layer.id) {
          this.selectedLayer = null;
        }
        this.updateLayersSidebar();
        this.updateInspectorPanel();
        this.drawTextfromInputs();
        historyManager.pushState(this.layers);
      });

      item.querySelector(".btn-vis").addEventListener("click", () => {
        layer.visible = !layer.visible;
        this.updateLayersSidebar();
        this.drawTextfromInputs();
        historyManager.pushState(this.layers);
      });

      item.querySelector(".btn-del").addEventListener("click", () => {
        this.selectedLayer = layer;
        this.deleteLayer();
      });

      list.appendChild(item);
    });
  }

  updateInspectorPanel() {
    const activePanel = document.getElementById("properties-active");
    const emptyPanel = document.getElementById("properties-empty");
    const label = document.getElementById("layer-type-label");

    if (!this.selectedLayer) {
      activePanel.style.display = "none";
      emptyPanel.style.display = "flex";
      label.innerText = "No Selection";
      return;
    }

    activePanel.style.display = "flex";
    emptyPanel.style.display = "none";
    label.innerText = "Text Layer";

    const l = this.selectedLayer;

    // Fill in typography inspector inputs
    document.getElementById("prop-text-val").value = l.text;
    document.getElementById("prop-pos-x").value = l.x;
    document.getElementById("prop-pos-y").value = l.y;
    document.getElementById("prop-width").value = l.width !== undefined && l.width !== null ? l.width : "";
    document.getElementById("prop-font-family").value = l.fontFamily;
    document.getElementById("prop-font-size").value = l.fontSize;
    document.getElementById("prop-opacity").value = l.opacity;
    document.getElementById("prop-letter-spacing").value = l.letterSpacing || 0;
    document.getElementById("prop-line-height").value = l.lineHeight || 1.2;

    // Bold / Italic toggles
    document.getElementById("prop-btn-bold").classList.toggle("active", l.fontWeight === "bold");
    document.getElementById("prop-btn-italic").classList.toggle("active", l.fontStyle === "italic");

    // Alignment states
    document.getElementById("prop-align-left").classList.toggle("active", l.textAlign === "left");
    document.getElementById("prop-align-center").classList.toggle("active", l.textAlign === "center");
    document.getElementById("prop-align-right").classList.toggle("active", l.textAlign === "right");

    // Colors: Solid vs Gradient toggles
    document.getElementById("color-mode-solid").classList.toggle("active", l.colorType === "solid");
    document.getElementById("color-mode-gradient").classList.toggle("active", l.colorType === "gradient");
    document.getElementById("solid-color-view").style.display = l.colorType === "solid" ? "block" : "none";
    document.getElementById("gradient-color-view").style.display = l.colorType === "gradient" ? "block" : "none";

    document.getElementById("prop-color-solid").value = l.color;
    document.getElementById("prop-color-solid-hex").value = l.color;

    document.getElementById("prop-gradient-color-1").value = l.gradientStart;
    document.getElementById("prop-gradient-color-2").value = l.gradientEnd;
    document.getElementById("prop-gradient-angle").value = l.gradientAngle;
    document.getElementById("gradient-angle-val").innerText = `${l.gradientAngle}°`;

    // Shadows
    document.getElementById("effect-shadow-enabled").checked = l.shadowEnabled;
    document.getElementById("shadow-params").style.display = l.shadowEnabled ? "flex" : "none";
    document.getElementById("prop-shadow-x").value = l.shadowX;
    document.getElementById("prop-shadow-y").value = l.shadowY;
    document.getElementById("prop-shadow-blur").value = l.shadowBlur;
    document.getElementById("prop-shadow-color").value = l.shadowColor;

    // Stroke
    document.getElementById("effect-stroke-enabled").checked = l.strokeEnabled;
    document.getElementById("stroke-params").style.display = l.strokeEnabled ? "block" : "none";
    document.getElementById("prop-stroke-width").value = l.strokeWidth;
    document.getElementById("prop-stroke-color").value = l.strokeColor;

    // Glow
    document.getElementById("effect-glow-enabled").checked = l.glowEnabled;
    document.getElementById("glow-params").style.display = l.glowEnabled ? "block" : "none";
    document.getElementById("prop-glow-radius").value = l.glowRadius;
    document.getElementById("prop-glow-color").value = l.glowColor;
  }
}

/* --- Template Library Manager --- */
class TemplateLibrary {
  constructor() {
    this.container = document.getElementById("template-library-container");
    this.templates = {
      black_a4: {
        title: "Black A4 Page",
        desc: "Minimalist landscape black canvas of A4 proportions.",
        bgType: "black_a4"
      },
      black_a4_portrait: {
        title: "Black A4 Portrait Page",
        desc: "Minimalist portrait black canvas of A4 proportions.",
        bgType: "black_a4_portrait"
      },
      classic: {
        title: "Classic Academy",
        desc: "Traditional diplomatic gold frames, serif fonts.",
        bgType: "classic"
      },
      corporate: {
        title: "Corporate Modern",
        desc: "Clean triangles and geometric modern grids.",
        bgType: "corporate"
      },
      creative: {
        title: "Creative Minimalist",
        desc: "Pastel gradient outer border and crosshairs.",
        bgType: "creative"
      },
      luxury: {
        title: "Luxury Diplomatic",
        desc: "Gold ornate border, dark slate canvas, stamp seal.",
        bgType: "luxury"
      }
    };
    this.renderTemplateList();
  }

  renderTemplateList() {
    this.container.innerHTML = "";
    Object.entries(this.templates).forEach(([key, value]) => {
      const card = document.createElement("div");
      card.className = "template-card";
      card.innerHTML = `
        <div class="temp-img-placeholder">
          <!-- Thumbnail preview drawn procedurally -->
          <canvas width="320" height="180" class="thumb-canvas" id="thumb-${key}"></canvas>
        </div>
        <div class="temp-info">
          <div class="temp-title">${value.title}</div>
          <div class="temp-desc">${value.desc}</div>
        </div>
      `;

      card.addEventListener("click", () => {
        this.loadTemplate(key);
      });

      this.container.appendChild(card);
      
      // Draw thumbnail preview asynchronously
      setTimeout(() => {
        const thumbCanvas = document.getElementById(`thumb-${key}`);
        if (thumbCanvas) {
          const tctx = thumbCanvas.getContext("2d");
          renderTemplateBackground(value.bgType, thumbCanvas, tctx);
        }
      }, 50);
    });
  }

  loadTemplate(key) {
    const t = this.templates[key];
    if (!t) return;
    
    canvasEditor.bgTemplateType = t.bgType;
    if (key === "black_a4") {
      canvasEditor.canvas.width = 1920;
      canvasEditor.canvas.height = 1358; // A4 aspect ratio 297/210
    } else if (key === "black_a4_portrait") {
      canvasEditor.canvas.width = 1358;
      canvasEditor.canvas.height = 1920; // A4 portrait aspect ratio 210/297
    } else {
      canvasEditor.canvas.width = DEFAULT_WIDTH;
      canvasEditor.canvas.height = DEFAULT_HEIGHT;
    }
    
    // Inject preset layers depending on design choice
    canvasEditor.layers = [];
    
    if (key === "black_a4" || key === "black_a4_portrait") {
      canvasEditor.layers = [];
    } else if (key === "classic") {
      canvasEditor.layers = [
        this.createLayer("CERTIFICATE OF COMPLETION", 50, 16, "Cinzel", 4.0, "bold"),
        this.createLayer("PROUDLY CONFERRED UPON", 50, 30, "Montserrat", 1.4, "normal", "#64748B", 3),
        this.createLayer("John Doe", 50, 42, "Alex Brush", 7.2, "normal", "#B45309"),
        this.createLayer("For outstanding performance in completing the Executive Software Engineering course.", 50, 58, "Playfair Display", 1.8, "normal", "#334155"),
        this.createLayer("DATE", 25, 78, "Montserrat", 1.0, "bold", "#94A3B8"),
        this.createLayer("June 12, 2026", 25, 73, "Playfair Display", 1.6, "normal", "#1E293B"),
        this.createLayer("AUTHORIZED SIGNATURE", 75, 78, "Montserrat", 1.0, "bold", "#94A3B8"),
        this.createLayer("Muhammed Ashad K", 75, 72, "Alex Brush", 2.6, "normal", "#B45309")
      ];
    } else if (key === "corporate") {
      canvasEditor.layers = [
        this.createLayer("CERTIFICATE OF EXCELLENCE", 50, 18, "Outfit", 4.8, "bold", "#1E3A8A"),
        this.createLayer("THIS RECOGNITION IS EARNED BY", 50, 31, "Montserrat", 1.3, "normal", "#64748B", 4),
        this.createLayer("Harrison Ford", 50, 42, "Outfit", 6.2, "bold", "#7C3AED"),
        this.createLayer("In appreciation of your technical leadership and innovation in building enterprise platforms.", 50, 56, "Montserrat", 1.6, "normal", "#334155"),
        this.createLayer("ISSUED DATE", 30, 78, "Montserrat", 1.0, "bold", "#94A3B8"),
        this.createLayer("06/12/2026", 30, 73, "Outfit", 1.6, "normal", "#0F172A"),
        this.createLayer("DIRECTOR GENERAL", 70, 78, "Montserrat", 1.0, "bold", "#94A3B8"),
        this.createLayer("Muhammed Ashad K", 70, 71, "Alex Brush", 2.6, "normal", "#7C3AED")
      ];
    } else if (key === "creative") {
      canvasEditor.layers = [
        this.createLayer("creative achievement", 50, 20, "Inter", 4.5, "bold", "#0F172A"),
        this.createLayer("congratulations to", 50, 32, "Montserrat", 1.2, "normal", "#94A3B8", 3),
        this.createLayer("Sarah Connor", 50, 44, "Inter", 6.0, "bold", "#EC4899"),
        this.createLayer("Whose designs pushed the limit of modern UI architecture and visual expression.", 50, 58, "Montserrat", 1.6, "normal", "#475569"),
        this.createLayer("06.12.2026", 50, 75, "Inter", 1.4, "bold", "#EC4899")
      ];
      // Make gradient text on title
      canvasEditor.layers[2].colorType = "gradient";
      canvasEditor.layers[2].gradientStart = "#EC4899";
      canvasEditor.layers[2].gradientEnd = "#7C3AED";
    } else if (key === "luxury") {
      canvasEditor.layers = [
        this.createLayer("DIPLOMA OF MERIT", 50, 18, "Cinzel", 5.2, "bold", "#D4AF37"),
        this.createLayer("BY DECREE OF THE SUPREME ACADEMIC BOARD", 50, 30, "Montserrat", 1.2, "normal", "#94A3B8", 4),
        this.createLayer("ALEX MERCER", 50, 41, "Cinzel", 5.4, "bold", "#F8FAFC"),
        this.createLayer("Conferred for extraordinary service and strategic diplomatic achievements.", 50, 53, "Playfair Display", 1.8, "normal", "#E2E8F0"),
        this.createLayer("BOARD CHAIR", 30, 78, "Montserrat", 1.0, "bold", "#64748B"),
        this.createLayer("Muhammed Ashad K", 30, 72, "Alex Brush", 2.4, "normal", "#D4AF37")
      ];
    }

    canvasEditor.selectedLayer = null;
    canvasEditor.drawTextfromInputs();
    canvasEditor.updateLayersSidebar();
    canvasEditor.updateInspectorPanel();
    
    // Fit zoom
    canvasEditor.fitToScreen();
    
    historyManager.pushState(canvasEditor.layers);
    themeManager.pushNotification(`Loaded ${t.title} Template`, "success");
  }

  createLayer(text, x, y, font, size, weight = "normal", color = "#0F172A", spacing = 0) {
    return {
      id: "layer_" + Math.random().toString(36).substr(2, 9),
      text: text,
      x: x,
      y: y,
      fontFamily: font,
      fontSize: size,
      fontWeight: weight,
      fontStyle: "normal",
      colorType: "solid",
      color: color,
      gradientStart: "#2563EB",
      gradientEnd: "#7C3AED",
      gradientAngle: 90,
      opacity: 100,
      letterSpacing: spacing,
      lineHeight: 1.2,
      shadowEnabled: false,
      shadowX: 2,
      shadowY: 2,
      shadowBlur: 4,
      shadowColor: "#000000",
      strokeEnabled: false,
      strokeWidth: 1,
      strokeColor: "#000000",
      glowEnabled: false,
      glowRadius: 10,
      glowColor: "#06B6D4",
      locked: false,
      visible: true,
      editable: true
    };
  }
}

/* --- Bulk processing asynchronous queue --- */
class BulkProcessor {
  constructor() {
    this.sheetData = [];
    this.titles = [];
    this.mappings = {}; // { layerId: colIndex }
    this.currentIndex = 0;
    this.status = "idle"; // running, paused, cancelled, idle
    
    this.successCount = 0;
    this.failedCount = 0;
    this.errors = [];
    
    this.startTime = null;
    this.elapsedTime = 0;
    this.zip = null;
    this.failedRecordsToReprocess = [];

    this.bindEvents();
  }

  bindEvents() {
    const csvZone = document.getElementById("csv-drag-zone");
    csvZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      csvZone.classList.add("dragover");
    });
    csvZone.addEventListener("dragleave", () => csvZone.classList.remove("dragover"));
    csvZone.addEventListener("drop", (e) => {
      e.preventDefault();
      csvZone.classList.remove("dragover");
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        this.parseDataFile(e.dataTransfer.files[0]);
      }
    });

    document.getElementById("uploadcsv").addEventListener("change", (e) => {
      if (e.target.files && e.target.files[0]) {
        this.parseDataFile(e.target.files[0]);
      }
    });

    document.getElementById("btn-remove-data").addEventListener("click", () => {
      this.resetDataState();
    });

    // Control buttons inside bulk modal dashboard
    document.getElementById("btn-bulk-start").addEventListener("click", () => this.startRun());
    document.getElementById("btn-bulk-pause").addEventListener("click", () => this.pauseRun());
    document.getElementById("btn-bulk-resume").addEventListener("click", () => this.resumeRun());
    document.getElementById("btn-bulk-cancel").addEventListener("click", () => this.cancelRun());
    document.getElementById("btn-clear-console").addEventListener("click", () => {
      document.getElementById("bulk-console-output").innerHTML = "";
    });
    document.getElementById("btn-download-error-report").addEventListener("click", () => this.downloadErrorReport());
    document.getElementById("btn-reprocess-failed").addEventListener("click", () => this.reprocessFailed());
  }

  parseDataFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: "binary", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Parse sheet to rows with automatic date formatting
        const rawRows = parseSheetToRows(worksheet);
        
        if (rawRows.length < 2) {
          throw new Error("The file must contain a header row and at least one data record.");
        }

        this.titles = rawRows[0];
        this.sheetData = rawRows.slice(1).filter(row => row.length > 0);
        
        // Setup initial mappings (intelligent auto-detect)
        this.mappings = {};
        canvasEditor.layers.forEach((layer) => {
          // Find matching index if name is close
          const idx = this.titles.findIndex(t => t.toString().toLowerCase().trim() === layer.text.toLowerCase().trim());
          this.mappings[layer.id] = idx !== -1 ? idx : -1;
        });

        // Toggle UI panels
        document.getElementById("csv-drag-zone").style.display = "none";
        document.getElementById("data-source-status").style.display = "flex";
        document.getElementById("data-file-name").innerText = `${file.name} (${this.sheetData.length} records)`;
        document.getElementById("data-mapping-section").style.display = "block";

        this.renderMappingFields();
        this.renderPreviewTable();
        
        themeManager.pushNotification("Data spreadsheet parsed successfully", "success");
      } catch (err) {
        alert("Failed to parse sheet: " + err.message);
        themeManager.pushNotification("Failed to load sheet file: " + err.message, "error");
      }
    };
    reader.readAsBinaryString(file);
  }

  resetDataState() {
    this.sheetData = [];
    this.titles = [];
    this.mappings = {};
    document.getElementById("csv-drag-zone").style.display = "block";
    document.getElementById("data-source-status").style.display = "none";
    document.getElementById("data-mapping-section").style.display = "none";
    
    // Disable export dashboard elements
    this.disableBulkDownloadWidgets();
  }

  renderMappingFields() {
    const container = document.getElementById("mapping-fields-container");
    container.innerHTML = "";

    canvasEditor.layers.forEach((layer) => {
      const row = document.createElement("div");
      row.className = "mapping-row";
      
      let selectOptions = `<option value="-1">-- Unmapped / Hardcoded --</option>`;
      this.titles.forEach((title, idx) => {
        const selected = this.mappings[layer.id] === idx ? "selected" : "";
        selectOptions += `<option value="${idx}" ${selected}>${title}</option>`;
      });

      row.innerHTML = `
        <span>${layer.text}</span>
        <select id="map-select-${layer.id}">
          ${selectOptions}
        </select>
      `;

      row.querySelector("select").addEventListener("change", (e) => {
        this.mappings[layer.id] = parseInt(e.target.value);
        this.updatePreviewForSelectedRow();
      });

      container.appendChild(row);
    });
  }

  renderPreviewTable() {
    const head = document.getElementById("data-preview-header");
    const body = document.getElementById("data-preview-body");
    
    head.innerHTML = "";
    body.innerHTML = "";

    // Headers
    this.titles.slice(0, 4).forEach((title) => {
      head.innerHTML += `<th>${title}</th>`;
    });

    // Top 5 rows preview
    this.sheetData.slice(0, 5).forEach((row, rIdx) => {
      const tr = document.createElement("tr");
      if (rIdx === 0) tr.className = "selected";

      row.slice(0, 4).forEach((cell) => {
        tr.innerHTML += `<td>${cell || ''}</td>`;
      });

      tr.addEventListener("click", () => {
        body.querySelectorAll("tr").forEach(t => t.classList.remove("selected"));
        tr.classList.add("selected");
        this.previewRow(rIdx);
      });

      body.appendChild(tr);
    });

    // Preview the first record by default
    this.previewRow(0);
  }

  previewRow(idx) {
    if (!this.sheetData[idx]) return;
    const row = this.sheetData[idx];
    
    canvasEditor.layers.forEach((layer) => {
      const colIdx = this.mappings[layer.id];
      if (colIdx !== undefined && colIdx !== null && colIdx !== -1) {
        layer.tempValue = row[colIdx];
      } else {
        delete layer.tempValue;
      }
    });

    canvasEditor.drawTextfromInputs();
  }

  updatePreviewForSelectedRow() {
    const tableBody = document.getElementById("data-preview-body");
    const selectedTr = tableBody.querySelector("tr.selected");
    if (selectedTr) {
      // Find row index
      const trs = Array.from(tableBody.children);
      const idx = trs.indexOf(selectedTr);
      if (idx !== -1) this.previewRow(idx);
    }
  }

  logConsole(message, type = "info") {
    const consoleOut = document.getElementById("bulk-console-output");
    const span = document.createElement("div");
    span.className = `log-line ${type}`;
    span.innerText = `[${new Date().toLocaleTimeString()}] ${message}`;
    consoleOut.appendChild(span);
    consoleOut.scrollTop = consoleOut.scrollHeight;
  }

  startRun() {
    if (this.sheetData.length === 0) return;
    
    // Clear old cache for active run resolution
    exportManager.bgCache = null;

    // Reset Counters
    this.currentIndex = 0;
    this.successCount = 0;
    this.failedCount = 0;
    this.errors = [];
    this.startTime = Date.now();
    this.elapsedTime = 0;
    this.zip = new JSZip();
    this.failedRecordsToReprocess = [];

    this.status = "running";
    this.toggleControlButtons();
    this.disableBulkDownloadWidgets();

    document.getElementById("dashboard-total-records").innerText = this.sheetData.length;
    document.getElementById("dashboard-success-count").innerText = "0";
    document.getElementById("dashboard-failed-count").innerText = "0";
    document.getElementById("dashboard-speed-count").innerText = "0/s";
    
    // Hide failed records view
    document.getElementById("failed-records-container").style.display = "none";
    document.getElementById("btn-download-error-report").disabled = true;

    this.logConsole(`Starting batch execution run for ${this.sheetData.length} records.`, "info");
    
    // Run loop
    this.runNext();
  }

  pauseRun() {
    if (this.status !== "running") return;
    this.status = "paused";
    this.elapsedTime += Date.now() - this.startTime;
    this.toggleControlButtons();
    this.logConsole("Generation paused by user.", "warning");
    themeManager.pushNotification("Generation run paused", "warning");
  }

  resumeRun() {
    if (this.status !== "paused") return;
    this.status = "running";
    this.startTime = Date.now();
    this.toggleControlButtons();
    this.logConsole("Resuming generation run.", "info");
    this.runNext();
  }

  cancelRun() {
    this.status = "cancelled";
    this.toggleControlButtons();
    this.logConsole("Generation run aborted by user.", "error");
    themeManager.pushNotification("Generation run cancelled", "error");
    
    // Clear background cache
    exportManager.bgCache = null;

    // Enable downloading whatever was processed so far
    if (this.successCount > 0) {
      this.enableBulkDownloadWidgets();
    }
  }

  runNext() {
    if (this.status !== "running") return;

    if (this.currentIndex >= this.sheetData.length) {
      this.finishRun();
      return;
    }

    const row = this.sheetData[this.currentIndex];
    const index = this.currentIndex;

    // Apply values
    canvasEditor.layers.forEach((layer) => {
      const colIdx = this.mappings[layer.id];
      if (colIdx !== undefined && colIdx !== null && colIdx !== -1) {
        layer.tempValue = row[colIdx] !== undefined ? row[colIdx].toString() : "";
      } else {
        delete layer.tempValue;
      }
    });

    try {
      // Validation check
      let invalid = false;
      let validationMsg = "";
      canvasEditor.layers.forEach((layer) => {
        if (this.mappings[layer.id] !== -1) {
          const val = layer.tempValue;
          if (val === undefined || val.trim() === "") {
            invalid = true;
            validationMsg = `Missing mapped content value for field: "${layer.text}"`;
          }
        }
      });

      if (invalid) {
        throw new Error(validationMsg);
      }

      // Draw active canvas text layout
      canvasEditor.drawTextfromInputs(true);

      // Render High-res canvas output
      const res = parseInt(document.getElementById("export-resolution").value) || 3;
      const quality = parseFloat(document.getElementById("export-compression").value) || 0.95;
      const format = document.getElementById("downloadtype").value;

      const outputCanvas = exportManager.generateHighResCanvas(res);
      const filename = `Certificate_${index + 1}`;

      if (format === "png" || format === "jpg") {
        const type = format === "jpg" ? "jpeg" : "png";
        const url = outputCanvas.toDataURL(`image/${type}`, quality);
        const data = url.split(",")[1];
        this.zip.file(`${filename}.${format}`, data, { base64: true });
      } else if (format === "pdf") {
        const pdf = exportManager.generatePDFForCanvas(outputCanvas);
        const blob = pdf.output("blob");
        this.zip.file(`${filename}.pdf`, blob);
      }

      this.successCount++;
      document.getElementById("dashboard-success-count").innerText = this.successCount;
      this.logConsole(`[Success] Record ${index + 1} generated.`, "success");
    } catch (err) {
      this.failedCount++;
      document.getElementById("dashboard-failed-count").innerText = this.failedCount;
      
      const failRecord = {
        index: index,
        row: row,
        reason: err.message
      };
      this.failedRecordsToReprocess.push(failRecord);
      
      this.errors.push({
        index: index + 1,
        reason: err.message
      });

      this.logConsole(`[Error] Record ${index + 1} failed: ${err.message}`, "error");
    }

    this.currentIndex++;
    this.updateProgressUI();

    // Async yield (prevents browser lock, renders progress animations smoothly)
    setTimeout(() => {
      requestAnimationFrame(() => this.runNext());
    }, 15);
  }

  finishRun() {
    this.status = "idle";
    this.toggleControlButtons();
    this.logConsole(`Finished run. Generated: ${this.successCount}, Failed: ${this.failedCount}.`, "success");
    themeManager.pushNotification("Bulk generation run completed!", "success");

    // Clear background cache
    exportManager.bgCache = null;

    // Clear temp preview variables
    canvasEditor.layers.forEach(l => delete l.tempValue);
    canvasEditor.drawTextfromInputs();

    // Render failed rows if existing
    if (this.failedCount > 0) {
      this.showFailedRecords();
    }

    // Enable download buttons if success > 0
    if (this.successCount > 0) {
      this.enableBulkDownloadWidgets();
    }
  }

  updateProgressUI() {
    const total = this.sheetData.length;
    const progressVal = (this.currentIndex / total) * 100;
    
    // Update Radial progress
    const circle = document.getElementById("dashboard-radial-progress");
    const pctLabel = document.getElementById("bulk-radial-pct");
    const countLabel = document.getElementById("bulk-current-progress-count");
    
    // Circle stroke dash offset calculation (R = 40, Circumference = 2 * PI * 40 = 251.2)
    const offset = 251.2 - (progressVal / 100) * 251.2;
    circle.style.strokeDashoffset = offset;
    pctLabel.innerText = `${Math.round(progressVal)}%`;
    countLabel.innerText = `${this.currentIndex} / ${total}`;

    // Speed & ETA calculations
    const timeDiff = (this.status === "paused") 
      ? this.elapsedTime 
      : (this.elapsedTime + (Date.now() - this.startTime));
      
    if (timeDiff > 0) {
      const speed = (this.currentIndex / (timeDiff / 1000));
      document.getElementById("dashboard-speed-count").innerText = `${speed.toFixed(1)}/s`;
      
      const remainingCerts = total - this.currentIndex;
      const etaSeconds = speed > 0 ? (remainingCerts / speed) : 0;
      
      const m = Math.floor(etaSeconds / 60);
      const s = Math.round(etaSeconds % 60);
      document.getElementById("bulk-eta-display").innerText = `ETA: ${m}m ${s}s`;
    }
  }

  toggleControlButtons() {
    const start = document.getElementById("btn-bulk-start");
    const pause = document.getElementById("btn-bulk-pause");
    const resume = document.getElementById("btn-bulk-resume");
    const cancel = document.getElementById("btn-bulk-cancel");

    if (this.status === "running") {
      start.style.display = "none";
      pause.style.display = "block";
      resume.style.display = "none";
      cancel.style.display = "block";
    } else if (this.status === "paused") {
      start.style.display = "none";
      pause.style.display = "none";
      resume.style.display = "block";
      cancel.style.display = "block";
    } else { // idle or cancelled
      start.style.display = "block";
      pause.style.display = "none";
      resume.style.display = "none";
      cancel.style.display = "none";
    }
  }

  enableBulkDownloadWidgets() {
    document.getElementById("bulk-btn-zip-png").disabled = false;
    document.getElementById("bulk-btn-zip-jpg").disabled = false;
    document.getElementById("bulk-btn-zip-pdf").disabled = false;
    document.getElementById("bulk-btn-combined-pdf").disabled = false;
  }

  disableBulkDownloadWidgets() {
    document.getElementById("bulk-btn-zip-png").disabled = true;
    document.getElementById("bulk-btn-zip-jpg").disabled = true;
    document.getElementById("bulk-btn-zip-pdf").disabled = true;
    document.getElementById("bulk-btn-combined-pdf").disabled = true;
  }

  showFailedRecords() {
    document.getElementById("failed-records-container").style.display = "block";
    document.getElementById("btn-download-error-report").disabled = false;
    document.getElementById("failed-records-badge").innerText = this.failedCount;

    const list = document.getElementById("failed-records-list");
    list.innerHTML = "";

    this.failedRecordsToReprocess.forEach((fail) => {
      const item = document.createElement("div");
      item.className = "failed-record-item";
      item.innerHTML = `
        <div class="failed-record-info">
          <span class="failed-record-name">Row ${fail.index + 1}: ${fail.row[0] || 'Unnamed Row'}</span>
          <span class="failed-record-reason"><i class="fa-solid fa-triangle-exclamation"></i> ${fail.reason}</span>
        </div>
      `;
      list.appendChild(item);
    });
  }

  downloadErrorReport() {
    if (this.errors.length === 0) return;
    
    let report = `CertifyPro Bulk Run Error Log\nGenerated: ${new Date().toLocaleString()}\n`;
    report += `==============================================\n\n`;
    this.errors.forEach((e) => {
      report += `Record Row [${e.index}]: ${e.reason}\n`;
    });

    const blob = new Blob([report], { type: "text/plain;charset=utf-8" });
    saveAs(blob, "CertifyPro_Error_Report.txt");
  }

  reprocessFailed() {
    if (this.failedRecordsToReprocess.length === 0) return;
    
    this.logConsole(`Reprocessing ${this.failedRecordsToReprocess.length} failed records.`, "info");
    
    // Set active sheetData to only failed rows
    this.sheetData = this.failedRecordsToReprocess.map(f => f.row);
    this.startRun();
  }
}

/* --- Export Manager Module --- */
class ExportManager {
  constructor() {
    this.format = "png";
      this.resolution = 3;
      this.compression = 0.95;
      this.bgCache = null; // Caches background image data URL during bulk runs to optimize speed

      this.bindEvents();
    }

    bindEvents() {
      // Single download trigger
      document.getElementById("btn-download-single").addEventListener("click", () => this.exportSingleCertificate());

      // Bind bulk modal dashboard export formats
      document.getElementById("bulk-btn-zip-png").addEventListener("click", () => this.exportZippedFormat("png"));
      document.getElementById("bulk-btn-zip-jpg").addEventListener("click", () => this.exportZippedFormat("jpg"));
      document.getElementById("bulk-btn-zip-pdf").addEventListener("click", () => this.exportZippedFormat("pdf"));
      document.getElementById("bulk-btn-combined-pdf").addEventListener("click", () => this.exportCombinedPDF());

      // Format dropdown change
      document.getElementById("downloadtype").addEventListener("change", (e) => {
        this.format = e.target.value;
        const qualityWrapper = document.getElementById("jpeg-quality-wrapper");
        if (this.format === "jpg") {
          qualityWrapper.style.display = "block";
        } else {
          qualityWrapper.style.display = "none";
        }
      });

      // Resolution change
      document.getElementById("export-resolution").addEventListener("change", (e) => {
        this.resolution = parseFloat(e.target.value) || 3;
      });

      // Compression change
      document.getElementById("export-compression").addEventListener("change", (e) => {
        this.compression = parseFloat(e.target.value) || 0.95;
      });
    }

    drawLayerOnCanvas(tctx, layer, canvasW, canvasH, scale) {
      tctx.save();

      const defaultFontSizeVal = canvasW / 100;
      const size = layer.fontSize * defaultFontSizeVal;

      tctx.font = `${layer.fontWeight === 'bold' ? 'bold' : ''} ${layer.fontStyle === 'italic' ? 'italic' : ''} ${size}px "${layer.fontFamily}", sans-serif`;
      tctx.globalAlpha = layer.opacity / 100;
      tctx.textAlign = layer.textAlign;
      tctx.textBaseline = "top";

      if ('letterSpacing' in tctx) {
        tctx.letterSpacing = `${(layer.letterSpacing || 0) * scale}px`;
      }

      const xPos = (layer.x / 100) * canvasW;
      const yPos = (layer.y / 100) * canvasH;
      const displayText = layer.tempValue !== undefined ? layer.tempValue : layer.text;
      
      let lines = [];
      let widthPx = 0;
      if (layer.width !== undefined && layer.width !== null) {
        const maxWidth = (layer.width / 100) * canvasW;
        lines = wrapText(tctx, displayText, maxWidth);
        widthPx = maxWidth;
      } else {
        lines = displayText.split("\n");
        let maxLineWidth = 0;
        lines.forEach((line) => {
          const metrics = tctx.measureText(line);
          if (metrics.width > maxLineWidth) {
            maxLineWidth = metrics.width;
          }
        });
        widthPx = maxLineWidth;
      }
      const lineHeights = size * (layer.lineHeight || 1.2);
      const heightPx = lineHeights * lines.length;

      // Shadows
      if (layer.shadowEnabled) {
        tctx.shadowOffsetX = (layer.shadowX || 2) * scale;
        tctx.shadowOffsetY = (layer.shadowY || 2) * scale;
        tctx.shadowBlur = (layer.shadowBlur || 4) * scale;
        tctx.shadowColor = layer.shadowColor || "rgba(0,0,0,0.5)";
      } else if (layer.glowEnabled) {
        tctx.shadowOffsetX = 0;
        tctx.shadowOffsetY = 0;
        tctx.shadowBlur = (layer.glowRadius || 10) * scale;
        tctx.shadowColor = layer.glowColor || "#06B6D4";
      }

      // Colors
      if (layer.colorType === "gradient") {
        const angleRad = ((layer.gradientAngle || 90) * Math.PI) / 180;
        let startX = xPos;
        if (layer.textAlign === "center") startX = xPos - widthPx / 2;
        else if (layer.textAlign === "right") startX = xPos - widthPx;

        const grad = tctx.createLinearGradient(
          startX, 
          yPos, 
          startX + widthPx, 
          yPos + heightPx
        );
        grad.addColorStop(0, layer.gradientStart || "#2563EB");
        grad.addColorStop(1, layer.gradientEnd || "#7C3AED");
        tctx.fillStyle = grad;
      } else {
        tctx.fillStyle = layer.color || "#000000";
      }

      // Draw each line
      lines.forEach((line, index) => {
        const lineY = yPos + index * lineHeights;

        // Stroke
        if (layer.strokeEnabled) {
          tctx.strokeStyle = layer.strokeColor || "#000000";
          tctx.lineWidth = (layer.strokeWidth || 1) * scale;
          tctx.strokeText(line, xPos, lineY);
        }

        tctx.fillText(line, xPos, lineY);
      });

      tctx.restore();
    }

    generateHighResCanvas(scale) {
      const tempCanvas = document.createElement("canvas");
      tempCanvas.width = canvasEditor.canvas.width * scale;
      tempCanvas.height = canvasEditor.canvas.height * scale;
      const tctx = tempCanvas.getContext("2d");

      if (canvasEditor.bgTemplateType === "custom" && canvasEditor.customBgImage.complete) {
        tctx.drawImage(canvasEditor.customBgImage, 0, 0, tempCanvas.width, tempCanvas.height);
      } else {
        renderTemplateBackground(canvasEditor.bgTemplateType, tempCanvas, tctx);
      }

      canvasEditor.layers.forEach((layer) => {
        if (layer.visible) {
          this.drawLayerOnCanvas(tctx, layer, tempCanvas.width, tempCanvas.height, scale);
        }
      });

      return tempCanvas;
    }

    generatePDFForCanvas(canvas) {
      const { jsPDF } = window.jspdf;

      const isPortrait = canvasEditor.canvas.width < canvasEditor.canvas.height;
      const orientation = isPortrait ? "portrait" : "landscape";
      const pageWidth = isPortrait ? 210 : 297;
      const pageHeight = isPortrait ? 297 : 210;

      const pdf = new jsPDF({
        orientation: orientation,
        unit: "mm",
        format: "a4",
        compress: true
      });

      const canvasRatio = canvasEditor.canvas.width / canvasEditor.canvas.height;
      const pageRatio = pageWidth / pageHeight;

      let imgW = pageWidth;
      let imgH = pageHeight;
      if (canvasRatio > pageRatio) {
        imgW = pageWidth;
        imgH = pageWidth / canvasRatio;
      } else {
        imgH = pageHeight;
        imgW = pageHeight * canvasRatio;
      }
      const x = (pageWidth - imgW) / 2;
      const y = (pageHeight - imgH) / 2;

      const quality = this.compression;
      const dataUrl = canvas.toDataURL("image/jpeg", quality);
      pdf.addImage(dataUrl, "JPEG", x, y, imgW, imgH, undefined, "FAST");

      return pdf;
    }

    exportSingleCertificate() {
      canvasEditor.drawTextfromInputs(true);

      const scale = this.resolution;
      const type = this.format;
      const comp = this.compression;
      const title = document.getElementById("project-title").value || "Certificate";

      themeManager.pushNotification(`Compiling single certificate...`);

      if (type === "png" || type === "jpg") {
        const outputCanvas = this.generateHighResCanvas(scale);
        const mime = type === "jpg" ? "image/jpeg" : "image/png";
        const dataUrl = outputCanvas.toDataURL(mime, comp);
        const link = document.createElement("a");
        link.download = `${title}.${type}`;
        link.href = dataUrl;
        link.click();
        themeManager.pushNotification("Export successfully downloaded", "success");
      } else if (type === "pdf") {
        const outputCanvas = this.generateHighResCanvas(scale);
        const pdf = this.generatePDFForCanvas(outputCanvas);
        pdf.save(`${title}.pdf`);
        themeManager.pushNotification("PDF downloaded successfully", "success");
      }

      canvasEditor.drawTextfromInputs();
    }

    exportZippedFormat(formatType) {
      if (!bulkProcessor.zip) return;

      themeManager.pushNotification(`Assembling and packaging ZIP archive...`);

      const activeFormat = document.getElementById("downloadtype").value;
      if (activeFormat !== formatType) {
        alert(`The bulk run was compiled using the ${activeFormat.toUpperCase()} format override. Select ${formatType.toUpperCase()} under overrides and re-run generation.`);
        return;
      }

      bulkProcessor.zip.generateAsync({ type: "blob" }).then((content) => {
        saveAs(content, "CertifyPro_Zipped_Run.zip");
        themeManager.pushNotification("ZIP archive successfully downloaded!", "success");
      });
    }

    exportCombinedPDF() {
      if (bulkProcessor.sheetData.length === 0) return;

      themeManager.pushNotification("Starting Combined multipage PDF compiler...");
      const res = this.resolution;
      const { jsPDF } = window.jspdf;
      const quality = this.compression;

      let pdf = null;

      bulkProcessor.sheetData.forEach((row, idx) => {
        canvasEditor.layers.forEach((layer) => {
          const colIdx = bulkProcessor.mappings[layer.id];
          if (colIdx !== undefined && colIdx !== null && colIdx !== -1) {
            layer.tempValue = row[colIdx] !== undefined ? row[colIdx].toString() : "";
          } else {
            delete layer.tempValue;
          }
        });

        canvasEditor.drawTextfromInputs(true);

        const isPortrait = canvasEditor.canvas.width < canvasEditor.canvas.height;
        const orientation = isPortrait ? "portrait" : "landscape";
        const pageWidth = isPortrait ? 210 : 297;
        const pageHeight = isPortrait ? 297 : 210;
        const canvasRatio = canvasEditor.canvas.width / canvasEditor.canvas.height;
        const pageRatio = pageWidth / pageHeight;

        let imgW = pageWidth;
        let imgH = pageHeight;
        if (canvasRatio > pageRatio) {
          imgW = pageWidth;
          imgH = pageWidth / canvasRatio;
        } else {
          imgH = pageHeight;
          imgW = pageHeight * canvasRatio;
        }
        const x = (pageWidth - imgW) / 2;
        const y = (pageHeight - imgH) / 2;

        if (!pdf) {
          pdf = new jsPDF({
            orientation: orientation,
            unit: "mm",
            format: "a4",
            compress: true
          });
        } else {
          pdf.addPage("a4", orientation);
        }

        const highResCanvas = this.generateHighResCanvas(res);
        const dataUrl = highResCanvas.toDataURL("image/jpeg", quality);
        pdf.addImage(dataUrl, "JPEG", x, y, imgW, imgH, undefined, "FAST");
    });

    pdf.save("CertifyPro_Combined_Run.pdf");
    themeManager.pushNotification("Combined PDF downloaded successfully!", "success");
    canvasEditor.layers.forEach(l => delete l.tempValue);
    canvasEditor.drawTextfromInputs();
  }

  updateEstimates() {
    // Dummy method to prevent errors from other manager references
  }
}

/* --- Global Event Binding and Layout handlers --- */
function setupGlobalBindings() {
  // Top Navbar renaming project titles
  const pTitle = document.getElementById("project-title");
  pTitle.addEventListener("change", () => {
    themeManager.pushNotification(`Project renamed to: ${pTitle.value}`);
    historyManager.autoSave();
  });

  // Save/Export Template logic
  document.getElementById("btn-export-template").addEventListener("click", () => {
    // Clean up temporary run values before exporting layout
    const layersToSave = canvasEditor.layers.map(layer => {
      const cleanLayer = { ...layer };
      delete cleanLayer.tempValue;
      return cleanLayer;
    });

    const templateData = {
      projectTitle: document.getElementById("project-title").value || "Custom Template Design",
      bgTemplateType: canvasEditor.bgTemplateType,
      customBgImageSrc: canvasEditor.customBgImageSrc,
      canvasWidth: canvasEditor.canvas.width,
      canvasHeight: canvasEditor.canvas.height,
      layers: layersToSave
    };

    const blob = new Blob([JSON.stringify(templateData, null, 2)], { type: "application/json" });
    const titleClean = templateData.projectTitle.toLowerCase().replace(/[^a-z0-9]/gi, '_');
    saveAs(blob, `template_${titleClean}.json`);
    themeManager.pushNotification("Custom template saved successfully", "success");
  });

  // Load/Import Template logic
  const fileInput = document.getElementById("input-import-template");
  document.getElementById("btn-import-template-trigger").addEventListener("click", () => {
    fileInput.click();
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = JSON.parse(evt.target.result);
          if (!data || !Array.isArray(data.layers)) {
            throw new Error("Invalid template format: 'layers' array is missing.");
          }

          document.getElementById("project-title").value = data.projectTitle || "Imported Project";
          canvasEditor.bgTemplateType = data.bgTemplateType || "classic";
          
          if (data.canvasWidth && data.canvasHeight) {
            canvasEditor.canvas.width = data.canvasWidth;
            canvasEditor.canvas.height = data.canvasHeight;
          } else {
            // fallback defaults
            if (canvasEditor.bgTemplateType === "black_a4") {
              canvasEditor.canvas.width = 1920;
              canvasEditor.canvas.height = 1358;
            } else if (canvasEditor.bgTemplateType === "black_a4_portrait") {
              canvasEditor.canvas.width = 1358;
              canvasEditor.canvas.height = 1920;
            } else {
              canvasEditor.canvas.width = DEFAULT_WIDTH;
              canvasEditor.canvas.height = DEFAULT_HEIGHT;
            }
          }

          canvasEditor.layers = data.layers;

          if (canvasEditor.bgTemplateType === "custom" && data.customBgImageSrc) {
            canvasEditor.customBgImageSrc = data.customBgImageSrc;
            canvasEditor.customBgImage.src = data.customBgImageSrc;
            canvasEditor.customBgImage.onload = () => {
              canvasEditor.drawTextfromInputs();
              canvasEditor.updateLayersSidebar();
              canvasEditor.updateInspectorPanel();
              canvasEditor.fitToScreen();
              historyManager.pushState(canvasEditor.layers);
              themeManager.pushNotification("Custom template design loaded successfully", "success");
            };
          } else {
            canvasEditor.customBgImageSrc = null;
            canvasEditor.drawTextfromInputs();
            canvasEditor.updateLayersSidebar();
            canvasEditor.updateInspectorPanel();
            canvasEditor.fitToScreen();
            historyManager.pushState(canvasEditor.layers);
            themeManager.pushNotification("Custom template design loaded successfully", "success");
          }

        } catch (err) {
          alert("Failed to load template file: " + err.message);
          themeManager.pushNotification("Failed to load template file: " + err.message, "error");
        }
      };
      reader.readAsText(file);
      // Reset input value to allow uploading the same file again
      fileInput.value = "";
    }
  });

  // Global search layers/presets
  document.getElementById("global-search").addEventListener("input", (e) => {
    const val = e.target.value.toLowerCase().trim();
    if (val === "") {
      canvasEditor.layers.forEach(l => l.visible = true);
    } else {
      canvasEditor.layers.forEach(l => {
        l.visible = l.text.toLowerCase().includes(val);
      });
    }
    canvasEditor.drawTextfromInputs();
  });

  // Collapsible sidebar
  const sidebar = document.getElementById("app-sidebar");
  document.getElementById("btn-collapse-sidebar").addEventListener("click", () => {
    sidebar.classList.toggle("collapsed");
    document.getElementById("btn-collapse-sidebar").querySelector("i").classList.toggle("fa-bars");
    document.getElementById("btn-collapse-sidebar").querySelector("i").classList.toggle("fa-chevron-right");
    setTimeout(() => canvasEditor.fitToScreen(), 250); // fit canvas screen size after sidebar finishes folding animation
  });

  // Tab controls binding
  const tabs = document.querySelectorAll(".sidebar-tabs .tab-btn");
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      
      const target = tab.dataset.tab;
      const panes = document.querySelectorAll(".sidebar-content .tab-pane");
      panes.forEach(pane => {
        pane.classList.toggle("active", pane.id === `pane-${target}`);
      });

      // Expand sidebar if it was collapsed
      sidebar.classList.remove("collapsed");
      document.getElementById("btn-collapse-sidebar").querySelector("i").className = "fa-solid fa-bars";
    });
  });

  // Modals toggling
  const bulkModal = document.getElementById("bulk-dashboard-modal");
  document.getElementById("btn-open-bulk-dashboard").addEventListener("click", () => {
    bulkModal.style.display = "flex";
  });
  document.getElementById("btn-close-bulk").addEventListener("click", () => {
    bulkModal.style.display = "none";
  });

  const shortModal = document.getElementById("shortcuts-modal");
  document.getElementById("btn-shortcuts-toggle").addEventListener("click", () => {
    shortModal.style.display = "flex";
  });
  document.getElementById("btn-close-shortcuts").addEventListener("click", () => {
    shortModal.style.display = "none";
  });

  // Notifications bell toggler
  const notifBtn = document.getElementById("notifications-btn");
  const notifMenu = document.getElementById("notifications-menu");
  notifBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    notifMenu.classList.toggle("active");
  });
  
  document.getElementById("clear-notifications").addEventListener("click", () => {
    document.getElementById("notifications-list").innerHTML = `<div class="no-notifications">No new notifications</div>`;
    document.getElementById("notification-count").innerText = "0";
    document.getElementById("notification-count").style.display = "none";
  });

  window.addEventListener("click", () => {
    notifMenu.classList.remove("active");
  });

  // Settings Accessibility JoyStick toggle
  document.getElementById("setting-show-joystick").addEventListener("change", (e) => {
    document.getElementById("joystick-ui-container").style.display = e.target.checked ? "block" : "none";
  });

  // Bind inspector panel inputs back to model state updates
  document.getElementById("prop-text-val").addEventListener("input", (e) => {
    updateSelectedLayerField("text", e.target.value);
  });

  document.getElementById("prop-pos-x").addEventListener("input", (e) => {
    updateSelectedLayerField("x", parseFloat(e.target.value) || 0);
  });

  document.getElementById("prop-pos-y").addEventListener("input", (e) => {
    updateSelectedLayerField("y", parseFloat(e.target.value) || 0);
  });

  document.getElementById("prop-width").addEventListener("input", (e) => {
    const val = e.target.value === "" ? null : parseFloat(e.target.value);
    updateSelectedLayerField("width", val);
  });

  document.getElementById("prop-font-family").addEventListener("change", (e) => {
    updateSelectedLayerField("fontFamily", e.target.value);
  });

  document.getElementById("prop-font-size").addEventListener("input", (e) => {
    updateSelectedLayerField("fontSize", parseFloat(e.target.value) || 1);
  });

  document.getElementById("prop-opacity").addEventListener("input", (e) => {
    updateSelectedLayerField("opacity", parseInt(e.target.value) || 100);
  });

  document.getElementById("prop-letter-spacing").addEventListener("input", (e) => {
    updateSelectedLayerField("letterSpacing", parseInt(e.target.value) || 0);
  });

  document.getElementById("prop-line-height").addEventListener("input", (e) => {
    updateSelectedLayerField("lineHeight", parseFloat(e.target.value) || 1.2);
  });

  // Styles formatting bold/italic
  document.getElementById("prop-btn-bold").addEventListener("click", () => {
    const boldState = canvasEditor.selectedLayer.fontWeight === "bold" ? "normal" : "bold";
    updateSelectedLayerField("fontWeight", boldState);
    document.getElementById("prop-btn-bold").classList.toggle("active", boldState === "bold");
  });

  document.getElementById("prop-btn-italic").addEventListener("click", () => {
    const italicState = canvasEditor.selectedLayer.fontStyle === "italic" ? "normal" : "italic";
    updateSelectedLayerField("fontStyle", italicState);
    document.getElementById("prop-btn-italic").classList.toggle("active", italicState === "italic");
  });

  // Horizontal inspector alignments
  document.getElementById("prop-align-left").addEventListener("click", () => {
    updateSelectedLayerField("textAlign", "left");
    canvasEditor.updateInspectorPanel();
  });
  document.getElementById("prop-align-center").addEventListener("click", () => {
    updateSelectedLayerField("textAlign", "center");
    canvasEditor.updateInspectorPanel();
  });
  document.getElementById("prop-align-right").addEventListener("click", () => {
    updateSelectedLayerField("textAlign", "right");
    canvasEditor.updateInspectorPanel();
  });

  // Align toolbar actions
  document.getElementById("btn-align-left").addEventListener("click", () => {
    updateSelectedLayerField("textAlign", "left");
    canvasEditor.updateInspectorPanel();
  });
  document.getElementById("btn-align-center").addEventListener("click", () => {
    updateSelectedLayerField("textAlign", "center");
    canvasEditor.updateInspectorPanel();
  });
  document.getElementById("btn-align-right").addEventListener("click", () => {
    updateSelectedLayerField("textAlign", "right");
    canvasEditor.updateInspectorPanel();
  });
  document.getElementById("btn-distribute-h").addEventListener("click", () => {
    updateSelectedLayerField("x", 50.0);
    canvasEditor.updateInspectorPanel();
  });
  document.getElementById("btn-distribute-v").addEventListener("click", () => {
    updateSelectedLayerField("y", 50.0);
    canvasEditor.updateInspectorPanel();
  });

  // Solid vs Gradient color inspector toggles
  document.getElementById("color-mode-solid").addEventListener("click", () => {
    updateSelectedLayerField("colorType", "solid");
    canvasEditor.updateInspectorPanel();
  });

  document.getElementById("color-mode-gradient").addEventListener("click", () => {
    updateSelectedLayerField("colorType", "gradient");
    canvasEditor.updateInspectorPanel();
  });

  document.getElementById("prop-color-solid").addEventListener("input", (e) => {
    updateSelectedLayerField("color", e.target.value);
    document.getElementById("prop-color-solid-hex").value = e.target.value;
  });

  document.getElementById("prop-color-solid-hex").addEventListener("input", (e) => {
    const val = e.target.value;
    if (val.match(/^#[0-9A-Fa-f]{6}$/)) {
      updateSelectedLayerField("color", val);
      document.getElementById("prop-color-solid").value = val;
    }
  });

  document.getElementById("prop-gradient-color-1").addEventListener("input", (e) => {
    updateSelectedLayerField("gradientStart", e.target.value);
  });

  document.getElementById("prop-gradient-color-2").addEventListener("input", (e) => {
    updateSelectedLayerField("gradientEnd", e.target.value);
  });

  document.getElementById("prop-gradient-angle").addEventListener("input", (e) => {
    const deg = parseInt(e.target.value) || 0;
    updateSelectedLayerField("gradientAngle", deg);
    document.getElementById("gradient-angle-val").innerText = `${deg}°`;
  });

  // Shadows
  document.getElementById("effect-shadow-enabled").addEventListener("change", (e) => {
    updateSelectedLayerField("shadowEnabled", e.target.checked);
    document.getElementById("shadow-params").style.display = e.target.checked ? "flex" : "none";
  });
  document.getElementById("prop-shadow-x").addEventListener("input", (e) => {
    updateSelectedLayerField("shadowX", parseInt(e.target.value) || 0);
  });
  document.getElementById("prop-shadow-y").addEventListener("input", (e) => {
    updateSelectedLayerField("shadowY", parseInt(e.target.value) || 0);
  });
  document.getElementById("prop-shadow-blur").addEventListener("input", (e) => {
    updateSelectedLayerField("shadowBlur", parseInt(e.target.value) || 0);
  });
  document.getElementById("prop-shadow-color").addEventListener("input", (e) => {
    updateSelectedLayerField("shadowColor", e.target.value);
  });

  // Stroke
  document.getElementById("effect-stroke-enabled").addEventListener("change", (e) => {
    updateSelectedLayerField("strokeEnabled", e.target.checked);
    document.getElementById("stroke-params").style.display = e.target.checked ? "block" : "none";
  });
  document.getElementById("prop-stroke-width").addEventListener("input", (e) => {
    updateSelectedLayerField("strokeWidth", parseInt(e.target.value) || 1);
  });
  document.getElementById("prop-stroke-color").addEventListener("input", (e) => {
    updateSelectedLayerField("strokeColor", e.target.value);
  });

  // Glow
  document.getElementById("effect-glow-enabled").addEventListener("change", (e) => {
    updateSelectedLayerField("glowEnabled", e.target.checked);
    document.getElementById("glow-params").style.display = e.target.checked ? "block" : "none";
  });
  document.getElementById("prop-glow-radius").addEventListener("input", (e) => {
    updateSelectedLayerField("glowRadius", parseInt(e.target.value) || 10);
  });
  document.getElementById("prop-glow-color").addEventListener("input", (e) => {
    updateSelectedLayerField("glowColor", e.target.value);
  });

  // Duplicate / delete action panel bindings
  document.getElementById("prop-layer-duplicate").addEventListener("click", () => canvasEditor.duplicateLayer());
  document.getElementById("prop-layer-delete").addEventListener("click", () => canvasEditor.deleteLayer());

  // Bind signatures & seals drag insertions
  const decItems = document.querySelectorAll(".decoration-item");
  decItems.forEach((dec) => {
    dec.addEventListener("click", () => {
      const type = dec.dataset.type;
      if (type === "gold_seal") {
        // We draw gold seals procedurally by adding a custom text layer using standard text stamps
        canvasEditor.addTextLayer("★ PREMIUM QUALITY ★", "body");
        canvasEditor.selectedLayer.fontFamily = "Cinzel";
        canvasEditor.selectedLayer.color = "#D4AF37";
        canvasEditor.selectedLayer.glowEnabled = true;
        canvasEditor.selectedLayer.glowColor = "#F59E0B";
        canvasEditor.selectedLayer.glowRadius = 15;
        canvasEditor.drawTextfromInputs();
        canvasEditor.updateInspectorPanel();
      } else if (type === "signature_line") {
        canvasEditor.addTextLayer("_______________________", "body");
        canvasEditor.selectedLayer.color = "#94A3B8";
        canvasEditor.drawTextfromInputs();
        canvasEditor.updateInspectorPanel();
      }
    });
  });

  // Bind corporate palette kit insertions
  setupBrandKit();

  // Bind global keyboard shortcuts
  setupKeyboardShortcuts();

  // Try parsing autosave data from localstorage
  loadAutosave();
}

function updateSelectedLayerField(fieldName, value) {
  if (!canvasEditor.selectedLayer) return;
  canvasEditor.selectedLayer[fieldName] = value;
  canvasEditor.drawTextfromInputs();
  
  // Update sidebar list title if we edited text
  if (fieldName === "text") {
    canvasEditor.updateLayersSidebar();
  }

  if (exportManager) {
    exportManager.updateEstimates();
  }
}

function setupBrandKit() {
  const container = document.getElementById("brand-palettes");
  const palettes = {
    elegantGold: {
      name: "Elegant Gold & Navy",
      colors: ["#0F172A", "#D4AF37", "#1E293B", "#FCFBF7"]
    },
    techModern: {
      name: "Modern Tech Vibrant",
      colors: ["#2563EB", "#7C3AED", "#06B6D4", "#F8FAFC"]
    },
    natureOrganic: {
      name: "Organic Sage & Rose",
      colors: ["#15803D", "#EC4899", "#854D0E", "#F0FDF4"]
    }
  };

  Object.entries(palettes).forEach(([key, value]) => {
    const item = document.createElement("div");
    item.className = "brand-palette-item";
    
    let colorBlocks = "";
    value.colors.forEach((col) => {
      colorBlocks += `<div class="palette-color" style="background-color: ${col};"></div>`;
    });

    item.innerHTML = `
      <div class="palette-colors">${colorBlocks}</div>
      <span class="palette-name">${value.name}</span>
    `;

    item.addEventListener("click", () => {
      if (canvasEditor.selectedLayer) {
        // Apply primary color to active layer
        updateSelectedLayerField("colorType", "solid");
        updateSelectedLayerField("color", value.colors[1] || value.colors[0]);
        canvasEditor.updateInspectorPanel();
        themeManager.pushNotification(`Applied Brand Palette color: ${value.colors[1]}`);
        historyManager.pushState(canvasEditor.layers);
      }
    });

    container.appendChild(item);
  });

  // Bind Default Brand Typography setups
  const presetTypo = document.querySelectorAll(".preset-typography-btn");
  presetTypo.forEach((btn) => {
    btn.addEventListener("click", () => {
      const theme = btn.dataset.theme;
      if (theme === "classic") {
        canvasEditor.layers.forEach((l) => {
          if (l.fontSize > 3.5) l.fontFamily = "Cinzel";
          else l.fontFamily = "Playfair Display";
        });
      } else if (theme === "corporate") {
        canvasEditor.layers.forEach((l) => {
          if (l.fontSize > 3.5) l.fontFamily = "Outfit";
          else l.fontFamily = "Montserrat";
        });
      } else if (theme === "minimalist") {
        canvasEditor.layers.forEach((l) => {
          if (l.fontSize > 3.5) l.fontFamily = "Inter";
          else l.fontFamily = "Montserrat";
        });
      }
      canvasEditor.drawTextfromInputs();
      canvasEditor.updateInspectorPanel();
      themeManager.pushNotification(`Applied ${theme} typography branding presets`, "success");
      historyManager.pushState(canvasEditor.layers);
    });
  });
}

function setupKeyboardShortcuts() {
  window.addEventListener("keydown", (e) => {
    // Escape key removes selections
    if (e.key === "Escape") {
      canvasEditor.selectedLayer = null;
      canvasEditor.updateInspectorPanel();
      canvasEditor.updateLayersSidebar();
      canvasEditor.drawTextfromInputs();
    }

    // Checking key combinations when textareas are NOT focused
    if (document.activeElement.tagName === "INPUT" || document.activeElement.tagName === "TEXTAREA") {
      return;
    }

    const nudgeSpeed = e.shiftKey ? 1.0 : 0.1;

    // Layer selection nudge bindings
    if (canvasEditor.selectedLayer) {
      if (e.key === "ArrowUp") {
        e.preventDefault();
        canvasEditor.selectedLayer.y = parseFloat(Math.max(0, canvasEditor.selectedLayer.y - nudgeSpeed).toFixed(1));
        canvasEditor.updateInspectorPanel();
        canvasEditor.drawTextfromInputs();
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        canvasEditor.selectedLayer.y = parseFloat(Math.min(100, canvasEditor.selectedLayer.y + nudgeSpeed).toFixed(1));
        canvasEditor.updateInspectorPanel();
        canvasEditor.drawTextfromInputs();
      } else if (e.key === "ArrowLeft") {
        e.preventDefault();
        canvasEditor.selectedLayer.x = parseFloat(Math.max(0, canvasEditor.selectedLayer.x - nudgeSpeed).toFixed(1));
        canvasEditor.updateInspectorPanel();
        canvasEditor.drawTextfromInputs();
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        canvasEditor.selectedLayer.x = parseFloat(Math.min(100, canvasEditor.selectedLayer.x + nudgeSpeed).toFixed(1));
        canvasEditor.updateInspectorPanel();
        canvasEditor.drawTextfromInputs();
      } else if (e.key === "Delete" || e.key === "Backspace") {
        e.preventDefault();
        canvasEditor.deleteLayer();
      } else if (e.ctrlKey && e.key === "d") {
        e.preventDefault();
        canvasEditor.duplicateLayer();
      }
    }

    // Undo / Redo key bindings
    if (e.ctrlKey && e.key === "z") {
      e.preventDefault();
      historyManager.undo();
    } else if (e.ctrlKey && e.key === "y") {
      e.preventDefault();
      historyManager.redo();
    }
  });

  // Setup joystick bindings if Joystick Controller CDN/Vendor loaded
  if (window.JoystickController) {
    let prevX = 0;
    let prevY = 0;
    
    // Joystick nudge loop
    const stickNudge = new JoystickController("stick", 64, 8);
    const loop = () => {
      requestAnimationFrame(loop);
      if (!canvasEditor.selectedLayer) return;

      const jX = stickNudge.value.x;
      const jY = stickNudge.value.y;

      if (jX !== 0 || jY !== 0) {
        if (Math.abs(jX - prevX) > 0.05) {
          prevX = jX;
          canvasEditor.selectedLayer.x = parseFloat(Math.min(100, Math.max(0, canvasEditor.selectedLayer.x + (jX * 0.4))).toFixed(1));
        }
        if (Math.abs(jY - prevY) > 0.05) {
          prevY = jY;
          canvasEditor.selectedLayer.y = parseFloat(Math.min(100, Math.max(0, canvasEditor.selectedLayer.y + (jY * -0.4))).toFixed(1)); // joystick Y inversed
        }
        document.getElementById("prop-pos-x").value = canvasEditor.selectedLayer.x;
        document.getElementById("prop-pos-y").value = canvasEditor.selectedLayer.y;
        canvasEditor.drawTextfromInputs();
      }
    };
    loop();
  }
}

function loadAutosave() {
  try {
    const raw = localStorage.getItem("certify-autosave");
    if (raw) {
      const data = JSON.parse(raw);
      if (data && data.layers) {
        document.getElementById("project-title").value = data.projectTitle || "Untitled Certificate Project";
        canvasEditor.layers = data.layers;
        canvasEditor.bgTemplateType = data.bgTemplateType || "classic";
        if (data.bgTemplateType === "custom" && data.customBgImageSrc) {
          canvasEditor.customBgImageSrc = data.customBgImageSrc;
          canvasEditor.customBgImage.src = data.customBgImageSrc;
          canvasEditor.customBgImage.onload = () => {
            canvasEditor.canvas.width = canvasEditor.customBgImage.width;
            canvasEditor.canvas.height = canvasEditor.customBgImage.height;
            canvasEditor.drawTextfromInputs();
          };
        }
        
        canvasEditor.drawTextfromInputs();
        canvasEditor.updateLayersSidebar();
        canvasEditor.updateInspectorPanel();
        
        themeManager.pushNotification("Autosaved project restored", "info");
      }
    }
  } catch (err) {
    console.error("Autosave load failed: ", err);
  }
}

/* --- Procedural Canvas Background Renders (Print-ready Resolution sharp vectors) --- */
function renderTemplateBackground(type, canvas, ctx) {
  const w = canvas.width;
  const h = canvas.height;

  if (type === "black_a4" || type === "black_a4_portrait") {
    // Fill background with black color
    ctx.fillStyle = "#000000";
    ctx.fillRect(0, 0, w, h);
  } else if (type === "luxury") {
    // Dark luxury theme Slate base
    ctx.fillStyle = "#0F172A";
    ctx.fillRect(0, 0, w, h);

    // Ornate gold double stroke borders
    ctx.strokeStyle = "#D4AF37";
    ctx.lineWidth = w * 0.005;
    ctx.strokeRect(w * 0.03, h * 0.04, w * 0.94, h * 0.92);

    ctx.lineWidth = w * 0.0015;
    ctx.strokeRect(w * 0.04, h * 0.05, w * 0.92, h * 0.90);

    // Corner decorative triangles
    ctx.fillStyle = "rgba(212, 175, 55, 0.12)";
    const cornerOffset = w * 0.04;
    const cornerSize = w * 0.06;

    // Top-left corner flourish
    ctx.beginPath();
    ctx.moveTo(cornerOffset, cornerOffset);
    ctx.lineTo(cornerOffset + cornerSize, cornerOffset);
    ctx.lineTo(cornerOffset, cornerOffset + cornerSize);
    ctx.closePath();
    ctx.fill();

    // Top-right corner
    ctx.beginPath();
    ctx.moveTo(w - cornerOffset, cornerOffset);
    ctx.lineTo(w - cornerOffset - cornerSize, cornerOffset);
    ctx.lineTo(w - cornerOffset, cornerOffset + cornerSize);
    ctx.closePath();
    ctx.fill();

    // Bottom-left corner
    ctx.beginPath();
    ctx.moveTo(cornerOffset, h - cornerOffset);
    ctx.lineTo(cornerOffset + cornerSize, h - cornerOffset);
    ctx.lineTo(cornerOffset, h - cornerOffset - cornerSize);
    ctx.closePath();
    ctx.fill();

    // Bottom-right corner
    ctx.beginPath();
    ctx.moveTo(w - cornerOffset, h - cornerOffset);
    ctx.lineTo(w - cornerOffset - cornerSize, h - cornerOffset);
    ctx.lineTo(w - cornerOffset, h - cornerOffset - cornerSize);
    ctx.closePath();
    ctx.fill();

    // Procedural seal graphic at bottom right
    drawProceduralGoldSeal(ctx, w * 0.82, h * 0.76, w * 0.05);

  } else if (type === "corporate") {
    // Plain white base
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, w, h);

    // Blue geometric top left ribbon
    ctx.fillStyle = "#1E3A8A";
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(w * 0.22, 0);
    ctx.lineTo(0, h * 0.38);
    ctx.closePath();
    ctx.fill();

    // Cyan border line on triangle ribbon
    ctx.strokeStyle = "#06B6D4";
    ctx.lineWidth = w * 0.006;
    ctx.beginPath();
    ctx.moveTo(w * 0.22, 0);
    ctx.lineTo(0, h * 0.38);
    ctx.stroke();

    // Violet bottom-right ribbon
    ctx.fillStyle = "#7C3AED";
    ctx.beginPath();
    ctx.moveTo(w, h);
    ctx.lineTo(w * 0.78, h);
    ctx.lineTo(w, h * 0.62);
    ctx.closePath();
    ctx.fill();

    // Blue border line on violet ribbon
    ctx.strokeStyle = "#2563EB";
    ctx.lineWidth = w * 0.006;
    ctx.beginPath();
    ctx.moveTo(w * 0.78, h);
    ctx.lineTo(w, h * 0.62);
    ctx.stroke();

    // Thin slate border outline box
    ctx.strokeStyle = "#E2E8F0";
    ctx.lineWidth = w * 0.004;
    ctx.strokeRect(w * 0.02, h * 0.035, w * 0.96, h * 0.93);

  } else if (type === "creative") {
    // Light silver/grey base
    ctx.fillStyle = "#F8FAFC";
    ctx.fillRect(0, 0, w, h);

    // Modern linear gradient outer border frame
    const grad = ctx.createLinearGradient(0, 0, w, h);
    grad.addColorStop(0, "#FF8A00");
    grad.addColorStop(0.5, "#EC4899");
    grad.addColorStop(1, "#7C3AED");

    ctx.strokeStyle = grad;
    ctx.lineWidth = w * 0.015;
    ctx.strokeRect(w * 0.01, h * 0.018, w * 0.98, h * 0.964);

    // Inner minimal border
    ctx.strokeStyle = "rgba(15, 23, 42, 0.04)";
    ctx.lineWidth = w * 0.001;
    ctx.strokeRect(w * 0.03, h * 0.048, w * 0.94, h * 0.904);

    // Modern layout guidelines (crosshairs)
    ctx.strokeStyle = "rgba(15, 23, 42, 0.08)";
    ctx.lineWidth = 1;

    const crosshairs = [
      [w * 0.03, h * 0.048],
      [w * 0.97, h * 0.048],
      [w * 0.03, h * 0.952],
      [w * 0.97, h * 0.952]
    ];

    crosshairs.forEach(([cx, cy]) => {
      ctx.beginPath();
      ctx.moveTo(cx - 15, cy);
      ctx.lineTo(cx + 15, cy);
      ctx.stroke();

      ctx.beginPath();
      ctx.moveTo(cx, cy - 15);
      ctx.lineTo(cx, cy + 15);
      ctx.stroke();
    });

  } else {
    // Default: Classic Academy style
    ctx.fillStyle = "#FAF8F5";
    ctx.fillRect(0, 0, w, h);

    // Thick primary border box
    ctx.strokeStyle = "#1E293B";
    ctx.lineWidth = w * 0.009;
    ctx.strokeRect(w * 0.02, h * 0.03, w * 0.96, h * 0.94);

    // Thin gold inner double border
    ctx.strokeStyle = "#D4AF37";
    ctx.lineWidth = w * 0.0025;
    ctx.strokeRect(w * 0.03, h * 0.045, w * 0.94, h * 0.91);

    // Draw traditional corners
    const size = w * 0.02;
    drawFlourish(ctx, w * 0.03, h * 0.045, size, 1);
    drawFlourish(ctx, w * 0.97, h * 0.045, size, 2);
    drawFlourish(ctx, w * 0.03, h * 0.955, size, 3);
    drawFlourish(ctx, w * 0.97, h * 0.955, size, 4);
  }
}

function drawProceduralGoldSeal(ctx, cx, cy, r) {
  ctx.save();
  
  // Seal shadow
  ctx.shadowColor = "rgba(0, 0, 0, 0.25)";
  ctx.shadowBlur = 10;
  ctx.shadowOffsetX = 2;
  ctx.shadowOffsetY = 4;

  // Outer spikes scalloped path
  ctx.fillStyle = "#D97706";
  ctx.beginPath();
  const spikes = 36;
  const step = Math.PI / spikes;
  let rot = (Math.PI / 2) * 3;
  let x = cx;
  let y = cy;

  ctx.moveTo(cx, cy - r);
  for (let i = 0; i < spikes; i++) {
    x = cx + Math.cos(rot) * r;
    y = cy + Math.sin(rot) * r;
    ctx.lineTo(x, y);
    rot += step;

    x = cx + Math.cos(rot) * (r * 0.92);
    y = cy + Math.sin(rot) * (r * 0.92);
    ctx.lineTo(x, y);
    rot += step;
  }
  ctx.closePath();
  ctx.fill();

  // Reset shadows for inner graphics
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;

  // Bright gold base circle
  ctx.fillStyle = "#F59E0B";
  ctx.beginPath();
  ctx.arc(cx, cy, r * 0.82, 0, Math.PI * 2);
  ctx.fill();

  // Dark gold border outlines
  ctx.strokeStyle = "#D4AF37";
  ctx.lineWidth = r * 0.04;
  ctx.beginPath();
  ctx.arc(cx, cy, r * 0.74, 0, Math.PI * 2);
  ctx.stroke();

  // Decorative Inner star
  ctx.fillStyle = "#B45309";
  drawStar(ctx, cx, cy, 5, r * 0.32, r * 0.14);

  // Decorative circular ribbon texts lines
  ctx.strokeStyle = "rgba(180, 83, 9, 0.25)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.arc(cx, cy, r * 0.52, 0, Math.PI * 2);
  ctx.stroke();

  ctx.restore();
}

function drawStar(ctx, cx, cy, spikes, outerRadius, innerRadius) {
  let rot = (Math.PI / 2) * 3;
  let x = cx;
  let y = cy;
  const step = Math.PI / spikes;

  ctx.beginPath();
  ctx.moveTo(cx, cy - outerRadius);
  for (let i = 0; i < spikes; i++) {
    x = cx + Math.cos(rot) * outerRadius;
    y = cy + Math.sin(rot) * outerRadius;
    ctx.lineTo(x, y);
    rot += step;

    x = cx + Math.cos(rot) * innerRadius;
    y = cy + Math.sin(rot) * innerRadius;
    ctx.lineTo(x, y);
    rot += step;
  }
  ctx.lineTo(cx, cy - outerRadius);
  ctx.closePath();
  ctx.fill();
}

function drawFlourish(ctx, x, y, size, corner) {
  ctx.strokeStyle = "#D4AF37";
  ctx.lineWidth = size * 0.08;
  
  ctx.save();
  ctx.beginPath();
  if (corner === 1) { // top left
    ctx.moveTo(x + size, y);
    ctx.lineTo(x, y);
    ctx.lineTo(x, y + size);
    ctx.moveTo(x, y);
    ctx.lineTo(x + size * 0.6, y + size * 0.6);
  } else if (corner === 2) { // top right
    ctx.moveTo(x - size, y);
    ctx.lineTo(x, y);
    ctx.lineTo(x, y + size);
    ctx.moveTo(x, y);
    ctx.lineTo(x - size * 0.6, y + size * 0.6);
  } else if (corner === 3) { // bottom left
    ctx.moveTo(x + size, y);
    ctx.lineTo(x, y);
    ctx.lineTo(x, y - size);
    ctx.moveTo(x, y);
    ctx.lineTo(x + size * 0.6, y - size * 0.6);
  } else if (corner === 4) { // bottom right
    ctx.moveTo(x - size, y);
    ctx.lineTo(x, y);
    ctx.lineTo(x, y - size);
    ctx.moveTo(x, y);
    ctx.lineTo(x - size * 0.6, y - size * 0.6);
  }
  ctx.stroke();
  ctx.restore();
}
