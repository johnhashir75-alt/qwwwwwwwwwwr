// Set PDF.js worker path
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.11.338/pdf.worker.min.js';
const { jsPDF } = window.jspdf;

// DOM Elements (same as before)
const toolCards = document.querySelectorAll('.tool-card:not(.planned-tool)');
const fileInputs = document.querySelectorAll('.file-input');
const processingModal = document.getElementById('processing-modal');
const downloadModal = document.getElementById('download-modal');
const progressBar = document.getElementById('progress-bar');
const modalTitle = document.getElementById('modal-title');
const modalMessage = document.getElementById('modal-message');
const downloadTitle = document.getElementById('download-title');
const downloadMessage = document.getElementById('download-message');
const downloadBtn = document.getElementById('download-btn');
const cancelBtn = document.getElementById('cancel-btn');
const newFileBtn = document.getElementById('new-file-btn');

// Global variables
let currentTool = null;
let blob = null;
let processingInterval = null;
let abortController = new AbortController();

// Event listeners (same as before)
toolCards.forEach((card, index) => {
    const toolBtn = card.querySelector('.tool-btn');
    const fileInput = card.querySelector('.file-input');
    
    toolBtn.addEventListener('click', () => {
        currentTool = card.id;
        fileInput.click();
    });
});

fileInputs.forEach(input => {
    input.addEventListener('change', handleFileSelect);
});

cancelBtn.addEventListener('click', cancelProcessing);
newFileBtn.addEventListener('click', resetTool);
downloadBtn.addEventListener('click', downloadFile);

// File selection handler
async function handleFileSelect(e) {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    
    // Check if merge tool and multiple files are selected
    if (currentTool === 'merge-pdf' && files.length < 2) {
        alert('Please select at least 2 PDF files to merge.');
        return;
    }
    
    showProcessingModal();
    abortController = new AbortController();
    
    try {
        // Actual processing based on tool
        switch(currentTool) {
            case 'pdf-to-word':
                blob = await convertPdfToWord(files[0]);
                break;
            case 'pdf-to-excel':
                blob = await convertPdfToExcel(files[0]);
                break;
            case 'pdf-to-jpg':
                blob = await convertPdfToImages(files[0], 'jpg');
                break;
            case 'pdf-to-png':
                blob = await convertPdfToImages(files[0], 'png');
                break;
            case 'pdf-to-text':
                blob = await extractPdfText(files[0]);
                break;
            case 'jpg-to-pdf':
            case 'scan-to-pdf':
                blob = await imagesToPdf(Array.from(files));
                break;
            case 'rotate-pdf':
                blob = await rotatePdf(files[0]);
                break;
            case 'watermark-pdf':
                blob = await watermarkPdf(files[0]);
                break;
            case 'page-numbers':
                blob = await addPageNumbers(files[0]);
                break;
            case 'organize-pdf':
                blob = await organizePdf(files[0]);
                break;
            case 'pdf-to-markdown':
                blob = await pdfToMarkdown(files[0]);
                break;
            case 'merge-pdf':
                blob = await mergePdfFiles(Array.from(files));
                break;
            case 'split-pdf':
                // For split, we'll handle differently as it produces multiple files
                await splitPdfFile(files[0]);
                return;
            case 'compress-pdf':
                blob = await compressPdfFile(files[0]);
                break;
        }
        
        processingComplete();
    } catch (error) {
        if (error.name !== 'AbortError') {
            console.error('Conversion error:', error);
            alert(`Conversion failed: ${error.message}`);
            cancelProcessing();
        }
    }
}

// ====================================
// ACTUAL CONVERSION IMPLEMENTATIONS
// ====================================

async function convertPdfToWord(file) {
 updateProgress(10, 'Reading PDF text...');
 const pdf = await pdfjsLib.getDocument(await readFileAsArrayBuffer(file)).promise;
 const paragraphs=[];
 for(let i=1;i<=pdf.numPages;i++) {
  if(abortController.signal.aborted) throw new Error('Conversion cancelled');
  const page=await pdf.getPage(i); const content=await page.getTextContent();
  paragraphs.push(new docx.Paragraph({text:content.items.map(x=>x.str).join(' '),pageBreakBefore:i>1}));
  updateProgress(10+i/pdf.numPages*80,'Extracting page '+i);
 }
 return docx.Packer.toBlob(new docx.Document({sections:[{children:paragraphs}]}));
}

async function convertPdfToExcel(pdfFile) {
    updateProgress(10, 'Loading PDF document...');
    const pdf = await pdfjsLib.getDocument(await readFileAsArrayBuffer(pdfFile)).promise;
    
    updateProgress(30, 'Extracting tables...');
    const workbook = XLSX.utils.book_new();
    let worksheet;
    let tableData = [];
    
    for (let i = 1; i <= pdf.numPages; i++) {
        if (abortController.signal.aborted) throw new Error('Conversion cancelled');
        
        updateProgress(30 + (i / pdf.numPages * 50), `Processing page ${i} of ${pdf.numPages}...`);
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        // Simple table extraction - this could be enhanced with more sophisticated parsing
        const lines = textContent.items.reduce((acc, item) => {
            const lineIndex = Math.floor(item.transform[5] / 20); // Simple line grouping
            acc[lineIndex] = (acc[lineIndex] || '') + item.str + '\t';
            return acc;
        }, {});
        
        const pageTables = Object.values(lines).map(line => line.split('\t').filter(cell => cell.trim()));
        tableData = tableData.concat(pageTables);
    }
    
    updateProgress(90, 'Generating Excel file...');
    worksheet = XLSX.utils.aoa_to_sheet(tableData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

async function convertPdfToImages(pdfFile, format) {
    updateProgress(10, 'Loading PDF document...');
    const pdf = await pdfjsLib.getDocument(await readFileAsArrayBuffer(pdfFile)).promise;
    
    updateProgress(20, 'Preparing images...');
    const zip = new JSZip();
    const imgFolder = zip.folder("images");
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const dpi = 150;
    
    for (let i = 1; i <= pdf.numPages; i++) {
        if (abortController.signal.aborted) throw new Error('Conversion cancelled');
        
        updateProgress(20 + (i / pdf.numPages * 70), `Converting page ${i} of ${pdf.numPages}...`);
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: dpi / 72 });
        
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        
        await page.render({
            canvasContext: ctx,
            viewport: viewport
        }).promise;
        
        const imageData = canvas.toDataURL(`image/${format}`);
        const base64Data = imageData.split(',')[1];
        imgFolder.file(`page_${i}.${format}`, base64Data, { base64: true });
    }
    
    updateProgress(95, 'Creating download package...');
    const content = await zip.generateAsync({ type: 'blob' });
    return content;
}

async function extractPdfText(pdfFile) {
    updateProgress(10, 'Loading PDF document...');
    const pdf = await pdfjsLib.getDocument(await readFileAsArrayBuffer(pdfFile)).promise;
    
    updateProgress(30, 'Extracting text...');
    let fullText = '';
    
    for (let i = 1; i <= pdf.numPages; i++) {
        if (abortController.signal.aborted) throw new Error('Conversion cancelled');
        
        updateProgress(30 + (i / pdf.numPages * 60), `Processing page ${i} of ${pdf.numPages}...`);
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const text = textContent.items.map(item => item.str).join(' ');
        fullText += `=== Page ${i} ===\n${text}\n\n`;
    }
    
    updateProgress(95, 'Preparing text file...');
    return new Blob([fullText], { type: 'text/plain' });
}

async function mergePdfFiles(pdfFiles) {
    updateProgress(5, 'Initializing merger...');
    const { PDFDocument } = PDFLib;
    const mergedPdf = await PDFDocument.create();
    
    for (let i = 0; i < pdfFiles.length; i++) {
        if (abortController.signal.aborted) throw new Error('Conversion cancelled');
        
        updateProgress(5 + (i / pdfFiles.length * 90), `Merging file ${i + 1} of ${pdfFiles.length}...`);
        const pdfBytes = await readFileAsArrayBuffer(pdfFiles[i]);
        const pdfDoc = await PDFDocument.load(pdfBytes);
        
        const pages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
        pages.forEach(page => mergedPdf.addPage(page));
    }
    
    updateProgress(98, 'Finalizing merged document...');
    const mergedPdfBytes = await mergedPdf.save();
    return new Blob([mergedPdfBytes], { type: 'application/pdf' });
}

async function splitPdfFile(pdfFile) {
    updateProgress(10, 'Loading PDF document...');
    const { PDFDocument } = PDFLib;
    const pdfBytes = await readFileAsArrayBuffer(pdfFile);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const pageCount = pdfDoc.getPageCount();
    
    updateProgress(30, 'Creating individual pages...');
    const zip = new JSZip();
    
    for (let i = 0; i < pageCount; i++) {
        if (abortController.signal.aborted) throw new Error('Conversion cancelled');
        
        updateProgress(30 + (i / pageCount * 60), `Processing page ${i + 1} of ${pageCount}...`);
        const newPdf = await PDFDocument.create();
        const [page] = await newPdf.copyPages(pdfDoc, [i]);
        newPdf.addPage(page);
        
        const pageBytes = await newPdf.save();
        zip.file(`page_${i + 1}.pdf`, pageBytes);
    }
    
    updateProgress(95, 'Creating download package...');
    const content = await zip.generateAsync({ type: 'blob' });
    
    // For split PDF, we directly download the zip since it's multiple files
    processingModal.style.display = 'none';
    saveAs(content, 'split-pages.zip');
    resetFileInputs();
}

async function compressPdfFile(file) {
 updateProgress(20,'Optimising PDF structure...');
 const doc=await PDFLib.PDFDocument.load(await readFileAsArrayBuffer(file));
 const bytes=await doc.save({useObjectStreams:true});
 updateProgress(95,'Preparing optimised PDF...');
 return new Blob([bytes],{type:'application/pdf'});
}

async function imagesToPdf(files) {
    if (!files.length) throw new Error('Please select at least one image.');
    updateProgress(15, 'Preparing images...');
    const pdfDoc = await PDFLib.PDFDocument.create();
    for (let i=0;i<files.length;i++) {
        const bytes = await readFileAsArrayBuffer(files[i]);
        const type=(files[i].type||'').toLowerCase();
        const img = type.includes('png') ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes);
        const maxW=595.28, maxH=841.89, margin=28;
        const scale=Math.min((maxW-margin*2)/img.width,(maxH-margin*2)/img.height,1);
        const w=img.width*scale,h=img.height*scale;
        const page=pdfDoc.addPage([maxW,maxH]);
        page.drawImage(img,{x:(maxW-w)/2,y:(maxH-h)/2,width:w,height:h});
        updateProgress(15+Math.round((i+1)/files.length*70),`Adding image ${i+1} of ${files.length}...`);
    }
    const out=await pdfDoc.save(); updateProgress(100,'PDF ready');
    return new Blob([out],{type:'application/pdf'});
}

async function rotatePdf(file) {
    const angleRaw=prompt('Rotate every page by 90, 180, or 270 degrees:','90');
    if(angleRaw===null) throw new DOMException('Cancelled','AbortError');
    const angle=[90,180,270].includes(Number(angleRaw))?Number(angleRaw):90;
    updateProgress(25,'Loading PDF...');
    const doc=await PDFLib.PDFDocument.load(await readFileAsArrayBuffer(file));
    doc.getPages().forEach(p=>{ const current=p.getRotation().angle||0; p.setRotation(PDFLib.degrees((current+angle)%360)); });
    updateProgress(80,'Rotating pages...');
    return new Blob([await doc.save()],{type:'application/pdf'});
}

async function watermarkPdf(file) {
    const text=prompt('Watermark text:','PDFMivo');
    if(text===null) throw new DOMException('Cancelled','AbortError');
    updateProgress(25,'Loading PDF...');
    const doc=await PDFLib.PDFDocument.load(await readFileAsArrayBuffer(file));
    const font=await doc.embedFont(PDFLib.StandardFonts.HelveticaBold);
    doc.getPages().forEach(page=>{const {width,height}=page.getSize(); const size=Math.max(18,Math.min(48,width/10)); const tw=font.widthOfTextAtSize(text,size); page.drawText(text,{x:(width-tw)/2,y:height/2,size,font,color:PDFLib.rgb(.55,.55,.6),opacity:.22,rotate:PDFLib.degrees(35)});});
    updateProgress(85,'Applying watermark...');
    return new Blob([await doc.save()],{type:'application/pdf'});
}

async function addPageNumbers(file) {
    updateProgress(25,'Loading PDF...');
    const doc=await PDFLib.PDFDocument.load(await readFileAsArrayBuffer(file));
    const font=await doc.embedFont(PDFLib.StandardFonts.Helvetica);
    const pages=doc.getPages();
    pages.forEach((page,i)=>{const {width}=page.getSize(); const text=`${i+1} / ${pages.length}`; const size=10; page.drawText(text,{x:(width-font.widthOfTextAtSize(text,size))/2,y:18,size,font,color:PDFLib.rgb(.25,.3,.4)});});
    updateProgress(85,'Adding page numbers...');
    return new Blob([await doc.save()],{type:'application/pdf'});
}

async function organizePdf(file) {
    updateProgress(20,'Loading PDF...');
    const src=await PDFLib.PDFDocument.load(await readFileAsArrayBuffer(file));
    const n=src.getPageCount();
    const raw=prompt(`Enter page order using numbers 1-${n}, separated by commas. Example: 3,1,2\nUse only the pages you want to keep.`,Array.from({length:n},(_,i)=>i+1).join(','));
    if(raw===null) throw new DOMException('Cancelled','AbortError');
    const order=raw.split(',').map(x=>Number(x.trim())-1).filter(x=>Number.isInteger(x)&&x>=0&&x<n);
    if(!order.length) throw new Error('No valid page numbers entered.');
    const out=await PDFLib.PDFDocument.create();
    const copied=await out.copyPages(src,order); copied.forEach(p=>out.addPage(p));
    updateProgress(90,'Reordering pages...');
    return new Blob([await out.save()],{type:'application/pdf'});
}

async function pdfToMarkdown(file) {
    updateProgress(15,'Reading PDF text...');
    const pdf=await pdfjsLib.getDocument(await readFileAsArrayBuffer(file)).promise;
    let md=`# ${file.name.replace(/\.pdf$/i,'')}\n\n`;
    for(let i=1;i<=pdf.numPages;i++){
        const page=await pdf.getPage(i); const content=await page.getTextContent();
        const text=content.items.map(x=>x.str).join(' ').replace(/\s+/g,' ').trim();
        md+=`## Page ${i}\n\n${text}\n\n`;
        updateProgress(15+Math.round(i/pdf.numPages*75),`Extracting page ${i} of ${pdf.numPages}...`);
    }
    return new Blob([md],{type:'text/markdown;charset=utf-8'});
}

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function updateProgress(percent, message) {
    progressBar.style.width = `${percent}%`;
    if (message) modalMessage.textContent = message;
}

function showProcessingModal() {
    // Set modal content based on tool
    switch(currentTool) {
        // In your downloadFile function, add this case:
case 'pdf-to-word':
            modalTitle.textContent = 'Converting PDF to Word';
            modalMessage.textContent = 'Extracting content from your PDF...';
            break;
        case 'pdf-to-excel':
            modalTitle.textContent = 'Converting PDF to Excel';
            modalMessage.textContent = 'Extracting tables from your PDF to Excel format...';
            break;
        case 'pdf-to-jpg':
            modalTitle.textContent = 'Converting PDF to JPG';
            modalMessage.textContent = 'Converting PDF pages to JPG images...';
            break;
        case 'pdf-to-png':
            modalTitle.textContent = 'Converting PDF to PNG';
            modalMessage.textContent = 'Converting PDF pages to PNG images...';
            break;
        case 'pdf-to-text':
            modalTitle.textContent = 'Extracting Text from PDF';
            modalMessage.textContent = 'Extracting text content from your PDF...';
            break;
        case 'merge-pdf':
            modalTitle.textContent = 'Merging PDF Files';
            modalMessage.textContent = 'Combining your PDF files into one document...';
            break;
        case 'split-pdf':
            modalTitle.textContent = 'Splitting PDF';
            modalMessage.textContent = 'Splitting your PDF into multiple files...';
            break;
        case 'compress-pdf':
            modalTitle.textContent = 'Compressing PDF';
            modalMessage.textContent = 'Reducing your PDF file size...';
            break;
        case 'jpg-to-pdf':
        case 'scan-to-pdf': modalTitle.textContent='Creating PDF'; modalMessage.textContent='Combining your images into a PDF...'; break;
        case 'rotate-pdf': modalTitle.textContent='Rotating PDF'; modalMessage.textContent='Rotating your PDF pages...'; break;
        case 'watermark-pdf': modalTitle.textContent='Adding Watermark'; modalMessage.textContent='Applying your watermark...'; break;
        case 'page-numbers': modalTitle.textContent='Adding Page Numbers'; modalMessage.textContent='Numbering PDF pages...'; break;
        case 'organize-pdf': modalTitle.textContent='Organizing PDF'; modalMessage.textContent='Reordering selected pages...'; break;
        case 'pdf-to-markdown': modalTitle.textContent='Converting to Markdown'; modalMessage.textContent='Extracting PDF text...'; break;
    }
    
    processingModal.style.display = 'flex';
}

function processingComplete() {
    // Update download modal
    downloadTitle.textContent = 'Conversion Complete!';
    downloadMessage.textContent = `Your ${currentTool.replace(/-/g, ' ')} file is ready to download.`;
    
    let fileExtension = '';
    switch(currentTool) {
        case 'pdf-to-word': fileExtension = 'DOCX'; break;
        case 'pdf-to-excel': fileExtension = 'XLSX'; break;
        case 'pdf-to-jpg': fileExtension = 'JPG'; break;
        case 'pdf-to-png': fileExtension = 'PNG'; break;
        case 'pdf-to-text': fileExtension = 'TXT'; break;
        case 'merge-pdf': fileExtension = 'PDF'; break;
        case 'compress-pdf': fileExtension = 'PDF'; break;
        case 'jpg-to-pdf': case 'scan-to-pdf': case 'rotate-pdf': case 'watermark-pdf': case 'page-numbers': case 'organize-pdf': fileExtension = 'PDF'; break;
        case 'pdf-to-markdown': fileExtension = 'MD'; break;
    }
    
    downloadBtn.textContent = `Download ${fileExtension}`;
    
    // Show download modal
    processingModal.style.display = 'none';
    downloadModal.style.display = 'flex';
}

function cancelProcessing() {
    abortController.abort();
    processingModal.style.display = 'none';
    resetFileInputs();
}

function downloadFile() {
    if (!blob) return;
    
    let fileName = 'converted-file';
    switch(currentTool) {
        case 'pdf-to-word': fileName += '.docx'; break;
        case 'pdf-to-excel': fileName += '.xlsx'; break;
        case 'pdf-to-jpg': fileName += '-jpg.zip'; break;
        case 'pdf-to-png': fileName += '-png.zip'; break;
        case 'pdf-to-text': fileName += '.txt'; break;
        case 'merge-pdf': fileName = 'merged-document.pdf'; break;
        case 'compress-pdf': fileName = 'compressed.pdf'; break;
        case 'jpg-to-pdf': fileName = 'images.pdf'; break;
        case 'scan-to-pdf': fileName = 'scanned-document.pdf'; break;
        case 'rotate-pdf': fileName = 'rotated.pdf'; break;
        case 'watermark-pdf': fileName = 'watermarked.pdf'; break;
        case 'page-numbers': fileName = 'numbered.pdf'; break;
        case 'organize-pdf': fileName = 'organized.pdf'; break;
        case 'pdf-to-markdown': fileName = 'converted.md'; break;
    }
    
    saveAs(blob, fileName);
}

function resetTool() {
    downloadModal.style.display = 'none';
    resetFileInputs();
}

function resetFileInputs() {
    fileInputs.forEach(input => {
        input.value = '';
    });
}
