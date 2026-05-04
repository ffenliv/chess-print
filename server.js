const express = require('express');
const multer = require('multer');
const pdf = require('pdf-parse');
const PDFDocument = require('pdfkit');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static('public'));
app.use(express.json());

// Store for generated PDFs
const pdfStore = {};

// Counter file for nodemon restarts
const COUNTER_FILE = path.join(__dirname, '.restart-counter');

function getRestartCounter() {
  try {
    if (fs.existsSync(COUNTER_FILE)) {
      const count = parseInt(fs.readFileSync(COUNTER_FILE, 'utf8').trim());
      return isNaN(count) ? 1 : count + 1;
    }
  } catch (err) {
    console.error('Error reading counter file:', err);
  }
  return 1;
}

function saveRestartCounter(count) {
  try {
    fs.writeFileSync(COUNTER_FILE, count.toString());
  } catch (err) {
    console.error('Error writing counter file:', err);
  }
}

// Helper function to extract text from PDF
async function extractTextFromPDF(buffer) {
  const data = await pdf(buffer);
  return data.text;
}

// Helper function to parse pairings from XLSX
function parsePairingsFromXLSX(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    const pairings = [];
    let headerRowIdx = -1;
    let tournamentName = '';
    let roundInfo = '';
    
    // Tournament name is in row 2 (index 1)
    if (data[1] && data[1][0]) {
      tournamentName = String(data[1][0]).trim();
    }
    
    // Search all text for "Round" followed by a number pattern
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row && row[0]) {
        const fullText = String(row[0]);
        if (/round\s+\d+/i.test(fullText)) {
          roundInfo = fullText.trim();
          break;
        }
      }
    }
    
    console.log('Tournament Name:', tournamentName);
    console.log('Round Info:', roundInfo);
    
    // Find header row - scan for "Bo."
    for (let i = 0; i < Math.min(data.length, 20); i++) {
      const row = data[i];
      if (row && row[0]) {
        const firstCell = String(row[0]).toLowerCase();
        if (firstCell === 'bo.' || firstCell.includes('bo')) {
          headerRowIdx = i;
          break;
        }
      }
    }
    
    if (headerRowIdx >= 0) {
      // New format column positions:
      // Col 0: Bo., Col 1: (blank), Col 2: White, Col 3: Rtg, Col 4: Pts., 
      // Col 5: Result, Col 6: Pts., Col 7: (blank), Col 8: Black, Col 9: Rtg
      const boCol = 0;
      const whiteCol = 2;
      const whiteRtgCol = 3;
      const whitePtsCol = 4;
      const resultCol = 5;
      const blackPtsCol = 6;
      const blackCol = 8;
      const blackRtgCol = 9;
      
      for (let i = headerRowIdx + 1; i < data.length; i++) {
        const row = data[i];
        
        if (!row || row[boCol] === undefined || row[boCol] === null || row[boCol] === '') break;
        
        const board = parseInt(row[boCol]);
        const whiteName = String(row[whiteCol] || '').trim();
        const blackName = String(row[blackCol] || '').trim();
        
        if (!isNaN(board) && whiteName && blackName) {
          const resultStr = String(row[resultCol] || '').trim();
          const whitePts = row[whitePtsCol] !== undefined && row[whitePtsCol] !== null ? String(row[whitePtsCol]).trim() : '';
          const blackPts = row[blackPtsCol] !== undefined && row[blackPtsCol] !== null ? String(row[blackPtsCol]).trim() : '';
          
          pairings.push({
            board: board,
            whiteName: whiteName,
            whiteRating: parseInt(row[whiteRtgCol]) || 0,
            whitePoints: whitePts,
            result: resultStr,
            blackPoints: blackPts,
            blackRating: parseInt(row[blackRtgCol]) || 0,
            blackName: blackName
          });
          
          console.log(`Pairing ${board}: ${whiteName} (${parseInt(row[whiteRtgCol])}) vs ${blackName} (${parseInt(row[blackRtgCol])})`);
        }
      }
    }
    
    pairings.tournamentName = tournamentName;
    pairings.roundInfo = roundInfo;
    
    console.log('DEBUG - Tournament Name:', tournamentName);
    console.log('DEBUG - Round Info:', roundInfo);
    
    return pairings;
  } catch (error) {
    console.error('Error parsing XLSX:', error);
    return [];
  }
}

// Helper function to parse standings from XLSX
function parseStandingsFromXLSX(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    const standings = [];
    let headerRowIdx = -1;
    let tournamentName = '';
    let roundInfo = '';
    
    // Extract tournament name and round info from first few rows
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const row = data[i];
      if (row && row[0]) {
        const fullText = String(row[0]);
        const firstCell = fullText.toLowerCase();
        
        // Get tournament name
        if (firstCell.includes('collingwood') || firstCell.includes('chess') || firstCell.includes('somborac')) {
          tournamentName = fullText.trim();
        }
        
        // Get round info
        if (firstCell.includes('rank') && firstCell.includes('round')) {
          roundInfo = fullText.trim();
        }
      }
    }
    
    // Find header row by looking for "Rank" or "Rk" combined with "Name" and "Pts"
    for (let i = 0; i < Math.min(data.length, 20); i++) {
      const row = data[i];
      if (row && row.length > 3) {
        let headerScore = 0;
        for (let j = 0; j < row.length; j++) {
          const cell = String(row[j] || '').toLowerCase().trim();
          if (cell.includes('rank') || cell.includes('rk.') || cell === 'rk') headerScore++;
          if (cell.includes('name')) headerScore++;
          if (cell.includes('pts') && !cell.includes('tb')) headerScore++;
          if (cell.includes('rtg') || cell.includes('rating')) headerScore++;
        }
        // If we found at least 3 of these 4 keywords, it's likely the header row
        if (headerScore >= 3) {
          headerRowIdx = i;
          break;
        }
      }
    }
    
    if (headerRowIdx >= 0) {
      // Dynamically find column indices by searching header row
      const headerRow = data[headerRowIdx];
      let plCol = -1, nameCol = -1, scoreCol = -1, ratingCol = -1;
      
      // Search header row for column names - use exact abbreviated names
      for (let j = 0; j < headerRow.length; j++) {
        const cell = String(headerRow[j] || '').trim();
        if (cell === 'Rk.' || cell === 'Rk') plCol = j;
        if (cell === 'Name') nameCol = j;
        if (cell === 'Pts' || cell === 'Pts.') scoreCol = j;
        if (cell === 'Rtg') ratingCol = j;
      }
      
      console.log(`Found columns - Rank(${plCol}): "${headerRow[plCol]}", Name(${nameCol}): "${headerRow[nameCol]}", Score(${scoreCol}): "${headerRow[scoreCol]}", Rating(${ratingCol}): "${headerRow[ratingCol]}"`);
      
      // Only process if all columns were found
      if (plCol === -1 || nameCol === -1 || scoreCol === -1 || ratingCol === -1) {
        console.error('Could not find all required columns in header row');
        return standings;
      }
      
      for (let i = headerRowIdx + 1; i < data.length; i++) {
        const row = data[i];
        
        if (!row || row.length === 0) break;
        
        const name = String(row[nameCol] || '').trim();
        
        // Skip if name is empty or is a header text
        if (!name || name === 'Name' || name.length === 0) break;
        
        const rank = row[plCol];
        const scoreValue = parseFloat(row[scoreCol]) || 0;
        const rating = parseInt(row[ratingCol]) || 0;
        
        standings.push({
          rank: rank || i - headerRowIdx,
          name: name,
          score: scoreValue,
          rating: rating
        });
        
        console.log(`Standings ${rank}: ${name} - Score: ${scoreValue}, Rating: ${rating}`);
      }
    }
    
    standings.tournamentName = tournamentName;
    standings.roundInfo = roundInfo;
    
    return standings;
  } catch (error) {
    console.error('Error parsing standings XLSX:', error);
    return [];
  }
}

// Format date from YYYY-MM-DD to "Wednesday, May 1"
function formatDateNicely(dateStr) {
  if (!dateStr) return '';
  
  // dateStr is in format YYYY-MM-DD
  const [year, month, day] = dateStr.split('-');
  const date = new Date(year, parseInt(month) - 1, day);
  
  const options = { weekday: 'long', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

// Extract section name and round number
function extractSectionAndRound(tournamentName, roundInfo) {
  let section = '';
  let roundNum = '';
  let cleanedName = '';
  
  // Extract section - first try U format (u1500, u1400, etc)
  let sectionMatch = tournamentName.match(/[-\s](u\d+)\s*/i);
  if (sectionMatch) {
    section = sectionMatch[1].toUpperCase();
  } else {
    // If no U format, try to find other section names (Crown, etc)
    // Look for words that come after a dash or at the end
    sectionMatch = tournamentName.match(/[-\s]([A-Za-z]+)\s*$/);
    if (sectionMatch) {
      section = sectionMatch[1].toUpperCase();
    }
  }
  
  // Extract round number
  const roundMatch = roundInfo.match(/Round\s+(\d+)/i);
  if (roundMatch) {
    roundNum = `Round ${roundMatch[1]}`;
  }
  
  // Remove year and section from tournament name
  cleanedName = tournamentName
    .replace(/\s*-\s*u\d+\s*/i, '')
    .replace(/\s*-\s*[A-Za-z]+\s*$/, '')
    .replace(/\s+\d{4}\s*$/, '')
    .trim();
  
  return { section, roundNum, cleanedName };
}

// Generate pairings PDF with single result column
function generatePairingsPDF(pairings, tournamentInfo, section, roundNum, cleanedName, roundDate) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({
        size: 'Letter',
        margin: 30
      });
      
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);
      
      const pageWidth = doc.page.width - 60;
      const pageHeight = doc.page.height;
      const margins = 30;
      
      // Find longest player name
      let maxNameLength = 0;
      for (const pairing of pairings) {
        maxNameLength = Math.max(
          maxNameLength,
          pairing.whiteName.length,
          pairing.blackName.length
        );
      }
      
      // Calculate font size
      doc.fontSize(10).font('Helvetica');
      const testWidth = doc.widthOfString('A'.repeat(maxNameLength));
      
      let fontSize = 10;
      let nameColumnWidth = testWidth;
      
      while (nameColumnWidth > pageWidth * 0.28 && fontSize > 7) {
        fontSize--;
        doc.fontSize(fontSize);
        nameColumnWidth = doc.widthOfString('A'.repeat(maxNameLength));
      }
      
      const headerFontSize = Math.max(fontSize - 1, 7);
      const headerHeight = Math.ceil(fontSize * 1.8);
      
      // Column widths - Bo., White, Rating, Points, Result, Points, Rating, Black
      // Add left padding (3%) to name columns
      const namePadding = pageWidth * 0.03;
      const colWidths = {
        board: pageWidth * 0.05,
        whiteName: nameColumnWidth + namePadding,
        whiteRating: pageWidth * 0.075,
        whitePoints: pageWidth * 0.06,
        result: pageWidth * 0.075,
        blackPoints: pageWidth * 0.06,
        blackRating: pageWidth * 0.075,
        blackName: nameColumnWidth + namePadding
      };
      
      // Adjust if columns are too wide
      const totalWidth = Object.values(colWidths).reduce((a, b) => a + b, 0);
      if (totalWidth > pageWidth) {
        const scale = pageWidth / totalWidth;
        Object.keys(colWidths).forEach(key => {
          colWidths[key] *= scale;
        });
      }
      
      const maxPairingsPerPage = 25;
      const availableHeightPerPage = pageHeight - margins - 60 - headerHeight;
      const actualRowHeight = (availableHeightPerPage / maxPairingsPerPage) * 0.80;
      
      const finalFontSize = Math.max(Math.min(Math.floor(actualRowHeight / 2.2), 14), 8);
      const finalHeaderFontSize = Math.round(finalFontSize * 1.4 * 0.85); // 40% larger, then 15% smaller
      
      function drawPageHeader() {
        doc.fontSize(16).font('Helvetica-Bold').text(`${section} - ${roundNum}`, {
          align: 'center'
        });
        
        const niceDate = formatDateNicely(roundDate);
        const headerLine2 = niceDate ? `${cleanedName} - ${niceDate}` : cleanedName;
        doc.fontSize(10).font('Helvetica').text(headerLine2, {
          align: 'center'
        });
        
        doc.moveDown(0.5);
      }
      
      function drawTableHeader() {
        const headerY = doc.y;
        const headerHeight = Math.ceil(finalHeaderFontSize * 2);
        
        doc.rect(30, headerY, pageWidth, headerHeight).fillAndStroke('#2c3e50', '#2c3e50');
        
        doc.fontSize(finalHeaderFontSize).font('Helvetica-Bold').fillColor('white');
        let x = 30;
        const textY = headerY + (headerHeight / 2) - (finalHeaderFontSize / 2);
        
        doc.text('Bo.', x, textY, { width: colWidths.board, align: 'center', lineBreak: false });
        x += colWidths.board;
        doc.text('White', x, textY, { width: colWidths.whiteName, align: 'center', lineBreak: false });
        x += colWidths.whiteName;
        doc.text('Rtg', x, textY, { width: colWidths.whiteRating, align: 'center', lineBreak: false });
        x += colWidths.whiteRating;
        doc.text('Pts', x, textY, { width: colWidths.whitePoints, align: 'center', lineBreak: false });
        x += colWidths.whitePoints;
        doc.text('Result', x, textY, { width: colWidths.result, align: 'center', lineBreak: false });
        x += colWidths.result;
        doc.text('Pts', x, textY, { width: colWidths.blackPoints, align: 'center', lineBreak: false });
        x += colWidths.blackPoints;
        doc.text('Rtg', x, textY, { width: colWidths.blackRating, align: 'center', lineBreak: false });
        x += colWidths.blackRating;
        doc.text('Black', x, textY, { width: colWidths.blackName, align: 'center', lineBreak: false });
        
        doc.y = headerY + headerHeight;
      }
      
      function drawFooter() {
        const footerY = pageHeight - margins - 15;
        doc.fontSize(8).font('Helvetica').fillColor('#888888');
        doc.text('From the Tournament-Database of Chess-Results https://chess-results.com', 
          margins, footerY, 
          { align: 'center', width: pageWidth, lineBreak: false }
        );
      }
      
      drawPageHeader();
      drawTableHeader();
      
      let rowCount = 0;
      let pairingsOnCurrentPage = 0;
      
      for (let i = 0; i < pairings.length; i++) {
        const pairing = pairings[i];
        
        if (pairingsOnCurrentPage >= maxPairingsPerPage) {
          drawFooter();
          doc.addPage();
          drawPageHeader();
          drawTableHeader();
          rowCount = 0;
          pairingsOnCurrentPage = 0;
        }
        
        const rowY = doc.y;
        
        if (rowCount % 2 === 1) {
          doc.rect(30, rowY, pageWidth, actualRowHeight).fill('#e8eaed');
        } else {
          doc.rect(30, rowY, pageWidth, actualRowHeight).fill('white');
        }
        
        // Draw faint vertical lines between columns
        doc.strokeColor('#e0e0e0').lineWidth(0.3);
        let lineX = 30;
        
        lineX += colWidths.board;
        // Darker line after board number
        doc.strokeColor('#999999').lineWidth(0.5);
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        // Darker dashed bottom border (from left of board to 3/4 across white player name)
        // Only draw if not the last row on the page
        const isLastRow = (pairingsOnCurrentPage === maxPairingsPerPage - 1) || (i === pairings.length - 1);
        if (!isLastRow) {
          const boardBottomY = rowY + actualRowHeight;
          const threeQuarterWhiteWidth = colWidths.whiteName * 0.75;
          doc.strokeColor('#000000').lineWidth(0.6);
          doc.dash(2, { space: 2 });
          doc.moveTo(30, boardBottomY).lineTo(30 + colWidths.board + threeQuarterWhiteWidth, boardBottomY).stroke();
          doc.undash();
        }
        
        // Back to light lines for the rest
        doc.strokeColor('#e0e0e0').lineWidth(0.3);
        
        lineX += colWidths.whiteName;
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        lineX += colWidths.whiteRating;
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        lineX += colWidths.whitePoints;
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        lineX += colWidths.result;
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        lineX += colWidths.blackPoints;
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        lineX += colWidths.blackRating;
        doc.moveTo(lineX, rowY).lineTo(lineX, rowY + actualRowHeight).stroke();
        
        doc.fillColor('black').fontSize(finalFontSize).font('Helvetica');
        let x = 30;
        const textY = rowY + (actualRowHeight / 2) - (finalFontSize / 2) + 1;
        
        // Board number - larger and bolder
        doc.font('Helvetica-Bold').fontSize(finalFontSize * 1.25).text(String(pairing.board), x, textY, { 
          width: colWidths.board, align: 'center', lineBreak: false
        });
        doc.fontSize(finalFontSize); // Reset font size
        x += colWidths.board;
        
        // White name (with left padding) - BOLD
        doc.font('Helvetica-Bold').fontSize(finalFontSize);
        const whiteNameWidth = doc.widthOfString(pairing.whiteName);
        const whiteNameFits = whiteNameWidth < (colWidths.whiteName - 10);
        const whiteNameFontSize = whiteNameFits ? finalFontSize : finalFontSize * 0.85;
        doc.fontSize(whiteNameFontSize).text(pairing.whiteName, x + 5, textY, { 
          width: colWidths.whiteName - 10, align: 'left', lineBreak: false
        });
        doc.fontSize(finalFontSize);
        x += colWidths.whiteName;
        
        // White rating
        doc.font('Helvetica').text(String(pairing.whiteRating), x, textY, { 
          width: colWidths.whiteRating, align: 'center', lineBreak: false
        });
        x += colWidths.whiteRating;
        
        // White points (show if exists)
        doc.text(pairing.whitePoints, x, textY, { 
          width: colWidths.whitePoints, align: 'center', lineBreak: false
        });
        x += colWidths.whitePoints;
        
        // Result (blank for both to write)
        x += colWidths.result;
        
        // Black points (show if exists)
        doc.font('Helvetica').text(String(pairing.blackPoints), x, textY, { 
          width: colWidths.blackPoints, align: 'center', lineBreak: false
        });
        x += colWidths.blackPoints;
        
        // Black rating
        doc.text(String(pairing.blackRating), x, textY, { 
          width: colWidths.blackRating, align: 'center', lineBreak: false
        });
        x += colWidths.blackRating;
        
        // Black name (with left padding) - BOLD
        doc.font('Helvetica-Bold').fontSize(finalFontSize);
        const blackNameWidth = doc.widthOfString(pairing.blackName);
        const blackNameFits = blackNameWidth < (colWidths.blackName - 10);
        const blackNameFontSize = blackNameFits ? finalFontSize : finalFontSize * 0.85;
        doc.fontSize(blackNameFontSize).text(pairing.blackName, x + 5, textY, { 
          width: colWidths.blackName - 10, align: 'left', lineBreak: false
        });
        
        doc.y = rowY + actualRowHeight;
        rowCount++;
        pairingsOnCurrentPage++;
      }
      
      drawFooter();
      doc.end();
    } catch (error) {
      reject(error);
    }
  });
}

// Generate standings PDF
function generateStandingsPDF(standings, tournamentInfo, section, standingsRound, cleanedName) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({
        size: 'Letter',
        margin: 30
      });
      
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);
      
      const pageWidth = doc.page.width - 60;
      const pageHeight = doc.page.height;
      const margins = 30;
      
      // Find longest player name (actual pixel width)
      let maxNameWidth = 0;
      doc.fontSize(10).font('Helvetica');
      for (const entry of standings) {
        const nameWidth = doc.widthOfString(entry.name);
        maxNameWidth = Math.max(maxNameWidth, nameWidth);
      }
      
      // Calculate font size
      let fontSize = 10;
      let nameColumnWidth = maxNameWidth;
      
      while (nameColumnWidth > pageWidth * 0.5 && fontSize > 7) {
        fontSize--;
        doc.fontSize(fontSize);
        nameColumnWidth = doc.widthOfString(standings.reduce((max, e) => e.name.length > max.length ? e : max).name);
      }
      
      const headerFontSize = Math.max(fontSize - 1, 7);
      const headerHeight = Math.ceil(fontSize * 1.8);
      
      // Column widths for standings
      const colWidths = {
        rank: pageWidth * 0.08,
        name: maxNameWidth + 20,
        score: pageWidth * 0.12,
        rating: pageWidth * 0.12
      };
      
      const maxEntriesPerPage = 20;
      const availableHeightPerPage = pageHeight - margins - 60 - headerHeight;
      const actualRowHeight = (availableHeightPerPage / maxEntriesPerPage) * 0.80;
      
      const finalFontSize = Math.max(Math.min(Math.floor(actualRowHeight / 2.2), 14), 8);
      const finalHeaderFontSize = Math.max(finalFontSize - 1, 7);
      
      function drawPageHeader() {
        doc.fontSize(16).font('Helvetica-Bold').text(`${section} - ${standingsRound}`, {
          align: 'center'
        });
        
        doc.fontSize(10).font('Helvetica').text(`${cleanedName}`, {
          align: 'center'
        });
        
        doc.moveDown(0.3);
      }
      
      function drawTableHeader() {
        const headerY = doc.y;
        
        doc.rect(30, headerY, pageWidth, headerHeight).fillAndStroke('#2c3e50', '#2c3e50');
        
        doc.fontSize(finalHeaderFontSize).font('Helvetica-Bold').fillColor('white');
        let x = 30;
        const textY = headerY + (headerHeight / 2) - (finalHeaderFontSize / 2);
        
        doc.text('Rank', x, textY, { width: colWidths.rank, align: 'center', lineBreak: false });
        x += colWidths.rank;
        doc.text('Name', x, textY, { width: colWidths.name, align: 'left', lineBreak: false });
        x += colWidths.name;
        doc.text('Score', x, textY, { width: colWidths.score, align: 'center', lineBreak: false });
        x += colWidths.score;
        doc.text('Rating', x, textY, { width: colWidths.rating, align: 'center', lineBreak: false });
        
        doc.y = headerY + headerHeight;
      }
      
      function drawFooter() {
        const footerY = pageHeight - margins - 15;
        doc.fontSize(8).font('Helvetica').fillColor('#888888');
        doc.text('From the Tournament-Database of Chess-Results https://chess-results.com', 
          margins, footerY, 
          { align: 'center', width: pageWidth, lineBreak: false }
        );
      }
      
      drawPageHeader();
      drawTableHeader();
      
      let rowCount = 0;
      let entriesOnCurrentPage = 0;
      
      for (let i = 0; i < standings.length; i++) {
        const entry = standings[i];
        
        if (entriesOnCurrentPage >= maxEntriesPerPage) {
          drawFooter();
          doc.addPage();
          drawPageHeader();
          drawTableHeader();
          rowCount = 0;
          entriesOnCurrentPage = 0;
        }
        
        const rowY = doc.y;
        
        if (rowCount % 2 === 1) {
          doc.rect(30, rowY, pageWidth, actualRowHeight).fill('#f8f9fa');
        } else {
          doc.rect(30, rowY, pageWidth, actualRowHeight).fill('white');
        }
        
        doc.fillColor('black').fontSize(finalFontSize).font('Helvetica');
        let x = 30;
        const textY = rowY + (actualRowHeight / 2) - (finalFontSize / 2);
        
        doc.font('Helvetica-Bold').text(String(entry.rank), x, textY, { 
          width: colWidths.rank, align: 'center', lineBreak: false
        });
        x += colWidths.rank;
        
        // Check if name fits on one line
        doc.font('Helvetica').fontSize(finalFontSize);
        const nameWidth = doc.widthOfString(entry.name);
        const nameFits = nameWidth < (colWidths.name - 10);
        const nameFontSize = nameFits ? finalFontSize : finalFontSize * 0.85;
        
        doc.fontSize(nameFontSize).text(entry.name, x, textY, { 
          width: colWidths.name, align: 'left', lineBreak: false
        });
        doc.fontSize(finalFontSize); // Reset
        x += colWidths.name;
        
        doc.font('Helvetica-Bold').text(String(entry.score === 0 ? '0' : entry.score.toFixed(1)), x, textY, { 
          width: colWidths.score, align: 'center', lineBreak: false
        });
        x += colWidths.score;
        
        doc.font('Helvetica').text(String(entry.rating), x, textY, { 
          width: colWidths.rating, align: 'center', lineBreak: false
        });
        
        doc.y = rowY + actualRowHeight;
        rowCount++;
        entriesOnCurrentPage++;
      }
      
      drawFooter();
      doc.end();
    } catch (error) {
      reject(error);
    }
  });
}

// API endpoint to process files
app.post('/api/process', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file provided' });
    }
    
    const tableType = req.body.tableType || 'pairings';
    
    let data = [];
    let tournamentInfo = { tournament: '', round: '', date: '' };
    let section = '';
    let roundNum = '';
    let cleanedName = '';
    let previewData = [];
    
    const fileExtension = req.file.originalname.split('.').pop().toLowerCase();
    
    if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      if (tableType === 'standings') {
        data = parseStandingsFromXLSX(req.file.buffer);
        
        if (data.length === 0) {
          return res.status(400).json({ 
            error: 'Could not extract standings from Excel file.' 
          });
        }
        
        if (data.tournamentName) {
          tournamentInfo.tournament = data.tournamentName;
        }
        
        if (data.roundInfo) {
          roundNum = data.roundInfo.trim();
        } else {
          roundNum = 'Standings';
        }
        
        const { section: sec, cleanedName: clean } = extractSectionAndRound(data.tournamentName || '', '');
        section = sec;
        cleanedName = clean;
        
        previewData = data.slice(0, 5).map(entry => ({
          rank: entry.rank,
          name: entry.name,
          score: entry.score,
          rating: entry.rating
        }));
        
        delete data.tournamentName;
        delete data.roundInfo;
      } else {
        data = parsePairingsFromXLSX(req.file.buffer);
        
        if (data.length === 0) {
          return res.status(400).json({ 
            error: 'Could not extract pairings from Excel file.' 
          });
        }
        
        if (data.tournamentName) {
          tournamentInfo.tournament = data.tournamentName;
        }
        
        if (data.roundInfo) {
          const match = data.roundInfo.match(/Round\s+(\d+)\s+on\s+(\d{4})\/(\d{2})\/(\d{2})/);
          if (match) {
            tournamentInfo.round = `Round ${match[1]}`;
            tournamentInfo.date = `${match[2]}-${match[3]}-${match[4]}`;
          }
        }
        
        const { section: sec, roundNum: rnd, cleanedName: clean } = extractSectionAndRound(data.tournamentName, data.roundInfo);
        section = sec;
        roundNum = rnd;
        cleanedName = clean;
        
        previewData = data.slice(0, 5).map(pairing => ({
          board: pairing.board,
          whiteName: pairing.whiteName,
          whiteRating: pairing.whiteRating,
          result: pairing.result,
          blackRating: pairing.blackRating,
          blackName: pairing.blackName
        }));
        
        delete data.tournamentName;
        delete data.roundInfo;
      }
    } else {
      return res.status(400).json({ error: 'Please upload a .xlsx or .xls file' });
    }
    
    // Generate PDF
    let newPdfBuffer;
    if (tableType === 'standings') {
      newPdfBuffer = await generateStandingsPDF(data, tournamentInfo, section, roundNum, cleanedName);
    } else {
      const finalRoundDate = tournamentInfo.date || '';
      newPdfBuffer = await generatePairingsPDF(data, tournamentInfo, section, roundNum, cleanedName, finalRoundDate);
    }
    
    // Store PDF
    const pdfId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
    pdfStore[pdfId] = newPdfBuffer;
    
    // Send response with preview data
    res.json({
      pdfId: pdfId,
      tableType: tableType,
      count: data.length,
      preview: previewData,
      tournament: cleanedName,
      roundNum: roundNum
    });
  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).json({ error: 'Error processing file: ' + error.message });
  }
});

// ============================================================================
// CHESS-RESULTS URL SCRAPER - Completely separate feature
// ============================================================================

const cheerio = require('cheerio');
const axios = require('axios');

// Scrape pairings from chess-results.com HTML
async function scrapePairingsFromURL(url) {
  try {
    console.log('Fetching URL:', url);
    const response = await axios.get(url, { timeout: 10000 });
    console.log('Response received, parsing HTML...');
    const $ = cheerio.load(response.data);
    
    const pairings = [];
    let tournamentName = 'Tournament';
    let roundInfo = '';
    
    // Extract tournament name from h1 or page title
    const h1 = $('h1').first().text().trim();
    const pageTitle = $('title').text().trim();
    tournamentName = h1 || pageTitle || 'Tournament';
    console.log('Tournament name:', tournamentName);
    
    // Find round info
    const roundMatch = pageTitle.match(/Round\s+(\d+)/i) || h1.match(/Round\s+(\d+)/i);
    if (roundMatch) {
      roundInfo = `Round ${roundMatch[1]} on ${new Date().toISOString().split('T')[0]}`;
    }
    console.log('Round info:', roundInfo);
    
    // Parse tables - look for pairings table
    $('table').each((tableIdx, table) => {
      const rows = $(table).find('tr');
      let headerFound = false;
      let boCol = -1, whiteCol = -1, whiteRtgCol = -1, whitePtsCol = -1, 
          resultCol = -1, blackPtsCol = -1, blackRtgCol = -1, blackCol = -1;
      
      rows.each((rowIdx, row) => {
        const cells = $(row).find('td, th');
        const cellTexts = cells.map((i, cell) => $(cell).text().trim()).get();
        
        if (cellTexts.length === 0) return;
        
        // Find header row - look for "Bo." or "Board"
        if (!headerFound && (cellTexts[0].toLowerCase().includes('bo') || cellTexts[0].toLowerCase().includes('board'))) {
          console.log('Found header at table', tableIdx, 'row', rowIdx, ':', cellTexts);
          headerFound = true;
          
          // Dynamically find column positions from header
          for (let i = 0; i < cellTexts.length; i++) {
            const header = cellTexts[i].toLowerCase().trim();
            if (header === 'bo.' || header === 'bo') boCol = i;
            if (header === 'w' || header.includes('white') || (i === 1 && boCol === 0)) whiteCol = i;
            if (header === 'rtg' && i < cellTexts.length / 2 && whiteRtgCol === -1) whiteRtgCol = i;
            if (header === 'pts' && i < cellTexts.length / 2 && whitePtsCol === -1) whitePtsCol = i;
            if (header.includes('result') || header === 'res') resultCol = i;
            if (header === 'pts' && i > cellTexts.length / 2 && blackPtsCol === -1) blackPtsCol = i;
            if (header === 'rtg' && i > cellTexts.length / 2 && blackRtgCol === -1) blackRtgCol = i;
            if (header === 'b' || header.includes('black') || (i === cellTexts.length - 2 && blackCol === -1)) blackCol = i;
          }
          
          console.log('Column mapping: bo:', boCol, 'white:', whiteCol, 'whiteRtg:', whiteRtgCol, 
            'whitePts:', whitePtsCol, 'result:', resultCol, 'blackPts:', blackPtsCol, 'blackRtg:', blackRtgCol, 'black:', blackCol);
          return;
        }
        
        // Parse data rows
        if (headerFound && cellTexts.length > boCol) {
          const boardNum = parseInt(cellTexts[boCol]);
          if (!isNaN(boardNum) && boardNum > 0) {
            const pairing = {
              board: boardNum,
              whiteName: whiteCol >= 0 ? cellTexts[whiteCol] || '' : '',
              whiteRating: whiteRtgCol >= 0 ? parseInt(cellTexts[whiteRtgCol]) || 0 : 0,
              whitePoints: whitePtsCol >= 0 ? cellTexts[whitePtsCol] || '' : '',
              result: resultCol >= 0 ? cellTexts[resultCol] || '' : '',
              blackPoints: blackPtsCol >= 0 ? cellTexts[blackPtsCol] || '' : '',
              blackRating: blackRtgCol >= 0 ? parseInt(cellTexts[blackRtgCol]) || 0 : 0,
              blackName: blackCol >= 0 ? cellTexts[blackCol] || '' : ''
            };
            console.log('Pairing', boardNum, ':', pairing);
            pairings.push(pairing);
          }
        }
      });
    });
    
    console.log('Found', pairings.length, 'pairings');
    pairings.tournamentName = tournamentName;
    pairings.roundInfo = roundInfo;
    
    return pairings;
  } catch (error) {
    console.error('Error scraping pairings:', error.message);
    throw new Error('Failed to scrape pairings: ' + error.message);
  }
}

// Scrape standings from chess-results.com HTML
async function scrapeStandingsFromURL(url) {
  try {
    console.log('Fetching URL:', url);
    const response = await axios.get(url, { timeout: 10000 });
    console.log('Response received, parsing HTML...');
    const $ = cheerio.load(response.data);
    
    const standings = [];
    let tournamentName = 'Tournament';
    let roundInfo = '';
    
    // Extract tournament name
    const h1 = $('h1').first().text().trim();
    const pageTitle = $('title').text().trim();
    tournamentName = h1 || pageTitle || 'Tournament';
    console.log('Tournament name:', tournamentName);
    
    // Find round info
    const roundMatch = pageTitle.match(/Round\s+(\d+)/i) || pageTitle.match(/After\s+Round\s+(\d+)/i) || h1.match(/Round\s+(\d+)/i);
    if (roundMatch) {
      roundInfo = `Rank after Round ${roundMatch[1]}`;
    }
    console.log('Round info:', roundInfo);
    
    // Parse standings table
    $('table').each((tableIdx, table) => {
      const rows = $(table).find('tr');
      let headerRowIdx = -1;
      let rankCol = -1, nameCol = -1, scoreCol = -1, ratingCol = -1;
      
      // First pass: find header row
      rows.each((rowIdx, row) => {
        if (headerRowIdx !== -1) return; // Already found, skip
        
        const cells = $(row).find('td, th');
        const cellTexts = cells.map((i, cell) => $(cell).text().trim()).get();
        
        if (cellTexts.length === 0) return;
        
        // Check if this row contains header keywords
        let foundHeader = false;
        for (let i = 0; i < cellTexts.length; i++) {
          const header = cellTexts[i].toLowerCase().trim();
          if (header === 'rk.' || header === 'rk' || header === 'rank') {
            rankCol = i;
            foundHeader = true;
          }
          if (header === 'name') {
            nameCol = i;
            foundHeader = true;
          }
          if (header === 'pts' || header === 'pts.' || header === 'pts. ' || header === 'score') {
            scoreCol = i;
            foundHeader = true;
          }
          if (header === 'rtg' || header === 'rating') {
            ratingCol = i;
            foundHeader = true;
          }
        }
        
        if (foundHeader && rankCol >= 0 && nameCol >= 0 && scoreCol >= 0 && ratingCol >= 0) {
          headerRowIdx = rowIdx;
          console.log('Found standings header at table', tableIdx, 'row', rowIdx);
          console.log('Column mapping: rank:', rankCol, 'name:', nameCol, 'score:', scoreCol, 'rating:', ratingCol);
        }
      });
      
      // Second pass: parse data rows (only after we found the header)
      if (headerRowIdx >= 0) {
        rows.each((rowIdx, row) => {
          // Skip rows up to and including the header row
          if (rowIdx <= headerRowIdx) return;
          
          const cells = $(row).find('td, th');
          const cellTexts = cells.map((i, cell) => $(cell).text().trim()).get();
          
          if (cellTexts.length === 0) return;
          if (cellTexts.length <= Math.max(rankCol, nameCol, scoreCol, ratingCol)) return;
          
          const name = cellTexts[nameCol] || '';
          const rankText = cellTexts[rankCol] || '';
          
          // Skip if name is empty, is header text, or row looks like a header
          if (!name || name === 'Name' || name.length === 0) return;
          if (name === 'FED' || name.toUpperCase() === name) return; // Skip header-looking rows
          
          const rankNum = parseInt(rankText);
          // Skip if rank is not a valid number
          if (isNaN(rankNum)) return;
          
          const entry = {
            rank: rankNum || standings.length + 1,
            name: name,
            score: parseFloat(cellTexts[scoreCol]) || 0,
            rating: parseInt(cellTexts[ratingCol]) || 0
          };
          console.log('Entry', entry.rank, ':', entry);
          standings.push(entry);
        });
      }
    });
    
    console.log('Found', standings.length, 'standings entries');
    standings.tournamentName = tournamentName;
    standings.roundInfo = roundInfo;
    
    return standings;
  } catch (error) {
    console.error('Error scraping standings:', error.message);
    throw new Error('Failed to scrape standings: ' + error.message);
  }
}

// API endpoint for URL scraping
app.post('/api/scrape', async (req, res) => {
  console.log('=== SCRAPER ENDPOINT CALLED ===');
  try {
    const { url, tableType } = req.body;
    console.log('1. Request received - URL:', url, 'Type:', tableType);
    
    if (!url || !tableType) {
      console.log('2. FAILED - Missing URL or tableType');
      return res.status(400).json({ error: 'URL and tableType required' });
    }
    
    let data;
    try {
      console.log('2. Starting scrape...');
      if (tableType === 'pairings') {
        console.log('2a. Calling scrapePairingsFromURL');
        data = await scrapePairingsFromURL(url);
        console.log('2b. scrapePairingsFromURL returned:', data.length, 'pairings');
      } else if (tableType === 'standings') {
        console.log('2a. Calling scrapeStandingsFromURL');
        data = await scrapeStandingsFromURL(url);
        console.log('2b. scrapeStandingsFromURL returned:', data.length, 'standings');
      } else {
        console.log('2. FAILED - Invalid tableType');
        return res.status(400).json({ error: 'Invalid tableType' });
      }
    } catch (scrapeErr) {
      console.error('2. SCRAPING FAILED:', scrapeErr.message);
      console.error('Stack:', scrapeErr.stack);
      return res.status(500).json({ error: 'Failed to scrape URL: ' + scrapeErr.message });
    }
    
    console.log('3. Scrape successful, extracting section/round info');
    const { section, roundNum, cleanedName } = extractSectionAndRound(data.tournamentName, data.roundInfo);
    console.log('3a. Extracted - Section:', section, 'Round:', roundNum, 'Name:', cleanedName);
    
    // Generate PDF
    let pdfBuffer;
    try {
      console.log('4. Starting PDF generation...');
      if (tableType === 'pairings') {
        console.log('4a. Calling generatePairingsPDF');
        pdfBuffer = await generatePairingsPDF(data, { name: cleanedName, date: '' }, section, roundNum, cleanedName, '');
        console.log('4b. generatePairingsPDF completed, buffer size:', pdfBuffer.length);
      } else {
        console.log('4a. Calling generateStandingsPDF');
        pdfBuffer = await generateStandingsPDF(data, { name: cleanedName }, section, roundNum);
        console.log('4b. generateStandingsPDF completed, buffer size:', pdfBuffer.length);
      }
    } catch (pdfErr) {
      console.error('4. PDF GENERATION FAILED:', pdfErr.message);
      console.error('Stack:', pdfErr.stack);
      return res.status(500).json({ error: 'Failed to generate PDF: ' + pdfErr.message });
    }
    
    console.log('5. Storing PDF in pdfStore');
    const pdfId = Date.now().toString();
    pdfStore[pdfId] = pdfBuffer;
    console.log('5a. PDF stored with ID:', pdfId);
    
    console.log('6. Sending success response');
    res.json({
      pdfId: pdfId,
      tableType: tableType,
      count: data.length,
      tournament: data.tournamentName,
      roundInfo: data.roundInfo
    });
    console.log('=== SCRAPER ENDPOINT COMPLETE ===');
  } catch (error) {
    console.error('=== SCRAPER ENDPOINT ERROR ===');
    console.error('Error:', error.message);
    console.error('Stack:', error.stack);
    res.status(500).json({ error: error.message });
  }
});

// Preview pairings columns and suggest names
app.post('/api/preview-pairings', upload.single('file'), (req, res) => {
  try {
    const buffer = req.file.buffer;
    console.log('3a. Parsing preview...');
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Find header row
    let headerRowIdx = -1;
    let headerRow = [];
    
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const row = data[i];
      if (row && row[0]) {
        const firstCell = String(row[0]).toLowerCase();
        if (firstCell === 'bo.' || firstCell === 'bo' || (row.some(cell => String(cell).toLowerCase().includes('white') || String(cell).toLowerCase().includes('black')))) {
          headerRowIdx = i;
          headerRow = row;
          break;
        }
      }
    }
    
    // Get 2-3 sample data rows
    const sampleRows = data.slice(headerRowIdx + 1, headerRowIdx + 4);
    
    // Suggest column names
    const suggestions = headerRow.map(header => {
      const h = String(header).toLowerCase().trim();
      if (h === 'bo.' || h === 'bo') return 'Board';
      if (h === 'w' || h.includes('white')) return 'White Player';
      if (h === 'b' || h.includes('black')) return 'Black Player';
      if (h === 'rtg' || h === 'rating') return 'Rating';
      if (h === 'pts' || h === 'points' || h === 'score') return 'Points';
      if (h === 'result' || h === 'res') return 'Result';
      return header; // Keep original if no match
    });
    
    console.log('3b. Preview data prepared');
    res.json({
      headerRowIdx: headerRowIdx,
      headers: headerRow,
      suggestions: suggestions,
      samples: sampleRows,
      totalRows: data.length
    });
  } catch (error) {
    console.error('Preview error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Process pairings with column configuration
app.post('/api/process-pairings-configured', upload.single('file'), async (req, res) => {
  try {
    const { columnMap } = req.body;
    console.log('Processing with column map:', columnMap);
    
    const buffer = req.file.buffer;
    const pairings = parsePairingsFromXLSX(buffer);
    
    console.log('Found', pairings.length, 'pairings');
    pairings.tournamentName = tournamentName;
    pairings.roundInfo = roundInfo;
    pairings.columnConfig = JSON.parse(columnMap);
    
    res.json({
      success: true,
      count: pairings.length,
      tournament: pairings.tournamentName,
      roundInfo: pairings.roundInfo
    });
  } catch (error) {
    console.error('Processing error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Download endpoint
app.get('/api/download/:pdfId', (req, res) => {
  const { pdfId } = req.params;
  
  if (!pdfStore[pdfId]) {
    return res.status(404).json({ error: 'PDF not found' });
  }
  
  const pdfBuffer = pdfStore[pdfId];
  
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="chess-tournament-formatted.pdf"');
  res.send(pdfBuffer);
  
  setTimeout(() => {
    delete pdfStore[pdfId];
  }, 5000);
});

// Test file endpoint
app.get('/test-file/:filename', (req, res) => {
  const filename = req.params.filename;
  const filepath = path.join(__dirname, filename);
  
  // Security: only allow test file
  if (filename !== 'chessResultsList_3_.xlsx') {
    return res.status(404).json({ error: 'File not found' });
  }
  
  if (fs.existsSync(filepath)) {
    res.download(filepath);
  } else {
    res.status(404).json({ error: 'Test file not found' });
  }
});

const PORT = process.env.PORT || 3000;
const restartCounter = getRestartCounter();
saveRestartCounter(restartCounter);

app.listen(PORT, () => {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`Chess Tournament Formatter #${restartCounter}`);
  console.log(`Running on http://localhost:${PORT}`);
  console.log(`${'='.repeat(60)}\n`);
});
