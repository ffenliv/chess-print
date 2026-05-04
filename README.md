# Chess Tournament Formatter

Convert chess tournament pairings and standings from Excel into beautiful, printer-friendly PDFs.

## Features

- ✅ Format **Pairings** tables (Chess-Results.com export)
- ✅ Format **Standings** tables
- ✅ Auto-detect tournament name and round information
- ✅ Professional PDF layout with alternating row colors
- ✅ Dynamic font sizing for readability
- ✅ Chess-Results.com attribution footer
- ✅ Multi-page support with repeating headers

## Installation

### Requirements
- Node.js 14+
- npm

### Setup

1. Clone or extract this project
2. Install dependencies:
```bash
npm install
```

3. Start the server:
```bash
npm start
```

Or for development with auto-reload:
```bash
npm run dev
```

4. Open http://localhost:3000 in your browser

## Usage

1. Select whether you're formatting a **Pairings** or **Standings** table
2. Upload your Excel file (.xlsx or .xls)
3. The app will automatically detect:
   - Tournament name
   - Round/Standings information
   - Section level (u1500, u1400, etc)
4. Click "Download PDF" to save the formatted document

## Excel File Format

### Pairings Table
Export from Chess-Results.com with these columns:
- Bo. (Board number)
- White player name
- Rtg (White rating)
- Result
- Black player name
- Rtg (Black rating)

### Standings Table
- Pl. or Rank (Player place)
- Name
- Rating
- Score

## File Structure

```
chess-tournament-formatter/
├── server.js              # Express backend
├── package.json          # Dependencies
├── README.md             # This file
└── public/
    └── index.html        # Web interface
```

## API Endpoints

### POST /api/process
Upload and process a file
- Body: multipart form with `file` and `tableType`
- Returns: `{ pdfId, count, tableType }`

### GET /api/download/:pdfId
Download the generated PDF

## Customization

Edit `server.js` to customize:
- Column widths
- Font sizes
- Colors (header background #2c3e50)
- Rows per page (currently 15 for pairings, 20 for standings)
- Footer text

## License

MIT
