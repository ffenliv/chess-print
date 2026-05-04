# Quick Start Guide

## 1. Extract Files
Extract all files to a folder called `chess-tournament-formatter`

## 2. Install Dependencies
```bash
cd chess-tournament-formatter
npm install
```

## 3. Start Server
```bash
npm start
```

You'll see:
```
Chess Tournament Formatter running on http://localhost:3000
```

## 4. Open in Browser
Go to http://localhost:3000

## 5. Use the App

### For Pairings:
1. Select "Pairings Table"
2. Upload your Chess-Results.com pairings export (.xlsx)
3. Click "Process File"
4. Click "Download PDF"

### For Standings:
1. Select "Standings Table"
2. Upload your standings file (.xlsx)
3. Click "Process File"
4. Click "Download PDF"

## Development Mode (Auto-reload on file changes)
```bash
npm run dev
```

## Troubleshooting

**Port 3000 already in use?**
```bash
PORT=3001 npm start
```
Then go to http://localhost:3001

**File not recognized?**
- Make sure it's .xlsx or .xls format
- For pairings: file must have "Bo.", "White", "Rtg", "Result", "Black" columns
- For standings: file must have "Rank"/"Pl." and "Name" columns

**Missing dependencies?**
```bash
npm install
```

## Support

Check the README.md file for more detailed information.
