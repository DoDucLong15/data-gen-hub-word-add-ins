# Word Data Mapper Add-in

## Overview
This project is a Microsoft Word Add-in that allows users to map regions of a Word document to database fields, configure naming formats, and generate a specification (`spec.json`) file describing the mapping. It is designed to assist with data extraction and integration between Word documents and databases, making it easier to automate data flows for academic or business processes.

## Features
- Select and map parts of a Word document to specific database fields.
- Configure custom naming formats for exported data.
- Support for mapping single fields, combined fields, constants, and comments.
- Preview the generated mapping specification as JSON.
- Export the mapping as a `spec.json` file for use in other systems.

## Prerequisites
- [Node.js](https://nodejs.org/) (v14 or higher recommended)
- [npm](https://www.npmjs.com/) (comes with Node.js)
- Microsoft Word (Office 2016 or later, or Microsoft 365)

## Getting Started

### 1. Clone the Repository
```bash
git clone <your-repo-url>
cd data-gen-hub/add-ins/data-gen-hub-word
```

### 2. Install Dependencies
```bash
npm install
```

### 3. Build the Project
For production build:
```bash
npm run build
```
For development (with hot reload):
```bash
npm run dev-server
```

### 4. Sideload the Add-in in Word
1. Build or start the dev server as above.
2. Open Word and go to **Insert > My Add-ins > Shared Folder** (or sideload via Office Add-in sideloading tools).
3. Select and add the add-in using the `manifest.xml` file in the project root.

### 5. Using the Add-in
- Open the add-in task pane in Word.
- Configure the naming format as needed.
- Select database tables and fields to map document regions.
- Use the UI to map, preview, and export the specification file.

## Scripts
- `npm run build` – Build the add-in for production.
- `npm run dev-server` – Start the development server with hot reload.
- `npm start` – Launch the add-in in Word for debugging (requires Office Add-in Debugging tools).
- `npm run lint` – Run linter.
- `npm run lint:fix` – Auto-fix lint issues.

## Project Structure
- `src/taskpane/` – Main UI and logic for the task pane add-in.
- `src/commands/` – Command functions for Office ribbon integration.
- `manifest.xml` – Office Add-in manifest file.

## License
This project is licensed under the MIT License.

## Contribution
Feel free to submit issues or pull requests to improve the add-in!
