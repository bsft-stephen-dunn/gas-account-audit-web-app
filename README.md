# Account Audit Web App

A Google Apps Script web application for generating comprehensive account audit documents for Blueshift accounts. This tool streamlines the process of extracting and formatting account configuration data from Rails console.

## 🚀 Features

- **Automated Script Generation**: Dynamically generates Rails extraction scripts based on the provided site name
- **JSON Data Processing**: Processes complex nested JSON structures from Rails console outputs
- **Document Generation**: Creates formatted Google Docs with comprehensive account audit information
- **Dark Mode Support**: Toggle between light and dark themes for comfortable viewing
- **Mobile Responsive**: Optimized for both desktop and mobile devices
- **Syntax Highlighting**: Color-coded Ruby scripts for better readability

## 📋 Prerequisites

- Google Account with access to Google Apps Script
- Access to Blueshift Rails console

## 🛠️ Installation

### Using CLASP (Command Line Apps Script)

1. Clone this repository:
```bash
git clone https://github.com/bsft-stephen-dunn/gas-account-audit-web-app.git
cd gas-account-audit-web-app
```

2. Install CLASP globally (if not already installed):
```bash
npm install -g @google/clasp
```

3. Login to CLASP:
```bash
clasp login
```

4. Push the code to your Google Apps Script project:
```bash
clasp push
```

5. Deploy the web app:
```bash
clasp deploy
```

### Manual Installation

1. Go to [Google Apps Script](https://script.google.com)
2. Create a new project
3. Copy the contents of each file in this repository to the corresponding file in your Apps Script project:
   - `Code.js` → Code.gs
   - `Index_Styled.html` → Index_Styled.html
   - `JavaScript.html` → JavaScript.html
   - `Stylesheet.html` → Stylesheet.html
   - `appsscript.json` → (File > Project Settings > Show "appsscript.json" manifest file)

## 📖 Usage

### Folder Selection Feature
The app includes a folder picker that lets you choose where to save your audit documents:
- **Default**: Uses a dropdown menu to select from your Google Drive folders (no setup required)
- **Optional**: Can use Google's native Picker UI (requires API configuration - see [GOOGLE_PICKER_SETUP.md](GOOGLE_PICKER_SETUP.md))

### Step 1: Configure Site Name
1. Open the deployed web app
2. Enter the site name (e.g., `example.com`) in the input field
3. Click "Update Script" to generate a customized Rails extraction script

### Step 2: Extract Data from Rails Console
1. Copy the generated Ruby script using the "Copy Script" button
2. Access your Blueshift Rails console
3. Paste and execute the entire script
4. The script will automatically run and output JSON data
5. Copy the JSON output that appears after "EXTRACTION COMPLETE"

### Step 3: Generate Audit Document
1. Paste the Rails JSON output in the "Rails Extraction Data" field
2. Click "Generate Audit Document"
3. The app will create a Google Doc with the formatted audit information
4. A link to the document will be provided upon completion

## 🏗️ Project Structure

```
gas-account-audit-web-app/
├── Code.js              # Server-side Google Apps Script code
├── Index_Styled.html    # Main HTML template
├── JavaScript.html      # Client-side JavaScript
├── Stylesheet.html      # CSS styles with Blueshift branding
├── appsscript.json      # Apps Script manifest
├── .clasp.json          # CLASP configuration
└── README.md           # This file
```

## 🎨 Branding

This application follows Blueshift's brand guidelines with a professional blue color scheme and clean, modern design. The interface is optimized for both light and dark modes.

## 🔧 Development

### Local Development with CLASP

1. Make changes to the files locally
2. Push changes to Google Apps Script:
```bash
clasp push
```

3. Test the web app by opening the deployment URL

### Version Control

1. Stage your changes:
```bash
git add .
```

2. Commit with a descriptive message:
```bash
git commit -m "Your descriptive commit message"
```

3. Push to GitHub:
```bash
git push origin main
```

## 📱 Mobile Responsiveness

The application is fully responsive and optimized for:
- Desktop computers (1200px+)
- Tablets (768px - 1199px)
- Mobile phones (<768px)

## 🔒 Security

- All data is processed client-side before being sent to Google Apps Script
- Generated documents are stored in the user's Google Drive
- No data is stored permanently by the application

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is proprietary to Blueshift and for internal use only.

## 👥 Contact

For questions or support, please contact the Blueshift development team.

---

Built with ❤️ for the Blueshift team