# Navigation Hub - Flask Web Application

A beautiful web application built with Flask that provides quick navigation links to popular websites.

## Features

- Modern, responsive design with gradient backgrounds
- Interactive hover effects
- Links to popular websites (Google, GitHub, YouTube, etc.)
- Development resources (MDN, W3Schools, CodePen)
- All external links open in new tabs
- Mobile-friendly responsive layout

## Setup and Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Installation Steps

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the Flask application:**
   ```bash
   python app.py
   ```

3. **Open your browser:**
   - Go to `http://127.0.0.1:5000` or `http://localhost:5000`
   - The navigation hub will be displayed

## Project Structure

```
interface/
├── app.py                 # Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main HTML template
├── static/
│   └── styles.css        # CSS styles
└── README.md             # This file
```

## Usage

1. Start the Flask server by running `python app.py`
2. Open your web browser and navigate to `http://127.0.0.1:5000`
3. Click on any navigation card to jump to the respective website
4. All external links will open in new tabs

## Stopping the Server

Press `Ctrl+C` in the terminal to stop the Flask development server.

## Development

The Flask application runs in debug mode by default, which means:
- Automatic reloading when you make changes to the code
- Detailed error messages in the browser
- Interactive debugger for troubleshooting

## Customization

You can easily customize the website by:
- Adding new links in `templates/index.html`
- Modifying styles in `static/styles.css`
- Adding new routes in `app.py`
