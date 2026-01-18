# Automated Weighing System v2.0

A Flask-based web application for automated weighing and monitoring of cargo boxes at Tifra Fish facilities.

## Features

- **Real-time Weight Monitoring**: Live weight display from serial-connected scales
- **Automated Weighing**: Automatic weight validation and recording with configurable tolerances
- **Stability Detection**: Wind/movement detection with stability verification over 3-second windows
- **Data Export**: CSV export functionality for reports and analysis
- **Web Dashboard**: Responsive web interface for monitoring and control
- **Multi-Scale Support**: Framework for supporting multiple scales (currently configured for single scale)
- **Session Management**: Session tracking with weighing history and statistics

## System Requirements

- Windows OS (uses `winsound` for audio feedback)
- Python 3.7+
- Serial port connection to digital scale

## Installation

1. Clone the repository:
```bash
git clone https://github.com/karakoceila/automated-weighing-system.git
cd automated-weighing-system
```

2. Create a virtual environment:
```bash
python -m venv .venv
.venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

Edit the configuration section in `balance_web.py`:

```python
PORT_1 = "COM3"              # Serial port for the scale
BAUD = 9600                  # Baud rate
EXCEL_FILE = "poids_caisses.xlsx"  # Output Excel file

# Weight range (kg)
POIDS_MIN = 13.080
POIDS_MAX = 13.220

# Stability settings
STABLE_SECONDS = 3.0         # Duration to observe for stability
TOL_STABILITE = 0.005        # Wind tolerance (5g in this example)

# Empty detection
SEUIL_VIDE_KG = 0.20         # Threshold to detect empty scale
```

## Usage

### Run the application:

```bash
python balance_web.py
```

The web interface will be available at `http://localhost:5000`

### Environment variables:

- `SINGLE_SCALE`: Set to "1" or "2" to specify which scale to run
- `FLASK_PORT`: Specify Flask port (default: 5000)

```bash
set SINGLE_SCALE=1
set FLASK_PORT=5000
python balance_web.py
```

## Web Interface

### Main Dashboard
- Current weight display with live updates
- Number of items weighed in current session
- List of last 2 validated weights
- Session start timestamp

### Controls
- **Commencer / Réinitialiser la session**: Start or reset the current weighing session
- **Télécharger le rapport (CSV)**: Download session report as CSV

### Status Indicators
- Green LED: Serial connection OK
- Red blinking LED: Serial connection lost

## API Endpoints

### GET `/balance/<scale_id>`
Display the web interface for a specific scale.

### POST `/reset/<scale_id>`
Reset the session for a specific scale.

### GET `/csv/<scale_id>`
Download session data as CSV file.

### GET `/status/<scale_id>`
JSON endpoint returning current status and history.

Response:
```json
{
  "com_ok": true,
  "weight": 13.15,
  "status": "OK - enregistrée ✅",
  "history": [...]
}
```

## Weighing Workflow

1. **ATTENTE_CAISSE** (Waiting for box)
   - System waits for a box to be placed on the scale
   - Collects weight samples over 3 seconds
   - Verifies stability (amplitude ≤ 5g)

2. **Validation**
   - If weight is within range (13.080 - 13.220 kg), records to Excel
   - Plays beep sound on success
   - Displays status message

3. **ATTENTE_VIDE** (Waiting for removal)
   - Waits for box to be removed (weight < 0.20 kg)
   - Returns to waiting state

## Data Output

Weights are automatically recorded in `poids_caisses.xlsx` with columns:
- Date/Heure (Timestamp)
- Balance (Scale ID)
- Poids caisse (kg) (Weight)
- Port (Serial port)
- Plage (Acceptable range)
- Note (Status)

## Troubleshooting

### "COM OFF - vérifier câble/port"
- Check serial cable connection
- Verify correct COM port in configuration
- Ensure scale is powered on and in correct mode (ST output)

### Scale not reading
- Scale must output lines starting with "ST" (e.g., "ST 13.156 kg")
- Check scale output format matches regex pattern

### Files
- `balance_web.py`: Main Flask application
- `affichage.py`: Additional display utilities (if used)
- `Lancer_Pesee.bat`: Batch file to launch the application
- `static/`: Directory for images and static assets

## Support

**AgynTech Security Support**: 0554752037

## License

Proprietary - Client: Tifra Fish

## Author

AgynTech Security
