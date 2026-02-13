# PRTG APC Sensor Analysis

Excel report generator for APC environmental sensor data from PRTG Network Monitor. Creates multi-sheet reports with statistical analysis, hourly/daily patterns, and threshold recommendations for data center temperature monitoring.

## Features

### Excel Report Generation
- **Professional Multi-Sheet Reports**: Summary + detailed stats + raw data per sensor
- **Dark Theme Styling**: Modern appearance with color-coded status indicators
- **Summary Dashboard**: Overview table of all sensors with current status
- **Detailed Statistics**: Per-sensor sheets with complete analysis
- **Raw Data Export**: Formatted tables with all datetime/temperature readings
- **Color-Coded Thresholds**: Red/orange/green indicators for temperature status

### Analysis Capabilities
- **APC Sensor Support**: Designed for APC environmental sensors (AP9335T, AP9335TH, NetBotz)
- **Historical Data Retrieval**: Fetches temperature sensor data from PRTG API
- **Statistical Analysis**: Mean, min, max, standard deviation, percentiles
- **Pattern Detection**: Hourly and daily temperature patterns
- **Trend Analysis**: 24-hour temperature trends
- **Threshold Suggestions**: Preliminary warning/error threshold recommendations
- **Multiple Sensor Support**: Analyze multiple APC sensors in one run

## Use Case

Ideal for data center operators using APC environmental monitoring equipment who need to:
- Establish baseline temperature behavior for APC sensors in server rooms and data centers
- Identify temperature patterns and anomalies in APC-monitored environments
- Set appropriate alert thresholds based on actual APC sensor data
- Monitor cooling system effectiveness with APC environmental sensors
- Document environmental conditions for compliance and capacity planning
- Analyze APC NetBotz, AP series sensor data

## Requirements

- Python 3.7+
- PRTG Network Monitor with API access
- Valid PRTG credentials with sensor read permissions

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/prtg-apc-sensor-analysis.git
cd prtg-apc-sensor-analysis
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Create your configuration file:
```bash
cp config.json.example config.json
```

4. Edit `config.json` with your PRTG details:
```json
{
  "prtg": {
    "url": "https://your-prtg-server.com",
    "username": "your_username",
    "password": "your_password"
  },
  "sensors": {
    "12345": "APC Sensor - Server Room A (AP9335T)",
    "12346": "APC NetBotz - Data Center Aisle Hot Spot"
  },
  "days_to_analyze": 7
}
```

## Configuration

### PRTG Settings

- `url`: Your PRTG server URL
- `username`: PRTG username with API access
- `password`: PRTG password 

### Sensor Configuration

- Add APC sensor IDs and descriptive names to the `sensors` object
- To find sensor IDs: In PRTG, navigate to the APC sensor → the ID is in the URL

### Analysis Period

- `days_to_analyze`: Number of days to analyze (recommended: 7-14 for baseline)
- Script uses 5-minute averages for analysis

## Usage

Generate an Excel report with all configured sensors:

```bash
python prtg_report.py
```

Output: `sensor_report_{timestamp}.xlsx`

### Sample Output

```
PRTG SENSOR REPORT
======================================================================

[APC Sensor - Server Room A (AP9335T)] (Sensor 12345)
  Fetching data (last 7 days)...
  ✓ Retrieved (45829 bytes)
  ✓ Column: 'Temperature (°C)'
  ✓ 2016 valid readings
  Current: 22.3°C  |  Avg: 22.5°C  |  Range: 21.1-23.9°C  |  VERY STABLE

──────────────────────────────────────────────────────────────────────
Generating Excel report...

✓ Report saved: sensor_report_20260213_1430.xlsx
  2 sensor(s) included.
======================================================================
```

### Excel Report Structure

The Excel report includes multiple tabs:

**Summary Tab:**
```
PRTG Temperature Sensor Report
February 13, 2026  14:30

Sensor                          Current  Avg   Min   Max   Range  Std Dev  Stability      Upper Err  Upper Warn  Lower Warn  Lower Err
APC Sensor - Room A (12345)     22.3°C   22.5  21.1  23.9  2.8    0.58     VERY STABLE    23.8°C     23.5°C      21.5°C      21.3°C
APC NetBotz - Aisle Hot (12346) 24.1°C   24.0  22.8  25.2  2.4    0.62     VERY STABLE    25.0°C     24.8°C      23.2°C      23.0°C
```

**Stats Tabs:**
- Current temperature, average, min, max, range, std dev
- Stability classification
- Full percentile distribution (1st, 5th, 25th, 50th, 75th, 95th, 99th)
- Threshold recommendations (Upper/Lower Error/Warning)
- Hourly averages table with min/max per hour

**Raw Data Tabs:**
- Complete DateTime and Temperature data
- Formatted tables with professional styling

## Output Files

The script generates an Excel workbook:

**File:** `sensor_report_{timestamp}.xlsx`

**Sheet Structure:**
- **Summary Sheet**: Overview table of all sensors with current status
- **Stats Sheets**: One per sensor with detailed statistics, percentiles, and hourly patterns
- **Raw Data Sheets**: One per sensor with complete datetime/temperature data

## Recommendations

### Data Collection Period
- **Minimum**: 2-3 days for quick analysis
- **Recommended**: 7-14 days for baseline establishment
- **Seasonal**: 30+ days to capture weekly patterns
