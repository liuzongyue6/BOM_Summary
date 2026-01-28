# BOM Processor - Deployment Guide

**Quick deployment guide for the BOM Processor Web Application**

---

## 📋 Quick Start

### Local Setup (5 minutes)

**Windows:**
```powershell
python -m venv bom_streamlit
.\bom_streamlit\Scripts\Activate.ps1
pip install -r requirements.txt
run_app.bat
```

**Linux/Mac:**
```bash
python3 -m venv bom_streamlit
source bom_streamlit/bin/activate
pip install -r requirements.txt
./run_app.sh
```

**Access:**
- Local: http://localhost:8501
- Network: http://YOUR_IP:8501

---

## 🖥️ VM Deployment

### Windows VM

```powershell
# 1. Install dependencies
cd C:\BOM_Summary
python -m venv bom_streamlit
.\bom_streamlit\Scripts\activate.bat
pip install -r requirements.txt

# 2. Open firewall port
New-NetFirewallRule -DisplayName "BOM Streamlit" -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow

# 3. Run application
run_app.bat
```

### Linux VM

```bash
# 1. Install dependencies
sudo apt update && sudo apt install python3 python3-pip python3-venv -y
cd /opt/BOM_Summary
python3 -m venv bom_streamlit
source bom_streamlit/bin/activate
pip install -r requirements.txt

# 2. Open firewall port
sudo ufw allow 8501/tcp

# 3. Run application
./run_app.sh
```

**Access from other machines:** `http://VM_IP:8501`

---

## ⚙️ Configuration

### Basic Settings (`streamlit_app/config.py`)

```python
MAX_UPLOAD_SIZE_MB = 100      # Maximum file size
CLEANUP_HOURS = 24            # Auto-cleanup interval
```

### Advanced Settings (`.streamlit/config.toml`)

```toml
[server]
port = 8501
address = "0.0.0.0"
maxUploadSize = 100
```

---

## 🚀 Auto-Start Service

### Windows Service (NSSM)

```cmd
# Download NSSM from https://nssm.cc/download
nssm install BOMStreamlitApp "C:\BOM_Summary\bom_streamlit\Scripts\python.exe" "-m streamlit run C:\BOM_Summary\streamlit_app\bom_app.py --server.port=8501 --server.address=0.0.0.0"
nssm set BOMStreamlitApp AppDirectory "C:\BOM_Summary"
nssm set BOMStreamlitApp Start SERVICE_AUTO_START
nssm start BOMStreamlitApp
```

**Manage service:**
```cmd
nssm status BOMStreamlitApp    # Check status
nssm restart BOMStreamlitApp   # Restart
nssm stop BOMStreamlitApp      # Stop
```

### Linux Service (systemd)

Create `/etc/systemd/system/bom-streamlit.service`:

```ini
[Unit]
Description=BOM Streamlit Application
After=network.target

[Service]
Type=simple
User=YOUR_USERNAME
WorkingDirectory=/opt/BOM_Summary
Environment="PATH=/opt/BOM_Summary/bom_streamlit/bin"
ExecStart=/opt/BOM_Summary/bom_streamlit/bin/streamlit run /opt/BOM_Summary/streamlit_app/bom_app.py --server.port=8501 --server.address=0.0.0.0
Restart=always

[Install]
WantedBy=multi-user.target
```

**Enable and start:**
```bash
sudo systemctl daemon-reload
sudo systemctl enable bom-streamlit
sudo systemctl start bom-streamlit
```

**Manage service:**
```bash
sudo systemctl status bom-streamlit   # Check status
sudo systemctl restart bom-streamlit  # Restart
sudo systemctl stop bom-streamlit     # Stop
```

---

## 🔧 Troubleshooting

### Port Already in Use

**Windows:**
```cmd
netstat -ano | findstr :8501
taskkill /PID <PID> /F
```

**Linux:**
```bash
sudo lsof -i :8501
sudo kill -9 <PID>
```

### Cannot Access from Network

**Checklist:**
- ✅ Firewall allows port 8501
- ✅ Application uses `--server.address=0.0.0.0`
- ✅ VM network mode is Bridge or NAT with port forwarding
- ✅ Correct VM IP address

### Module Not Found

```bash
# Make sure virtual environment is activated
pip install -r requirements.txt
```

### Get VM IP Address

**Windows:** `ipconfig`  
**Linux:** `ip addr show` or `hostname -I`

---

## 🛡️ Security (Optional)

### IP Whitelist

**Windows:**
```powershell
New-NetFirewallRule -DisplayName "BOM Restricted" -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow -RemoteAddress 192.168.1.0/24
```

**Linux:**
```bash
sudo ufw allow from 192.168.1.0/24 to any port 8501
```

### Add Authentication

```bash
pip install streamlit-authenticator
```

### Use HTTPS

Configure Nginx reverse proxy (see Streamlit documentation)

---

## 📊 System Requirements

- Python 3.8+
- 2GB RAM minimum (4GB recommended)
- 500MB disk space
- Windows 10+, Linux (Ubuntu 20.04+), or macOS 10.14+

---

## 📖 Usage

1. **Upload File** - Browse or drag Excel file (.xlsm/.xlsx, max 100MB)
2. **Validate** - System checks required columns automatically
3. **Process** - Click "Start Processing" and view real-time logs
4. **Download** - Get three output files: extracted workbook, BOM summary, and processing log

---

**Version:** 1.0 | **Updated:** January 2026
