# Nium Capability Explorer 🌍

A modern, interactive web application built with Streamlit that allows users to explore Nium's global payment capabilities, filter by multiple criteria, preview detailed information, and download customized reports.

---

## 📋 Features

✅ **Dual Dataset Support**
- Toggle between FI (Financial Institution) and Non-FI datasets
- 923+ total payment corridors across 223 countries

✅ **Advanced Filtering**
- Filter by Country (with "All" option)
- Filter by Payment Mode (Bank Account, SWIFT, UPI, Visa Direct, etc.)
- Filter by Currency (Global currencies supported)
- Filter by TAT (Turn-Around Time: T0, T1, T2, Realtime)
- Multi-select for Transaction Types (B2B, B2P, P2P, P2B)
- Smart "All" option for each filter

✅ **Interactive Results**
- Results summary with metrics (Total results, Countries, Payment Modes, Currencies)
- Clean, sortable data table with key information
- Real-time result count updates

✅ **Detailed Preview**
- Expandable rows to view all 24 data fields
- Organized sections:
  - Core Information
  - Requirements & Documentation
  - Transaction Limits (per transaction type)
  - Additional Information & Notes

✅ **Smart Download**
- User-selectable columns for export
- Organized by category (Core Info, Transaction Limits, Requirements, All)
- Export formats: CSV or Excel
- Automatic filename with timestamp
- Only download selected data

✅ **Professional UI**
- Modern, responsive design
- Color-coded sections
- Intuitive navigation
- Mobile-friendly interface

---

## 🛠️ Installation & Setup

### **Prerequisites**
- Python 3.8 or higher
- pip (Python package manager)

### **Step 1: Install Python Packages**

```bash
pip install -r requirements.txt
```

### **Step 2: Prepare Your Data Files**

Place the Excel files in the same directory as the script:
```
.
├── nium_capability_explorer.py
├── 11th_April_FI_Payout_data.xlsx
├── 11th_April_Non_FI_Data.xlsx
└── requirements.txt
```

**Note:** The app currently expects files at:
```
/mnt/user-data/uploads/11th_April_FI_Payout_data.xlsx
/mnt/user-data/uploads/11th_April_Non_FI_Data.xlsx
```

To change the file paths, edit lines 70-71 in `nium_capability_explorer.py`:
```python
fi_path = '/path/to/your/FI_file.xlsx'
non_fi_path = '/path/to/your/NonFI_file.xlsx'
```

### **Step 3: Run the App Locally**

```bash
streamlit run nium_capability_explorer.py
```

The app will open in your default browser at: `http://localhost:8501`

---

## 🚀 Usage Guide

### **Basic Workflow**

1. **Select Dataset**
   - Choose FI or Non-FI from the radio button on the top-left

2. **Apply Filters**
   - Select values from dropdowns (or leave as "All" for no filter)
   - Select transaction types as needed (B2B, B2P, P2P, P2B)
   - Click "🔍 Apply Filters" button

3. **View Results**
   - See summary metrics at the top
   - Scroll through the results table
   - Check "Total matches" to understand your filter effectiveness

4. **Preview Details**
   - Click any row expander (🔹) to see all 24 data fields
   - View transaction limits for all types
   - Check requirements and documentation

5. **Download Data**
   - Select columns you want to include
   - Choose column tabs: Core Info, Transaction Limits, Requirements, or All
   - Click "📥 Download as CSV" or "📥 Download as Excel"
   - File downloads with timestamp

### **Filter Logic**

- **OR Logic within categories**: Select multiple values in a single filter? (Not currently available, but can be added)
- **AND Logic across categories**: Country + Payment Mode + Currency = must match ALL conditions
- **Supported Modes**: Only shows rows where ALL selected transaction types are supported
- **"All" option**: Returns results for every value in that dimension

### **Example Scenarios**

#### Scenario 1: "Show me all corridors to India"
1. Select Country: "India"
2. Leave Payment Mode: "All"
3. Leave Currency: "All"
4. Leave TAT: "All"
5. Click Apply Filters
→ Shows all payment corridors available for India

#### Scenario 2: "What B2B transaction corridors support Realtime TAT?"
1. Leave Country: "All"
2. Leave Payment Mode: "All"
3. Leave Currency: "All"
4. Select TAT: "Realtime"
5. Check only "B2B" transaction type
6. Click Apply Filters
→ Shows only B2B corridors with Realtime delivery

#### Scenario 3: "India to USA via UPI in INR with transaction limits"
1. Select Country: "India"
2. Select Payment Mode: "Proxy (UPI)"
3. Select Currency: "INR"
4. Leave TAT: "All"
5. Click Apply Filters
6. For download: Select "Core Info" + "Transaction Limits" tabs
7. Download as Excel

---

## 📊 Data Fields Explained

### **Core Fields**
- **Country**: Destination country for payout
- **Payment Mode**: Method of payment (Bank Transfer, Card, Wallet, etc.)
- **Currency**: Currency of the transaction
- **TAT**: Time-to-settlement
  - `T0` = Same day
  - `T1` = Next business day
  - `T2` = 2 business days
  - `Realtime` = Instant

### **Transaction Type Fields**
- **B2B**: Business to Business payments
- **B2P**: Business to Person (Employee payouts, vendors, etc.)
- **P2P**: Person to Person (Remittances, transfers)
- **P2B**: Person to Business (Marketplace payments)

### **Limits Fields**
- **Transaction limit per end-user - [TYPE] - Min**: Minimum amount per transaction
- **Transaction limit per end-user - [TYPE] - Max**: Maximum amount per transaction

### **Other Fields**
- **Supported Modes**: Which transaction types are available
- **Network Participant**: Payment network operator
- **Channels**: Available channels/methods
- **Mandatory data requirements**: Required information for setup
- **Supporting Documents**: Documentation needed
- **Cutoff & delivery timing**: Details on cutoff times and delivery schedules
- **Notes**: Additional important information

---

## ⚙️ Configuration & Customization

### **Change File Paths**
Edit the `load_data()` function (lines 70-76):
```python
def load_data():
    """Load both FI and Non-FI datasets"""
    fi_path = 'path/to/your/fi_file.xlsx'
    non_fi_path = 'path/to/your/non_fi_file.xlsx'
    # ... rest of function
```

### **Modify Display Columns**
Edit the `display_columns` variable (line 188):
```python
display_columns = ['Country', 'Payment Mode', 'Currency', 'TAT', 
                  'Supported Modes', 'Network Participant',
                  'Transaction limit per end-user - B2B - Max',
                  'Transaction limit per end-user - B2P - Max',
                  'Transaction limit per end-user - P2P - Max',
                  'Transaction limit per end-user - P2B - Max']
```

### **Change Color Scheme**
Edit the CSS in the `st.markdown()` section at the top:
```python
# Change primary color from #1f77b4 to your brand color
h1 { color: #YOUR_HEX_COLOR; }
```

### **Adjust Table Height**
Edit line 191:
```python
st.dataframe(..., height=400)  # Change 400 to your preferred height
```

---

## 🚢 Deployment Options

### **Option 1: Local Network Share**
```bash
streamlit run nium_capability_explorer.py
# Share the http://localhost:8501 URL within your network
```

### **Option 2: Streamlit Cloud (Recommended for teams)**
1. Push your code to GitHub
2. Go to https://share.streamlit.io/
3. Connect your GitHub repo
4. Select your app file
5. Deploy in one click
6. Share the public URL with your team

### **Option 3: Docker Container**
Create a `Dockerfile`:
```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8501
CMD streamlit run nium_capability_explorer.py --server.port=8501 --server.address=0.0.0.0
```

Build and run:
```bash
docker build -t nium-explorer .
docker run -p 8501:8501 nium-explorer
```

### **Option 4: AWS/Azure/Google Cloud**
All cloud providers support Streamlit apps. Deploy using their CLI tools or use services like:
- **AWS**: Streamlit on EC2 or App Runner
- **Azure**: Streamlit on App Service
- **Google Cloud**: Streamlit on Cloud Run

---

## 🔄 Updating Data

### **Method 1: Replace Excel Files**
Simply replace the old Excel files with updated ones:
```bash
# Replace these files:
/mnt/user-data/uploads/11th_April_FI_Payout_data.xlsx
/mnt/user-data/uploads/11th_April_Non_FI_Data.xlsx
```

The app has data caching with `@st.cache_data` decorator. To refresh:
1. Press `R` in the app
2. Or stop and restart the Streamlit server

### **Method 2: Clear Cache Programmatically**
Add a button in the sidebar to clear cache:
```python
if st.sidebar.button("Clear Cache"):
    st.cache_data.clear()
    st.rerun()
```

---

## 🐛 Troubleshooting

### **Issue: "No module named 'streamlit'"**
```bash
pip install -r requirements.txt
```

### **Issue: "File not found" error**
Check that your Excel files are in the correct path specified in the code. Update paths in `load_data()` function if needed.

### **Issue: Filters not working**
Try clearing the Streamlit cache:
1. Press `C` in the app
2. Or restart the server: `Ctrl+C` and run again

### **Issue: Download button not showing**
Ensure at least one column is selected in the download section. Multiple tabs need to have checkboxes checked.

### **Issue: App is slow**
- First load may be slow due to Excel file reading. This is cached for subsequent runs.
- If using large files, consider splitting data or optimizing Excel format.

---

## 📈 Performance Notes

- **Data Loading**: ~2-3 seconds on first run (cached after that)
- **Filtering**: Real-time, instant response for < 1000 rows
- **Download**: Instant for all sizes
- **Memory Usage**: Minimal (~50MB for current data)

For datasets > 10,000 rows, consider:
- Using a database instead of Excel
- Pre-filtering data before export
- Implementing pagination

---

## 🔐 Security & Privacy

- **No data is sent externally** (runs locally)
- **No user accounts required**
- **Data stored only in Excel files** (as provided)
- **Downloads are local** (stored on user's computer)

---

## 📝 License & Support

This app is created for Nium internal use. For support or feature requests, contact your development team.

---

## 🎨 Future Enhancements

Potential features for future versions:

- [ ] Database backend instead of Excel
- [ ] User authentication & role-based access
- [ ] Save/share filter presets
- [ ] Compare FI vs Non-FI side-by-side
- [ ] Matrix/heatmap view (Countries vs Payment Modes)
- [ ] Advanced search with regex
- [ ] Scheduled data updates from backend
- [ ] API endpoint for programmatic access
- [ ] Dark mode toggle
- [ ] Multi-language support
- [ ] Email export functionality

---

## 📞 Quick Reference

| Action | How-to |
|--------|--------|
| **Start app** | `streamlit run nium_capability_explorer.py` |
| **Clear cache** | Press `C` in app |
| **Reset filters** | Click "🔄 Reset All" button |
| **Update data** | Replace Excel files & restart app |
| **Download data** | Select columns → Click "📥 Download" button |
| **View all fields** | Click expander (🔹) on any row |
| **Export settings** | Manually select columns each time (future: save presets) |

---

**Version**: 1.0  
**Last Updated**: December 2024  
**Python**: 3.8+  
**Dependencies**: Streamlit 1.28.1, Pandas 2.0.3, OpenPyXL 3.1.2
