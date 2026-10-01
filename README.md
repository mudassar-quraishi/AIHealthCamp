# UiPath Health Camp Registration Workflow

## Overview

This UiPath workflow reads pending registration requests (JSON files), processes them, and produces:
1. A unique Registration ID
2. A new row in `registrations.xlsx`
3. A confirmation text file
4. A result JSON file for the FastAPI backend to read

## Prerequisites

- UiPath Studio installed
- UiPath.Excel.Activities package installed
- Access to the project's `data/` directory

## Folder Structure

```
data/
├── pending/          ← FastAPI writes registration request JSON files here
├── results/          ← UiPath writes result JSON files here
├── confirmations/    ← UiPath writes confirmation .txt files here
└── registrations.xlsx ← UiPath manages this Excel file
```

## Workflow Steps (Main.xaml)

### Step 1: Set Data Path
- Assign `dataPath` = Path to the project's `data/` folder
- Use an absolute path or a relative path from the UiPath project directory

### Step 2: Check for Pending Files
- Use **Directory.GetFiles** to list all `.json` files in `data/pending/`
- If no files found → Log "No pending registrations" → End

### Step 3: Process Each Pending File
For each JSON file found:

#### 3a: Read the JSON File
- Use **Read Text File** activity to read the file content
- Use **Deserialize JSON** activity to parse into a JObject

#### 3b: Extract Fields
- `request_id` = jObject("request_id").ToString
- `full_name` = jObject("full_name").ToString
- `age` = jObject("age").ToString
- `gender` = jObject("gender").ToString
- `mobile_number` = jObject("mobile_number").ToString
- `address` = jObject("address").ToString
- `health_service` = jObject("health_service").ToString

#### 3c: Validate (Basic)
- Check that `full_name` is not empty
- Check that `age` is not empty
- If validation fails → write error result → skip to next file

#### 3d: Generate Registration ID
Format: `HCYYYY####` (e.g., `HC20260001`)

Logic:
1. Check if `registrations.xlsx` exists
   - If not, create it with headers and start at `0001`
2. If it exists, read the last row
3. Extract the numeric part of the last Registration ID
4. Increment by 1
5. Format: `"HC" + Year + SequenceNumber.ToString("D4")`

#### 3e: Update Excel
- Use **Excel Application Scope** (or **Use Excel File**)
- **Write Row** at the end of the sheet:

| Column | Value |
|--------|-------|
| A | Registration ID |
| B | Full Name |
| C | Age |
| D | Gender |
| E | Mobile Number |
| F | Address |
| G | Health Service |
| H | Today's Date (yyyy-MM-dd) |
| I | "Confirmed" |

#### 3f: Write Confirmation File
- Create file at: `data/confirmations/{registration_id}.txt`
- Content:

```
HEALTH CAMP REGISTRATION CONFIRMATION
======================================

Registration ID: HC20260001
Name: Rahul Sharma
Age: 35
Gender: Male
Mobile: 9876543210
Address: Aligarh, Uttar Pradesh
Health Service: General Checkup
Registration Date: 2026-09-23
Status: Confirmed

Thank you for registering for the Health Camp!
```

#### 3g: Write Result JSON
- Create file at: `data/results/{request_id}.json`
- Content:

```json
{
    "request_id": "req_abc12345",
    "status": "success",
    "registration_id": "HC20260001",
    "full_name": "Rahul Sharma",
    "health_service": "General Checkup",
    "confirmation_file": "data/confirmations/HC20260001.txt"
}
```

#### 3h: Clean Up
- Delete or move the processed file from `data/pending/`

### Step 4: Log Completion
- Log "Registration processed successfully: {registration_id}"

## Error Handling

Wrap the main processing logic in a **Try/Catch** block:

**On Error:**
- Log the error message
- Write a failure result JSON:

```json
{
    "request_id": "req_abc12345",
    "status": "failed",
    "message": "Error description here"
}
```

## How to Run

1. Open UiPath Studio
2. Open the `HealthCampRegistration` project
3. Make sure the `dataPath` variable points to your project's `data/` folder
4. Click **Run** (or **Debug**)
5. The workflow will process any pending registration files

For the demo:
1. Complete the chat registration in the React frontend
2. Click "Confirm Registration"
3. The backend writes a pending JSON file
4. Run the UiPath workflow
5. The frontend automatically detects the result and shows the confirmation

## Creating the Excel File

The first time UiPath runs, it should create `registrations.xlsx` with these headers in Row 1:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| Registration ID | Full Name | Age | Gender | Mobile Number | Address | Health Service | Registration Date | Status |

## Tips

- Use **Append Range** or **Write Range** to add data to Excel
- Use `Now.ToString("yyyy-MM-dd")` for the registration date
- Use `Now.Year.ToString` for the year part of the Registration ID
- Use **Path.Combine** for building file paths
- Test with a single registration first before processing multiple
