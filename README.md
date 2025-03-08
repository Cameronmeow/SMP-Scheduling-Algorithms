# Scheduling Algorithm Steps

## Overview

This project automates the scheduling process using Python. It consists of three main scripts that preprocess data, execute scheduling logic, and send automated email notifications.

### Step 1 - Prerequistes
Ensure you have Visual Studio Code and Python installed. You can follow this guide: [Install VS Code and Python](https://youtu.be/D2cwvpJSBX4?si=YaYW3ITyThKGQ82D).

### Step 2 - Installation
To run this project, install the required dependencies by executing the following commands:

``` python
pip install pandas openpyxl smtplib ssl logging
```

```
import pandas as pd
from datetime import datetime, timedelta
import logging
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
```


## Step 3

Run the files in the following order:

- 1. Preprocessing.py

- - Reads the form responses file.
- - Cleans and processes the input data.
- - Ensures data consistency and format compatibility.
- - Prepares data for scheduling.
- - Test with colors.xlsx file is the output file.

- 2. Scheduling.py

- - Reads the Test with colours
- - Implements the scheduling algorithm.
- - Assigns tasks based on predefined rules.
- - Generates an output file with the final schedule - Weekly Interview Schedule.

- 3. AutoEmail.py

- - Reads the generated schedule.
- - Composes and sends email notifications to relevant stakeholders.
- - Ensures secure email delivery using SSL.