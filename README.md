# Kyocera MFP Document Box Transfer Automation

## Overview

This script automates the transfer of documents to document boxes on Kyocera MFPs, eliminating the need for manual printing and renaming. By setting up the correct environment, users can efficiently send multiple documents to specific document boxes without manual intervention.

## Implementation

Currently only the .py file in /python, and the .exe in /python/dist are complete enough to be run successfully. May or may not complete the .ps1 script in /ps1. Plan to create a separate .py file to allow for python script to run by itself and be separate from the script used to compile the .exe.

## Prerequisites

Before running the script, ensure the following setup is completed:

1) **Install and Configure a Print Driver**
   - A new print driver must be added that uses a FILE port (this is critical for the script to function correctly).

2) **Network Access**
   - The workstation must have network access to the Kyocera MFP that will receive the documents. This can be achieved via:
     - Standard network connection
     - Direct crossover cable connection

3) **Folder Structure**
   - A main directory should be created (e.g., boxes).
   - Inside this directory, subfolders must be named according to the **document box numbers** on the Kyocera MFP.
   - Documents to be transferred should be placed inside their respective subfolders.
   - Document box numbers can be found via:
     - The machine's display panel
     - Kyocera NetViewer (easiest)
     - Web UI of the MFP

## Usage Instructions

1) Run the script.

2) The script will prompt for three inputs:
   - The name of the printer using the **FILE port**.
   - The **IP address** of the Kyocera MFP.
   - The **directory** containing the structured folders with documents.

3) Once inputs are provided, the script will:
   - Process each document and send it to the corresponding document box.
   - Run automatically without requiring further user interaction.

## Important Notes
- **Do not interrupt the script while it is running**.
- **Do not interact with the workstation during execution**, as this may disrupt GUI automation aspects of the script.
- Ensure that all required files are correctly placed in their respective subfolders before starting the script.

## Future enhancements
- Better error tracking, handling, and edge case testing
- Optional GUI for easier setup of environment
- Optional GUI to setup script
- Optional CLI argument passing