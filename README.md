# AssayScreeningDataProcessor
Processes assay screening data outputted from Perkin-Elmer Envision plate readers for compatibility to upload into a database for further analyses and comparisons across experiments.

This script is intended to provide users with an easy-to-use interface and to automate processing and converting raw data from plate readers amenable to upload to a database.

This processor also provides a plate summary spreadsheet that summarizes statistical information: high/low control averages, standard deviations, z' scores, etc. to help assess assay quality across all plates run. The processor was intended for use in high-throughput screening assays.

# Instructions
1. Under "Browse to raw screening data folder" use the "Browse" button to navigate to your Envision plate reader results files (.csv).
2. Under "Browse to the chemical database file" use the "Browse" button to select the chemical database file (ex. 20220419_Chemical_Database_Example.xlsx).
3. Select the assay plate format either 1536-well or 384-well plates.
4. Check at least one box of any of the chemical classes used in the experiment (ex. XA, XB, etc.).
5. Under the "Enter metadata..." section enter experiment details.
6. Under "Enter any assay plate numbers you'd like to exclude:" enter integers 1-27 that represent the assay plate ID numbers you'd like to exclude from data processing.
7. Click the "Submit" button after making your choices. Under the "Status:" window a message will notify you when data processing completes.
8. Exit by clicking the 'X' in the upper-right corner or by using the 'Cancel' button.
