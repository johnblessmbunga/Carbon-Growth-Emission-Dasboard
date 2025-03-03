# Carbon-Growth-Emission-Dasboard
Carbon-Growth-Emission-Dasboard is a data analysis dashboard for carbon growth emissions in different regions and years.
##Table of contents
1. [Installation](#installation)
2. [Usage](#usage)
4. [License](#license)
## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/johnblessmbunga/Carbon-Growth-Emission-Dasboard.git
   cd Carbon-Growth-Emission-Dasboard
   npm install
## Usage
To strat application ,run:

npm start
### Features
__-Dashboard__: A dashboard is used for user input. Users can choose whether to analyse carbon growth emissions in a particular year or year range and in a particular region or country as shown in Figure 1. 
### Figure 1: Dasboard
![Dashboard](images/Dashboard.png)


__-Write and Save Report__: Writes a report table of different parameters regarding the carbon growth emissions in the selected time frame and region when the button enter is pressed as shown in Figure 2.
### Figure 2 : Dashboard with Report
![Dashboard_Report](images/Dashboard_with_Report.png)
The parameters vary depending on selection used. The common parameters are region/country, year/year range, and total growth emissions. If year range is used the average, maximum , and minimum annual carbon growth emissions are calculated and presented. If region is selected the average, maximum, and minimum national carbon growth emissions are calculated and presented. Additionally, the width of the table changes depending on string length.

__-Save Report__: The report can be saved into an excel file by clicking the save button. The filename will be generated based on time frame and region selected.

__-Error Handling__: The calculations of the carbon growth emission parameters only begin when enter button pressed and all necessary inputs are valid.
## License
This project is licensed under MIT license.

### Acknowledgements
Thanks to the Kagglehub, os, pandas, tkinter, and openpyxl  library for providing backend framework. 
Addtionally I extend my thanks to Samith Chimminiyan for creating and sharing dataset on Kagglehub
