# Spec Reviewer
Simple Python script for spec reviews. 

The intention is to quickly identify and record minor issues in a specification file.

The script detects the following: 
 - Empty cells 
 - Absence of formatting comments when the content is center aligned
 - Issues with the CDASH column content - redundant spaces, missing hyphen, use of a dash instead of a hyphen
 - Line breaks in cells/paragraphs

 User can also choose to skip some of the checks by adding one or more flags when running the script.


## Getting Started

To get a local copy clone the repo: 
```sh
   git clone https://github.com/iburazin-ag/spec-reviewer.git
```


### Prerequisites

To be able to run this script, be sure to check if you have Python3 installed on your machine by running the following command in the terminal:
```sh
   python3 --version
```

If you get the following error:
```sh
    Error: No developer tools installed.
    Install the Command Line Tools:
    xcode-select --install
```
Run the command mentioned above:
```sh
   xcode-select --install
```
After the installation is complete, check the version of Python3 again, and if there are no errors, run the command below:
```sh
   pip3 install python-docx
```

## Running the scripts

Make sure you are in the same directory where the script is, and run the following command in the terminal:
```sh
   python3 spec-reviwer.py <file_local_path>
```
If you want to skip some of the checks, you can use optional flags like in the example below:
```sh
   python3 spec-reviwer.py <file_local_path> --skip-line-breaks 
```
Run the script with the --help flag for the list and a brief description of the arguments.

**NOTES:** 
 - For easier use, make sure that the document you want to run the script on is in the same directory as the script, that way you'll only have to provide the full name of the document
 - There should be no spaces in the title of the document.


If there are no findings, the script will output the following in the terminal:
```sh
   No findings found!
```

Otherwise, the following message will be displayed and the file scanned will be opened:
```sh
   Findings recorded in the file.
```
The file will contain all the findings in the cells where they occurred. The findings are in all caps red letters for better visibility.

Once the file has been updated and saved, run the script again. If there are any residual comments left, that will be treated as a finding. 