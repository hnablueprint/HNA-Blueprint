# Generating HNA Word reports

This folder contains the files and code to produce `Housing Needs Assessment` reports at municipal level.

Follow this instructions to set up the `Virtual Environment` and install the required packages to run this code.

Please, copy, paste and run the following code lines into Terminal.

## Set up

### Create Virtual Environment

```bash
python -m venv .venv

```

### Activate Virtual Environment

```bash
.venv\Scripts\activate
```

### Create folders

```bash
New-Item -Path "output" -ItemType Directory
```

### Install Packages

```bash
pip install -r scripts/requirements.txt
```

## To run

Add the `HNA_databooks` Excel files corresponding to the desired subdivisions to the `input` folder. These can be generated using the `generate_hna_excel` folder also available in this repository.

You will need to have Excel and Word in your computer.

Then, run the script `generate_word.py`. The script will automatically generate and save the Word report from every `HNA_databook` file added to the `input` folder into the `output` folder.