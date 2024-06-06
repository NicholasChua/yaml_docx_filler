# Project README

Welcome to the `yaml_docx_filler` project!

## Description

The `yaml_docx_filler` is a tool designed to fill content from YAML files into DOCX template files. It provides a simple and efficient way to automate the process of populating DOCX files with data from YAML files.

## How It Works

See the example folder for an example of a YAML template file and a DOCX template file. The YAML template file contains the data to be filled into the DOCX template file. The `yaml_docx_filler` tool reads the YAML template file and fills the DOCX template file with the data.

## Why I Built This

I wanted to fill out a bunch of policy documents with a similar structure but different content. I decided to build this tool to ease the process of diffing and updating the documents. YAML being both human-readable and machine-readable made it a good choice for the data source, and it also happens to be valid JSON, which can be built upon later. DOCX and PDF was the format that the documents were in, and would be expected to be in.

I wanted to learn how to use the python-docxtpl libraries, which relies on Jinja templating. I also wanted to approach this project taking advantage of Git's version control system, and potentially use it as a way to track changes to the documents, which is infinitely easier using text-based markup languages like YAML, as opposed to DOCX and PDF.

## Installation

#### Dependencies

- Python 3.12

#### Steps

To install `yaml_docx_filler`, follow these steps:

1. Clone the repository: `git clone https://github.com/NicholasChua/yaml_docx_filler`
2. Navigate to the project directory: `cd yaml_docx_filler`
3. Install the dependencies: `pip install -r requirements.txt`

## Usage

To use `yaml_docx_filler`, follow these steps:

1. Prepare your yml and docx template files.
2. Modify filter.py's example_document_generation function to point at your yml and docx template files. By default, it points at the example files in the example folder.
3. Run filter.py.
4. The filled docx file will be generated in the specified output path.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
