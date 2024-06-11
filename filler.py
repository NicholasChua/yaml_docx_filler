"""
This Python code is used to fill a Word document template with the extracted information from a YAML file.

This is the assumed structure of the document template:
- Document Header and Footer Items
    - Header
        - Document Type
        - Document Number
        - Document Revision Number
        - Document Effective Date
        - Document Title
    - Footer
        - Document Code
- Document Control Items
    - Revision History
    - Document Review and Approval
        - Prepared By
        - Reviewed and Approved By
- Document Body Items
    - Purpose           (1.0)
    - Scope             (2.0)
    - Responsibility    (3.0)
    - Definition        (4.0)
    - Procedure         (5.0)
    - Reference         (6.0)
    - Attachment        (7.0)
"""

import os
from docxtpl import DocxTemplate
import yaml
from typing import Dict, List, Union, Optional, Any


# Define type aliases for known sections
RevisionHistory = [List[Dict[str, str]]]
PreparedBy = [List[Dict[str, str]]]
ReviewedApprovedBy = [List[Dict[str, str]]]
Purpose = [List[str]]
Scope = [List[str]]
Responsibility = [List[str]]
Definition = [List[str]]

# Define a generic type alias for Procedure subsections
GenericProcedureSubSection = Optional[
    Union[List[str], List[Dict[str, Any]], Dict[str, Any]]
]

# Define the type alias for Procedure section
Procedure = [Dict[str, GenericProcedureSubSection]]

# Define the overall document type
DocumentType = Dict[
    str,
    Union[
        str,
        RevisionHistory,
        PreparedBy,
        ReviewedApprovedBy,
        Purpose,
        Scope,
        Responsibility,
        Definition,
        Procedure,
        List[str],
    ],
]


def read_yaml_file(input_file: str) -> DocumentType:
    """
    Used to read a YAML file and return the content as a dictionary
    Usage: content = read_yaml_file('text.yml')

    Inputs:
    input_file (str): The file path of the YAML file to be read

    Returns:
    yaml_content (dict): The content of the YAML file as a dictionary
    """
    with open(input_file, "r", encoding="utf-8") as file:
        yaml_content = yaml.safe_load(file)
    return yaml_content


def fill_common_items(yaml_content: DocumentType) -> dict:
    """
    Base function to read YAML data and return common items in a document template as a context dictionary.
    Extended functions will add on to the context dictionary and fill the document template.
    This is the assumed structure of the document template:
    - Document Header and Footer Items
        - Header
            - Document Type
            - Document Number
            - Document Revision Number
            - Document Effective Date
            - Document Title
        - Footer
            - Document Code
    - Document Control Items
        - Revision History
        - Document Review and Approval
            - Prepared By
            - Reviewed and Approved By
    - Document Body Items
        - Purpose           (1.0)
        - Scope             (2.0)
        - Responsibility    (3.0)
        - Definition        (4.0)
        - Reference         (6.0)
        - Attachment        (7.0)

    Usage: fill_common_items(content)

    Inputs:
    yaml_content (dict): The extracted information from the YAML file from read_yaml_file()

    Returns:
    context (dict): The context dictionary with the common items filled in
    """
    context = {
        # Document Header and Footer Items
        "type": yaml_content["type"],
        "document_no": yaml_content["document_no"],
        "document_code": yaml_content["document_code"],
        "effective_date": yaml_content["effective_date"],
        "document_rev": yaml_content["document_rev"],
        "title": yaml_content["title"],
        # Document Control Items
        "revision_history": yaml_content["revision_history"],
        "prepared_by": yaml_content["prepared_by"],
        "reviewed_approved_by": yaml_content["reviewed_approved_by"],
        # Document Body Items
        "purpose": yaml_content["purpose"],
        "scope": yaml_content["scope"],
        "responsibility": yaml_content["responsibility"],
        "definition": yaml_content["definition"],
        # Procedure (5.0) goes here, filled in by extended functions
        "reference": yaml_content["reference"],
        "attachment": yaml_content["attachment"],
    }

    return context


def get_mixed_types_str_keys(
    content: Union[List[Any], Dict[Any, Any]]
) -> Union[List[str], str]:
    """
    Helper function to extract a list of strings and dictionary keys from a list or dictionary of mixed strings and/or dictionaries.
    Usage: get_mixed_types_str_keys(content['list']['sub-list']) or get_mixed_types_str_keys(content['dictionary'])

    Inputs:
    content (list, dict): The list or dictionary containing mixed strings and/or dictionaries

    Returns:
    mixed_str_key_list (list, str): A list of strings and dictionary keys. If there is only one item, it is returned as a string instead of a list.
    """
    mixed_str_key_list = []
    # Interpret the content as a list if it is a dictionary
    if isinstance(content, dict):
        content = list(content.values())
    for item in content:
        # If the item is a string, strip the trailing newline and add it to the list
        if isinstance(item, str):
            item = item.rstrip("\n")
            mixed_str_key_list.append(item)
        # Else if the item is a dictionary, add the key to the list
        elif isinstance(item, dict):
            mixed_str_key_list.append(list(item.keys())[0])
        # Default case handling
        else:
            mixed_str_key_list.append("error here")

    # If there is only one item in the list, return it as a string.
    if len(mixed_str_key_list) == 1:
        return str(mixed_str_key_list[0])

    return mixed_str_key_list


def get_mixed_types_dict_values(content: List[Any]) -> List[str]:
    """
    Helper function that, given a list of mixed strings and dictionaries, extracts all dictionary values in the list in order.
    Returns a list of lists containing the values of the dictionaries in the input list.
    The inner lists contain the values of the dictionaries in the order they appear in the input list.
    Usage: get_mixed_types_dict_values(content['list']['sub-list'])

    Inputs:
    content (list): The list containing mixed strings and dictionaries

    Returns:
    dict_values (list): A list of lists containing the values of the dictionaries in the input list. If there is only one list within the list,
    it is returned as a single list instead of a list of lists.
    """
    dict_count = 0
    dict_values = []
    for item in content:
        if isinstance(item, dict):
            dict_count += 1
            temp_list = []
            for value in list(item.values())[0]:
                temp_list.append(value)
            dict_values.append(temp_list)

    # If no dictionaries are found in the list, return None
    if dict_count == 0:
        print("No dictionaries found in the list")
        return None

    # If there is only one list of dictionary values, return it as a single list
    if len(dict_values) == 1:
        return dict_values[0]

    return dict_values



def example_document_generation(
    yaml_content: DocumentType,
    input_file: str = os.path.join("example", "template_example.docx"),
    output_file: str = os.path.join("example", "filled_example.docx"),
) -> None:
    """
    Used to fill a Word document template with the extracted information.
    This is purely an example and should be customized based on the actual document template.
    Usage: process_document(content, input_file, output_file)

    Inputs:
    yaml_content (dict): The extracted information from the YAML file from read_yaml_file()
    input_file (str): The file path of the Word document template
    output_file (str): The file path to save the filled Word document

    Returns:
    None

    For the example document, the data structure is as follows:
    {'type': 'EXAMPLE DOCUMENT', 
    'document_no': 'EX01-100', 
    'document_code': 'EX01-100-01', 
    'effective_date': '06-JUN-2024', 
    'document_rev': '00', 
    'title': 'Example Document', 
    'revision_history': [{'revision_no': '00', 'description_of_changes': 'Initial release', 'effective_date': '06-JUN-2024'}], 
    'prepared_by': [{'name': 'Nicholas Chua', 'title': 'Preparer'}], 
    'reviewed_approved_by': [{'name': 'Not Nicholas Chua', 'title': 'Approver'}], 
    'purpose': ['Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'], 
    'scope': ['Lorem ipsum dolor sit.'], 
    'responsibility': ['Lorem ipsum,', 'dolor sit amet.'], 
    'definition': ['Lorem: ipsum,', 'Dolor: sit amet.'], 
    'procedure': {
    'example_policy': 
    [
        {
        'Lorem ipsum dolor sit amet:': ['consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.', 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.', 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.', 'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.']
        }, 
        {
        '"This is a very very very very long multi-line comment line 1.\nThis is a very very very very long multi-line comment line 2:"\n': ['This is a sub-item of the multi-line comment.']
        }, 
        'Lorem ipsum dolor sit amet.'
    ], 
    'another_policy': 
    [
        {
        'I have 9 sub items:': ['Item 1.', 'Item 2.', 'Item 3.', 'Item 4.', 'Item 5.', 'Item 6.', 'Item 7.', 'Item 8.', 'Item 9.']
        }
    ]
    }, 
    'reference': ['N/A'], 
    'attachment': ['N/A']}
    """
    # Open the Word document
    doc = DocxTemplate(input_file)

    # Prepare the context for the template
    # Call the base function to fill in common items
    context = fill_common_items(yaml_content)

    # Fill in the Procedure section
    context.update(
        {
            # The hard part is in the jinja tags within the word document
            # See the example_template for the structure of the tags and filled_example for the filled tags
            "procedure": yaml_content["procedure"],
        }
    )

    # Replace template jinja tags with corresponding extracted information
    doc.render(context)

    # Export the modified Word document
    doc.save(output_file)


def main():
    print("You are now running the example filler.py script!")

    # Read the YAML file
    yaml_file_path = os.path.join("example", "example.yml")
    yaml_content = read_yaml_file(yaml_file_path)

    # Generate the filled document
    example_document_generation(yaml_content)


if __name__ == "__main__":
    main()
