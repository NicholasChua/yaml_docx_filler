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
from typing import TypedDict, List, Dict, Any


# Define the types for the YAML content based on the assumed structure of the document template
class HeaderFooterItems(TypedDict):
    """
    Represents the structure of the header and footer items in a document.

    Fields:
    - document_type (str): The type of the document.
    - document_no (str): The document number.
    - document_code (str): The code identifying the document.
    - effective_date (str): The date the document becomes effective.
    - document_rev (str): The revision number of the document.
    - title (str): The title of the document.
    """

    document_type: str
    document_no: str
    document_code: str
    effective_date: str
    document_rev: str
    title: str


class RevisionHistoryItem(TypedDict):
    """
    Represents the structure of a revision history item in a document.

    Fields:
    - revision (str): The revision number.
    - date (str): The date of the revision.
    - description (str): The description of the changes made in the revision.
    """

    revision: str
    date: str
    description: str


class PreparedByItem(TypedDict, total=False):
    """
    Represents the structure of the Prepared By section in a document.

    Fields:
    - name (str): The name of the individual.
    - role (str): The role of the individual.
    - date (str | None): The date of preparation. Optional field as it may not be present in all documents.
    """

    name: str
    role: str
    date: str | None


class ReviewedApprovedByItem(TypedDict, total=False):
    """
    Represents the structure of the Reviewed and Approved By section in a document.

    Fields:
    - name (str): The name of the individual.
    - role (str): The role of the individual.
    - date (str | None): The date of review or approval. Optional field as it may not be present in all documents.
    """

    name: str
    role: str
    date: str | None


class ProcedureSection(TypedDict, total=False):
    """
    Represents the structure of the Procedure section (5.0) in a document.
    This is the part of the document that is most likely to vary significantly between documents,
    thus it is left flexible with optional fields, and custom handling may be required.
    Minimally, it can be expected to be a mixed list of lists and dictionaries.

    Fields:
    - title (str): The title of the procedure section.
    - content [List[str] | List[Dict[str, Any]] | Dict[str, Any]] | None: The content of the procedure section.
    """

    title: str
    content: List[str] | List[Dict[str, Any]] | Dict[str, Any] | None


class DocumentType(TypedDict):
    """
    Represents the structure of a standard document, including header/footer, document control items and content sections.

    Fields:
    - document_type (str): The type of the document.
    - document_no (str): The document number.
    - effective_date (str): The date the document becomes effective.
    - document_rev (str): The revision number of the document.
    - title (str): The title of the document.
    - document_code (str): The code identifying the document.
    - revision_history (List[RevisionHistoryItem]): A list of revision history items.
    - prepared_by (List[PreparedByItem]): A list of individuals who prepared the document.
    - reviewed_approved_by (List[ReviewedApprovedByItem]): A list of individuals who reviewed and approved the document.
    - purpose (List[str]): The purpose of the document.
    - scope (List[str]): The scope of the document.
    - responsibility (List[str]): The responsibilities outlined in the document.
    - definition (List[str]): Definitions of terms used in the document.
    - reference (List[str]): References to other documents.
    - attachment (List[str]): Attachments to the document.
    - procedure (List[ProcedureSection]): The procedural content of the document.
    """

    document_type: str
    document_no: str
    effective_date: str
    document_rev: str
    title: str
    document_code: str
    revision_history: List[RevisionHistoryItem]
    prepared_by: List[PreparedByItem]
    reviewed_approved_by: List[ReviewedApprovedByItem]
    purpose: List[str]
    scope: List[str]
    responsibility: List[str]
    definition: List[str]
    reference: List[str]
    attachment: List[str]
    procedure: List[ProcedureSection]


def transform_data(
    data: DocumentType | Dict[str, Any] | List[Any]
) -> DocumentType | Dict[str, Any] | List[Any]:
    """
    Helper function that recursively enters an unknown-amount nested dictionary or list and applies transformations to all string elements.
    This function can be flexibly edited to apply other transformations to the data structure as needed.
    In this example, we remove trailing newline characters from all string elements.
    Usage: transform_data(data)

    Args:
        data: The input data structure which can be a dictionary, list, string, or any other type.

    Returns:
        data: The transformed data structure with modifications applied to all string elements.
    """
    if isinstance(data, dict):
        return {
            transform_data(key): transform_data(value) for key, value in data.items()
        }
    elif isinstance(data, list):
        return [transform_data(element) for element in data]
    elif isinstance(data, str):
        # Edit the transformation here as needed
        return data.rstrip("\n")
    else:
        print("There are no string elements in the data structure to transform.")
        return data


def read_yaml_file(input_file: str) -> DocumentType | Dict:
    """
    Used to read a YAML file and return the content as a dictionary
    Usage: content = read_yaml_file('text.yml')

    Inputs:
    input_file: The file path of the YAML file to be read

    Returns:
    yaml_content: The content of the YAML file as a dictionary
    """
    with open(input_file, "r", encoding="utf-8") as file:
        yaml_content = yaml.safe_load(file)

    # Transform the data structure to remove trailing newline characters from all string elements
    yaml_content = transform_data(yaml_content)

    return yaml_content


def fill_common_items(yaml_content: DocumentType) -> Dict:
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
        - Procedure         (5.0) (filled in by extended functions)
        - Reference         (6.0)
        - Attachment        (7.0)

    Usage: fill_common_items(content)

    Inputs:
    yaml_content: The extracted information from the YAML file from read_yaml_file()

    Returns:
    context: The context dictionary with the common items filled in
    """
    context = {
        # Document Header and Footer Items
        "document_type": yaml_content["type"],
        "document_no": yaml_content["document_no"],
        "effective_date": yaml_content["effective_date"],
        "document_rev": yaml_content["document_rev"],
        "title": yaml_content["title"],
        "document_code": yaml_content["document_code"],
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


def get_mixed_types_str_keys(content: List[Any] | Dict[Any, Any]) -> List[str] | str:
    """
    Helper function to extract a list of strings and dictionary keys from a list or dictionary of mixed strings and/or dictionaries.
    Usage: get_mixed_types_str_keys(content['list']['sub-list']) or get_mixed_types_str_keys(content['dictionary'])

    Inputs:
    content: The list or dictionary containing mixed strings and/or dictionaries

    Returns:
    mixed_str_key_list: A list of strings and dictionary keys. If there is only one item, it is returned as a string instead of a list.
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
        # Else if the item is neither a string nor a dictionary, add an error message to the list
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
    content: The list containing mixed strings and dictionaries

    Returns:
    dict_values: A list of lists containing the values of the dictionaries in the input list. If there is only one list within the list,
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
    yaml_content: The extracted information from the YAML file from read_yaml_file()
    input_file: The file path of the Word document template
    output_file: The file path to save the filled Word document

    Returns:
    None

    For the example document, the data structure is as follows:
    {'type': 'EXAMPLE DOCUMENT',
    'document_no': 'EX01-100',
    'effective_date': '06-JUN-2024',
    'document_rev': '00',
    'title': 'Example Document',
    'document_code': 'EX01-100-01',
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
            'This is a very very very very long multi-line comment line 1.\nThis is a very very very very long multi-line comment line 2:': ['This is a sub-item of the multi-line comment.']
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
