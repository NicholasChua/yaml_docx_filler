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


def get_mixed_types_str_keys(content: List[Any]) -> List[str]:
    """
    Helper function to extract a list of strings and dictionary keys from a list of mixed strings and dictionaries.
    Usage: get_mixed_types_str_keys(content['list']['sub-list'])

    Inputs:
    content (list): The list containing mixed strings and dictionaries

    Returns:
    mixed_str_key_list (list): A list of strings and dictionary keys
    """
    mixed_str_key_list = []
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
    dict_values: A list of lists containing the values of the dictionaries in the input list
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

    if dict_count == 0:
        print("No dictionaries found in the list")
        return None

    return dict_values


def get_multi_level_content(
    content: List[Any],
) -> List[Union[List[str], List[List[Dict[str, Union[str, Dict[str, str]]]]]]]:
    """
    Helper function that returns a list of lists, the first contains top-level strings and dictionary keys.
    The second contains a list of lists of dictionary values.
    Used to form a data structure allowing access to multi-level content in a predictable way.
    Usage: get_multi_level_content(content['list']['sub-list'])

    E.g. To get content from the top-level > sub-list > sub-sub-list2,
    use the following structure of indices:
    - top-level: content['list']
        - sub-list-item-1: content['list']['sub-list'][0]
            - sub-sub-list1: content['list']['sub-list'][0][0]
                - sub-sub-sub-list1: content['list']['sub-list'][0][0][0]
        - sub-list-item-2: content['list']['sub-list'][1]

    Inputs:
    content (list): The list containing mixed strings and dictionaries

    Returns:
    list_of_lists (list): A list of lists containing the top-level strings and dictionary keys and the dictionary values
    """
    list_of_lists = []
    # Get all the strings and dictionary keys and append them to the list
    top_level_str = get_mixed_types_str_keys(content)
    list_of_lists.append(top_level_str)
    # Get all the dictionary values and append them to the list
    dict_values = get_mixed_types_dict_values(content)
    list_of_lists.append(dict_values)

    return list_of_lists


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
    """

    # Open the Word document
    doc = DocxTemplate(input_file)

    # Prepare the context for the template
    # Call the base function to fill in common items
    context = fill_common_items(yaml_content)

    # Fill in the Procedure section
    context.update(
        {
            # Splitting up the Procedure section into multiple levels due to its complexity
            "example_policy_1": str(
                get_mixed_types_str_keys(yaml_content["procedure"]["example_policy"])[0]
            ).strip("['']"),
            "sub_example_policy_1": get_mixed_types_dict_values(
                yaml_content["procedure"]["example_policy"]
            )[0],
            "example_policy_2": str(
                get_mixed_types_str_keys(yaml_content["procedure"]["example_policy"])[1]
            ).strip("['']"),
            "sub_example_policy_2": get_mixed_types_dict_values(
                yaml_content["procedure"]["example_policy"]
            )[1],
            "example_policy_3": str(
                get_mixed_types_str_keys(yaml_content["procedure"]["example_policy"])[2]
            ).strip("['']"),
            "another_policy_1": str(
                get_mixed_types_str_keys(yaml_content["procedure"]["another_policy"])
            ).strip("['']"),
            "sub_another_policy_1": get_mixed_types_dict_values(
                yaml_content["procedure"]["another_policy"]
            )[0],
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
