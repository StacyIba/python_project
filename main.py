from json_helper import JsonHelper

if __name__ == '__main__':
    """
    Default parameters in JsonHelper(), you can override them
        input_json_file_name="DM",
        output_json_file_name="result.json",
        output_excel_file_name="result.xlsx",
        
        root_calc_element_name="calculation",
        root_filter_element_name="filter",
        root_metadata_element_name="metadataTreeView",
        
        root_folder_element_name="folderItem",
        folder_element_name="folder",
        ref_element_name="ref",
        label_element_name="label",
        identifier_field_name="identifier",
        hidden_field_name="hidden",
        
        domains_to_exclude=("Enterprise_Performance_Management",),
        
        columns_to_output_xlsx = {
                "A": ("Structure", "identifier", None),
                "B": ("Items / Filters", "label", None),
                "C": ("Description", "screenTip", None),
                "D": ("Hidden away", "hidden", "FALSE")
        }
        
        delete_hidden=False
    """
    json_helper_developers = JsonHelper(
        output_excel_file_name="output_for_developers.xlsx",
        delete_hidden=False
    )
    json_helper_developers.write_to_excel()

    json_helper_users = JsonHelper(
        output_excel_file_name="output_for_users.xlsx",
        delete_hidden=True,
    )
    json_helper_users.write_to_excel()

    json_helper_with_expression = JsonHelper(
        columns_to_output_xlsx={
            "A": ("Structure", "identifier", None),
            "B": ("Items / Filters", "label", None),
            "C": ("Description", "screenTip", None),
            "D": ("Expression", "expression", None),
        },
        output_excel_file_name="output_with_expression.xlsx",
        hidden_field_name=None,
    )
    json_helper_with_expression.write_to_excel()
