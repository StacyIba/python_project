from json_helper import JsonHelper

if __name__ == '__main__':
    """
    Default parameters in JsonHelper(), you can override them
        input_json_file_name="DM.json",
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
        description_field_name="screenTip"
    """
    json_helper = JsonHelper(
        output_excel_file_name="output.xlsx"
    )
    json_helper.write_to_excel()
