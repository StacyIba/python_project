import json as json_lib
from openpyxl import Workbook
from openpyxl.styles import Font


class JsonHelper:
    def __init__(self,
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
                 description_field_name="screenTip",):
        self.input_json_file_name = input_json_file_name
        self.output_json_file_name = output_json_file_name
        self.output_excel_file_name = output_excel_file_name

        self.root_calc_element_name = root_calc_element_name
        self.root_filter_element_name = root_filter_element_name
        self.root_metadata_element_name = root_metadata_element_name
        self.root_folder_element_name = root_folder_element_name
        self.folder_element_name = folder_element_name
        self.ref_element_name = ref_element_name
        self.label_element_name = label_element_name
        self.identifier_field_name = identifier_field_name
        self.description_field_name = description_field_name

        self.folder_dict = {}
        self.folder_with_calc_filter = {}
        self.sheet_names = set()

        with open(self.input_json_file_name, encoding="utf-8") as f:
            self.json = json_lib.load(f)

    def write_to_json(self,
                      fields_list,
                      root_element_name,
                      input_json_file_name=None,
                      result_json_file_name=None,
                      ):
        if input_json_file_name:
            with open(input_json_file_name, encoding="utf-8") as f:
                json = json_lib.load(f)
        else:
            json = self.json

        if not result_json_file_name:
            result_json_file_name = self.output_json_file_name

        json_dict = {root_element_name: [JsonHelper.select_fields(
            fields=fields_list,
            dictionary=elem
        ) for elem in json[root_element_name]]}

        json.dump(json_dict, open(result_json_file_name, "w"), indent=2)

    def write_to_excel(self,
                       result_excel_file_name=None,):
        if not result_excel_file_name:
            result_excel_file_name = self.output_excel_file_name

        self.get_folder_to_calc_and_filter()
        self.sheet_names = sorted(self.sheet_names)
        sheet_dict = {sheet_name: {"calc_filter_list": []} for sheet_name in self.sheet_names}
        for folder, calc_filter in self.folder_with_calc_filter.items():
            main_folder, subfolder_label = folder.split("___")
            for sheet_name in self.sheet_names:
                if main_folder == sheet_name:
                    subfolder, label = subfolder_label.split("__")
                    sheet_dict[sheet_name]["calc_filter_list"].append({(subfolder, label): calc_filter})

        wb = Workbook()
        wb.remove(wb.active)
        for sheet_name in self.sheet_names:
            ws = wb.create_sheet(sheet_name)
            ws.title = sheet_name
            calc_filter_list = sheet_dict[sheet_name]["calc_filter_list"]
            # print(calc_filter_list)
            ws["A1"] = "Folder"
            ws["B1"] = "Items"
            ws["C1"] = "Description"
            row = 2
            subfolder_set = set()

            for calc_filter in calc_filter_list:
                for (subfolder, label), values in calc_filter.items():
                    if len(values) > 0:
                        if subfolder == "None":
                            if row != 2 and label not in subfolder_set:
                                row = row + 1
                            ws[f"A{row}"] = label
                            if label not in subfolder_set:
                                ws[f"A{row}"].font = Font(bold=True)
                        else:
                            if subfolder not in subfolder_set:
                                if row != 2:
                                    row = row + 1
                                ws[f"A{row}"] = subfolder
                                ws[f"A{row}"].font = Font(bold=True)
                                row = row + 1
                                subfolder_set.add(subfolder)
                            ws[f"A{row}"] = label
                        for elem_info in values:
                            for elem, desc in elem_info.items():
                                ws[f"B{row}"] = elem
                                ws[f"C{row}"] = desc
                                row = row + 1
                        row = row + 1

        wb.save(result_excel_file_name)

    def get_folder_to_calc_and_filter(self,
                                      fields_list=None,):
        if not fields_list:
            fields_list = [self.identifier_field_name,
                           self.label_element_name,
                           self.description_field_name]

        calculations_filter_list = []

        calculations_filter_list.extend([JsonHelper.select_fields(
            fields=fields_list,
            dictionary=calculation
        ) for calculation in self.json[self.root_calc_element_name]])

        calculations_filter_list.extend([JsonHelper.select_fields(
            fields=fields_list,
            dictionary=filter_dict
        ) for filter_dict in self.json[self.root_filter_element_name]])

        calculations_filter_dict = {
            calc_filter_info[self.identifier_field_name]:
                [calc_filter_info[self.label_element_name], calc_filter_info.get(self.description_field_name)]
            for calc_filter_info in calculations_filter_list
        }

        self.get_folder_hierarchy(
            folder_list=self.json[self.root_metadata_element_name][0][self.root_folder_element_name])

        for folder, refs in self.folder_dict.items():
            self.folder_with_calc_filter.update({
                folder: [{
                    calculations_filter_dict[ref][0]: calculations_filter_dict[ref][1]
                } for ref in refs]
            })

    @staticmethod
    def select_fields(fields, dictionary):
        output_dictionary = {}
        for field in fields:
            for key, value in dictionary.items():
                if dictionary.get(field, None):
                    output_dictionary[field] = dictionary[field]
        return output_dictionary

    def get_folder_hierarchy(self, folder_list):

        def get_refs(folder_dict,
                     folder_name,
                     subfolder_name=None):

            folders_list = folder_dict[self.folder_element_name][self.root_folder_element_name]
            folder_label = folder_dict[self.folder_element_name][self.label_element_name]
            refs = []
            for element in folders_list:
                if element.get(self.ref_element_name) is not None:
                    refs.append(element[self.ref_element_name])
                elif element.get(self.folder_element_name) is not None:
                    get_refs(element, folder_name, folder_label)
            self.sheet_names.add(folder_name)
            if folder_name == subfolder_name:
                pass
            else:
                self.folder_dict.update({f"{folder_name}___{subfolder_name}__{folder_label}": refs})

        for folder in folder_list:
            label = folder[self.folder_element_name][self.label_element_name]
            for subfolder in folder[self.folder_element_name][self.root_folder_element_name]:
                if label in ("Enterprise_Performance_Management", ):
                    pass
                else:
                    if subfolder.get(self.ref_element_name):
                        subfolder = {self.folder_element_name: folder[self.folder_element_name]}
                    get_refs(
                        folder_dict=subfolder,
                        folder_name=label)
