# -*- coding: utf-8 -*-
import os
import regex as re
import xlwings as xw
import shutil
ROOT_DIR = os.path.dirname(__file__)
#%%
class CodeStructure:
    
    def __init__(self, level_0_dir):
        """
        Keyword arguments:\n
            level_0_dir : The directory path at level 0.\n
        """
        self.level_0_dir = level_0_dir
        self.patterns_list = [r'(?<=import\s)(.*?)(?=\sas\b)', r'(?<=from)(.*)(?=\simport\b)', r'(?<=import)(.*)(?=\n)']
        self.xlsx_template = rf"{ROOT_DIR}\xlsx_dependencies_template\xlsx_dependencies.xlsx"
 
    def __match_import_pattern(self, module_name):
        """
        Method used for getting the imported modules of .py files.\n
        """ 
        with open(module_name, encoding="utf8") as fileInp:
            lines = fileInp.readlines()
         
        matched_imports = []
        for line in lines:
            line = re.sub(r"#.*$", "", line) # handle comments
            for i, pattern in enumerate(self.patterns_list):
                try:
                    match = re.search(pattern, line).group(0).strip()
                    matched_imports.append(match)
                    break
                except Exception:
                    pass
                
        return matched_imports
    
    def __index_to_alphanumeric(self, index):
        """
        Convert an index to an alphanumeric representation.\n
        Code from Chat-GPT.\n
        """
        if index < 0:
            raise ValueError("Index must be a non-negative integer.")
    
        result = ""
        while index > 0:
            # Convert the remainder to ASCII character and prepend to the result
            result = chr((index - 1) % 26 + ord('A')) + result
            # Update the index to the quotient
            index = (index - 1) // 26
    
        return result
    
    def __get_module_imports(self, include_libs):
        """
        Method used for getting the imports per python module.\n
        """
        code_structure_dict = self.code_structure_dict()
        dependencies_dict = {}
        imported_modules_inproject, imported_modules_aux = [], []
        
        xlsx_row, module_idx = 3, 1
        for dir_level in code_structure_dict["Directories"].keys():
            for directory in code_structure_dict["Directories"][dir_level]:
                if directory in code_structure_dict["Files"].keys():
                    for file, file_full in code_structure_dict["Files"][directory]:
                        if file_full.endswith(".py"):
                            imported_modules, file, folder_name = self.__module_imports_replacements(file, file_full)
                            if folder_name not in dependencies_dict.keys():
                                dependencies_dict[module_idx] = [folder_name, file, imported_modules]
                                module_idx +=1
                                imported_modules_inproject.append(file)
                                imported_modules_aux += imported_modules
                            xlsx_row += 1
        
        if include_libs == True:
            dependencies_dict = self.__include_libraries_imports(dependencies_dict, imported_modules_inproject, imported_modules_aux)
       
        return dependencies_dict
    
    def __include_libraries_imports(self, dependencies_dict, imported_modules_inproject, imported_modules_aux):
        """
        Method used for including libraries into the dependencies.\n
        """
        imported_modules_lib = []
        for module in imported_modules_aux:
            if module not in imported_modules_inproject and module not in imported_modules_lib:
                imported_modules_lib.append(module)
                
        imported_modules_lib.sort()
        xlsx_row_lib = len(imported_modules_inproject)
        for module in imported_modules_lib:
            dependencies_dict[xlsx_row_lib+1] = ["<library>", module, []]
            xlsx_row_lib += 1

        return dependencies_dict
                      
    def __module_imports_replacements(self, file, file_full):
        """
        Auxiliary method used for formatting folder name and module names.
        """
        imported_modules = self.__match_import_pattern(file_full)
        folder_name = file_full.replace(self.level_0_dir, ".")
        folder_name = folder_name.replace(file, "")[:-1]
        if ".py" in file:
            file = file.replace(".py","")
            
        return imported_modules, file, folder_name
            
    def __mapping_module_to_moduleidx(self, dependencies_dict):
        """
        Auxiliary method that maps in a dictionary a module name to an index.\n
        """
        map_module_to_idx = {}
        for xlsx_row in range(1, len(dependencies_dict)+1):
            map_module_to_idx[dependencies_dict[xlsx_row][1]] = xlsx_row

        return map_module_to_idx
    
    def __mapping_moduleidx_to_imported_idxs(self, dependencies_dict):
        """
        Auxiliary method that maps a module index to imported module indices.\n
        """
        map_module_to_idx = self.__mapping_module_to_moduleidx(dependencies_dict)
        map_moduleIdx_importedModuleIdxs = {}
        for xlsx_row in range(1, len(dependencies_dict)):
            for i in range(len(dependencies_dict[xlsx_row][2])):
                if dependencies_dict[xlsx_row][2][i] in map_module_to_idx.keys():
                    _aux = map_module_to_idx[dependencies_dict[xlsx_row][2][i]]
                    if xlsx_row not in map_moduleIdx_importedModuleIdxs.keys():
                        map_moduleIdx_importedModuleIdxs[xlsx_row] = [_aux]
                    else:
                        map_moduleIdx_importedModuleIdxs[xlsx_row].append(_aux)

        return map_moduleIdx_importedModuleIdxs
       
    def __write_xlsx_columns_A_C(self, dependencies_dict, sht):
        """
        Method used for writing the A and C columns of the .xlsx dependencies file.\n 
        """
        xlsx_cols_A_C = {}
        for module_idx in dependencies_dict.keys():
            xlsx_cols_A_C[f"A{module_idx+2}"] = dependencies_dict[module_idx][0]
            xlsx_cols_A_C[f"C{module_idx+2}"] = dependencies_dict[module_idx][1]
        for cell, val in xlsx_cols_A_C.items():
            sht.range(cell).value = val
            
        return None
    
    def __write_xlsx_inner_cells(self, dependencies_dict, sht):
        """
        Method used for writing the cells that indicate the dependencies.\n 
        """
        map_moduleIdx_importedModuleIdxs = self.__mapping_moduleidx_to_imported_idxs(dependencies_dict)
        xlsx_row, xlsx_col = 2, 3
        for j in range(1, len(dependencies_dict)+1):
            for i in range(1, len(dependencies_dict)+1):
                alpha_numeric = self.__index_to_alphanumeric(xlsx_col+j)
                if i in map_moduleIdx_importedModuleIdxs.keys():
                    if j in map_moduleIdx_importedModuleIdxs[i]:
                        sht.range(f"{alpha_numeric}{xlsx_row +i}").value = "T"
                    else:
                        if i != j:
                            sht.range(f"{alpha_numeric}{xlsx_row +i}").value = "F"
                else:
                    if i != j:
                        sht.range(f"{alpha_numeric}{xlsx_row +i}").value = "F"
        return None
    
    def __excel_file_copy(self):
        """
        Creation of an .xlsx copy file using the template.\n
        """
        shutil.copy2(self.xlsx_template, "__COPY.xlsx")
        book = xw.Book("__COPY.xlsx")
        sheet= book.sheets[0]
        
        return book, sheet
    
    def __excel_file_rename(self, book, xlsx_dir, xlsx_name):
        """
        Method used for renaming the copied .xlsx file and deleting the copy.\n
        """
        book.save()
        book.close()
        xlsx_file_renamed = os.path.join(xlsx_dir, xlsx_name)
        try:
            os.rename("__COPY.xlsx", xlsx_file_renamed)
        except FileExistsError:
            os.remove(xlsx_file_renamed)
            os.rename("__COPY.xlsx", xlsx_file_renamed)
        return None
    
    def code_structure_dict(self, lvls_to_account="all"):
        """
        Function used for visualizing the code structure of a whole program.\n
        Keyword arguments:\n
            lvls_to_account : Default value "all".\n
        Returns a dictionary of the following format:\n
            {
                "Directories" : { 
                    <dir_level> : [<dir_name_1>...] 
                    },
                "Files"      : { 
                    <dir_name_1> : [[module, filepath]] 
                    }
            }
        """
        code_structure_dict = {"Directories":{}, "Files":{}}
        if lvls_to_account == "all":
            lvls_to_account = 10**6
        
        for root, dirs, files in os.walk(self.level_0_dir):
            level = root.replace(self.level_0_dir, '').count(os.sep)
            if level <= lvls_to_account:
                current_dir = os.path.basename(root)
                if len(files) == 0:
                    code_structure_dict["Files"][current_dir] = [["", f"{root}"]]
            for file in files:
                if level <= lvls_to_account:
                    if current_dir not in code_structure_dict["Files"].keys():
                        code_structure_dict["Files"][current_dir] = [[file, f"{root}\\{file}"]]
                    else:
                        code_structure_dict["Files"][current_dir].append([file, f"{root}\\{file}"])
    
            if level not in code_structure_dict["Directories"].keys():
                code_structure_dict["Directories"][level] = [current_dir]
            else:
                code_structure_dict["Directories"][level].append(current_dir)
    
        return code_structure_dict
    
    def code_structure_file(self, txt_dir_out, txt_name_out, lvls_to_account="all", print_in_terminal=True):
        """
        Keyword arguments:\n
            txt_dir_out  : Outpout directory.\n
            txt_name_out : Output filename.\n
            lvls_to_account : Number of levels after level 0 to take into account.\n
            print_in_terminal : Boolean. Allow to print in terminal the code structure.\n
        Returns None.
        """
        file_name_out = os.path.join(txt_dir_out, txt_name_out)
        code_structure_dict = self.code_structure_dict()
        dir_levels_dict = {value: key for key, values in code_structure_dict["Directories"].items() for value in values}
        PrintInTerminal = []
        with open(file_name_out, mode="w", encoding="utf8") as fOut:
            for diR in code_structure_dict["Files"].keys(): # Directories are fetched from code_structure_dict["Files"] dictionary and not from dir_levels_dict
                level = dir_levels_dict[diR]
                indent = f"{4*' '*level}"
                stR = f"{indent}./{diR}"
                PrintInTerminal.append(stR)
                fOut.write(f"{stR}\n")
                try:
                    for file, _ in code_structure_dict["Files"][diR]:
                        indent2 = 4*" "*(level+1)
                        stR = f"{indent2}|{file}"
                        PrintInTerminal.append(stR)
                        fOut.write(f"{stR}\n")
                except KeyError:
                    pass
        if print_in_terminal == True:
            for line in PrintInTerminal:
                print(line)
        return None
    
    def module_dependencies(self, xlsx_dir_out, xlsx_name_out, include_libs=False):
        """
        Method used for creating an .xlsx file indicating the dependencies of a\n
        python based program.\n
        Keyword arguments:\n
            xlsx_dir_out  : Outpout directory.\n
            xlsx_name_out : Output .xlsx filename.\n
            include_libs  : Boolean. If True then library dependencies (like numpy\n
                            or non-project specific) are included.\n 
        Returns a dictionary of the following format:\n
            {
                <module idx> : [
                    <module relative_dir>, <module name>, 
                    [
                        <imported_module_1>,
                        <imported_module_2>,
                    ]
                ]
            }.\n
            - <module relative_dir> : eg. "." if module is located at directory level 0.\n
            - <module name> : Module name without .py.\n
            - <imported module_1> ...: Imported modules in .py.
        """
        book, sheet = self.__excel_file_copy()
        dependencies_dict= self.__get_module_imports(include_libs)
        self.__write_xlsx_columns_A_C(dependencies_dict, sheet)
        self.__write_xlsx_inner_cells(dependencies_dict, sheet)
        self.__excel_file_rename(book, xlsx_dir_out, xlsx_name_out)
        
        return dependencies_dict