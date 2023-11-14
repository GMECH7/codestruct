# -*- coding: utf-8 -*-
import os
import sys
ROOT_DIR = os.path.dirname(__file__)
#%%
try:
    from code_structure import CodeStructure
    from formatext      import FormatText
except ImportError:
    sys.path.append(os.path.realpath(os.path.join(ROOT_DIR, r"../codestruct")))
    from code_structure import CodeStructure
    sys.path.append(os.path.realpath(os.path.join(ROOT_DIR, r"../../formatext")))
    from formatext import FormatText
#%%
PROJECT_DIR = rf"{ROOT_DIR}\project"
xlsx_dir_out  = rf"{ROOT_DIR}\results"
xlsx_name_out = "dependencies_test.xlsx"
txt_dir_out   = rf"{ROOT_DIR}\results"
txt_name_out  = "project_structure_test.dat"
include_libs=True
printLen    = 80
#%%
code_structure_obj = CodeStructure(PROJECT_DIR)
#%%
msg = f"\n Checking {code_structure_obj.code_structure_dict.__name__} \n"
FormatText(msg, printLenLimit=printLen, align="c", fillInstr="#")
try:
    code_structure_dict = code_structure_obj.code_structure_dict()
    msg = "PASS"
    FormatText(msg, printLenLimit=20, align="l", fillInstr="", inLinechr=[":", "|"])
except Exception as ex:
    msg = f"Exception : {ex}"
    FormatText(msg, printLenLimit=20, align="l", fillInstr="", inLinechr=[":", "|"])
    FormatText(msg, printLenLimit=20, align="l", fillInstr="", inLinechr=[":", "|"])
#%%
msg = f"\n Checking {code_structure_obj.code_structure_file.__name__} \n"
FormatText(msg, printLenLimit=printLen, align="c", fillInstr="#")
try:
    code_structure_obj.code_structure_file(txt_dir_out, txt_name_out, lvls_to_account="all", print_in_terminal=False)
    msg = "PASS"
    FormatText(msg, printLenLimit=20, align="l", fillInstr="", inLinechr=[":", "|"])
except Exception as ex:
    msg = f"Exception : {ex}"
    FormatText(msg, printLenLimit=20, align="l", fillInstr="", inLinechr=[":", "|"])
#%%  
msg = f"\n Checking {code_structure_obj.module_dependencies.__name__} \n"
FormatText(msg, printLenLimit=printLen, align="c", fillInstr="#")
try:
    dependencies_dict = code_structure_obj.module_dependencies(xlsx_dir_out, xlsx_name_out, include_libs)
    msg = "PASS"
    FormatText(msg, printLenLimit=20, align="l", fillInstr="", inLinechr=[":", "|"])
except Exception as ex:
    msg = f"Exception : {ex}"
