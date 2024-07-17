# -*- coding: utf-8 -*-
import os
import sys
ROOT_DIR = os.path.dirname(__file__)
#%%
try:
    from code_structure import CodeStructure
    from formatext      import formatext
except ImportError:
    sys.path.append(os.path.realpath(os.path.join(ROOT_DIR, r"../codestruct")))
    from code_structure import CodeStructure
    sys.path.append(os.path.realpath(os.path.join(ROOT_DIR, r"../../formatext")))
    from formatext import formatext
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
formatext(msg, print_len_lim=printLen, align="c", fillin_str="#")
try:
    code_structure_dict = code_structure_obj.code_structure_dict()
    msg = "PASS"
    formatext(msg, print_len_lim=20, align="l", fillin_str="", inline_char=[":", "|"])
except Exception as ex:
    msg = f"Exception : {ex}"
    formatext(msg, print_len_lim=20, align="l", fillin_str="", inline_char=[":", "|"])
    formatext(msg, print_len_lim=20, align="l", fillin_str="", inline_char=[":", "|"])
#%%
msg = f"\n Checking {code_structure_obj.code_structure_file.__name__} \n"
formatext(msg, print_len_lim=printLen, align="c", fillin_str="#")
try:
    code_structure_obj.code_structure_file(txt_dir_out, txt_name_out, lvls_to_account="all", print_in_terminal=False)
    msg = "PASS"
    formatext(msg, print_len_lim=20, align="l", fillin_str="", inline_char=[":", "|"])
except Exception as ex:
    msg = f"Exception : {ex}"
    formatext(msg, print_len_lim=20, align="l", fillin_str="", inline_char=[":", "|"])
#%%  
msg = f"\n Checking {code_structure_obj.module_dependencies.__name__} \n"
formatext(msg, print_len_lim=printLen, align="c", fillin_str="#")
try:
    dependencies_dict = code_structure_obj.module_dependencies(xlsx_dir_out, xlsx_name_out, include_libs)
    msg = "PASS"
    formatext(msg, print_len_lim=20, align="l", fillin_str="", inline_char=[":", "|"])
except Exception as ex:
    msg = f"Exception : {ex}"
