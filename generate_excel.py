#!/usr/bin/env python3

import textract
import os, re

# REMOVE SPACES
def remove_space(name):
    return name.replace(" ", "")

# REMOVE SPACES
def fix_name_doc():
    path = "/PATH"
    files_name = []
    for filename in os.listdir(path):
        if " " in filename:
            fixed_name = remove_space(filename)

            src = os.path.join(path, filename)
            dst = os.path.join(path, fixed_name)

            # Only rename if names are different
            if filename != fixed_name:
                os.rename(src, dst)
                print("FIXED: " + filename + " → " + fixed_name)
            
        else:
            print("NAME LOOKS FINE.")

        files_name.append(filename)

    return files_name


def get_info(keyword):
    files_name = fix_name_doc()
    path = "/PATH"
    
    for file in files_name:
        full_path = path + file
        print(f"Processing file: {file}")

        try:
            text_bytes = textract.process(full_path)
            text = text_bytes.decode('utf-8', errors='ignore')
        except Exception as e:
            print(f"Error reading {file}: {e}")
            continue

        normalized_text = re.sub(r'\s+', ' ', text).strip()


        pattern = re.compile(rf"{re.escape(keyword)}\s*[:\-]?\s*(.+?)(?=(\.|$))", re.IGNORECASE)

        match = pattern.search(normalized_text)
        if match:
            project_name = match.group(1).strip()
            print(f"Found in {file}: {project_name}")
            return project_name
        
        print(f"No match found in {file}")

    return ""








#def generate_info_csv():
    #path = "/home/totto/scripting/python/TCU/docs/"
    #for filename in os.listdir(path):
        #nombre project=$(cat filename| grep -i nombre del proyecto)



#text = textract.process("docs/02-AntepTCU-ULAT-FundaciónRealMadridEmprendedores.docx")
get_info("NOMBRE del ANTEPROYECTO")
