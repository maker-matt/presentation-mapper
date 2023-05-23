import os, re, yaml


def extract_keywords(text):
    pattern = r"\{\{(.+?)\}\}"
    matches = re.findall(pattern, text)
    return matches

def save_mapping(mapping, mapping_filename, replace = False):
    # Save the mapping to 'mapping.yml'
    mapping_filepath = os.path.join(os.getcwd(), mapping_filename)
    if os.path.exists(mapping_filepath):
        raise FileExistsError(f"File '{mapping_filename}' already exists")
    with open(mapping_filepath, 'w') as f:
        yaml.dump(mapping, f)

def read_mapping(mapping_filepath):
    def flatten_mapping(mapping):
        flattened_mapping = {}
        for slide, placeholders in mapping.items():
            flattened_mapping.update(placeholders)
        return flattened_mapping
    try:
        with open(mapping_filepath, 'r') as f:
            mapping = yaml.safe_load(f)
        return flatten_mapping(mapping)
    except FileNotFoundError:
        raise FileNotFoundError(f"File '{mapping_filepath}' not found")

def get_pptx_from_dir(dir_path):
    all_pptx = [file for file in os.listdir(dir_path) if file.endswith('.pptx')]
    assert len(all_pptx) == 1, f"Expected 1 .pptx file in current directory, found {len(all_pptx)}"
    only_pptx_file = all_pptx.pop()
    # return path to pptx file
    return os.path.join(dir_path, only_pptx_file)

def get_xlsx_from_dir(dir_path):
    all_xlsx = [file for file in os.listdir(dir_path) if file.endswith('.xlsx')]
    assert len(all_xlsx) == 1, f"Expected 1 .pptx file in current directory, found {len(all_xlsx)}"
    only_xlsx_file = all_xlsx.pop()
    return os.path.join(dir_path, only_xlsx_file)