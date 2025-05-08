import pandas as pd
import xml.etree.ElementTree as ET

def extract_excel_columns(file):
    """
    Extract column names from all sheets in an Excel file
    
    Args:
        file (str): Path to Excel file
    
    Returns:
        dict: Dictionary mapping sheet names to lists of column names
    """
    try:
        xls = pd.ExcelFile(file)
        all_columns = {}
        for sheet in xls.sheet_names:
            # Read just the first row to get column names
            df = pd.read_excel(file, sheet_name=sheet, nrows=1)
            all_columns[sheet] = list(df.columns)
        return all_columns
    except Exception as e:
        print(f"Error extracting Excel columns: {str(e)}")
        return {}

def extract_xml_tags(xml_file):
    """
    Extracts paths for all attributes and structural leaf elements from an XML file.
    
    Args:
        xml_file (str): Path to XML file
    
    Returns:
        list: List of unique XML paths
    """
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        def get_xpath(element, current_path=""):
            """Get XPath for the element and its attributes"""
            paths = []
            
            # Create the path for this element
            if current_path:
                path = f"{current_path}/{element.tag}"
            else:
                path = element.tag
            
            # Add the element path if it's a leaf node (no children but has text)
            if len(element) == 0 and element.text and element.text.strip():
                paths.append(path)
            
            # Add paths for all attributes
            for attr_name in element.attrib:
                paths.append(f"{path}/@{attr_name}")
            
            # Recursively process child elements
            for child in element:
                paths.extend(get_xpath(child, path))
                
            return paths

        all_paths = get_xpath(root)
        return sorted(list(set(all_paths)))
    except Exception as e:
        print(f"Error extracting XML tags: {str(e)}")
        return []