import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
import tempfile

def read_xml_from_gsz(gsz_file_path):
    try:
        with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
            root_xml = [f for f in zip_ref.namelist() if f.endswith('.xml') and '/' not in f][0]
            xml_content = zip_ref.read(root_xml)
            return ET.ElementTree(ET.fromstring(xml_content))
    except Exception as e:
        print(f"Error reading XML from gsz: {e}")
        return None

def write_xml_to_gsz(gsz_file_path, tree):
    try:
        temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        xml_files = [f for f in os.listdir(temp_dir) if f.endswith('.xml')]
        if not xml_files:
            shutil.rmtree(temp_dir)
            return False
            
        root_xml_path = os.path.join(temp_dir, xml_files[0])
        tree.write(root_xml_path, encoding='UTF-8', xml_declaration=True)
        
        with zipfile.ZipFile(gsz_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)
                    
        shutil.rmtree(temp_dir)
        return True
    except Exception as e:
        if 'temp_dir' in locals() and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        return False

def extract_analysis_names(gsz_file_path):
    analysis_names = []
    try:
        with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
            root_xml_files = [f for f in zip_ref.namelist() if f.endswith('.xml') and '/' not in f]
            if not root_xml_files: return analysis_names
            
            with zip_ref.open(root_xml_files[0]) as xml_file:
                root = ET.fromstring(xml_file.read())
                analyses = root.find('Analyses')
                if analyses is not None:
                    for analysis in analyses.findall('Analysis'):
                        name_element = analysis.find('Name')
                        if name_element is not None and name_element.text:
                            analysis_names.append(name_element.text)
    except Exception as e:
        print(f"Error: {e}")
    return analysis_names

def extract_DataPoint_coordinates(gsz_file_path, analysis_name, DataPoint_index):
    try:
        with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            analysis_xml_files = [f for f in file_list if f.endswith('.xml') and f.startswith(f"{analysis_name}/")]
            if not analysis_xml_files:
                analysis_xml_files = [f for f in file_list if f.endswith('.xml') and '/' not in f]

            with zip_ref.open(analysis_xml_files[0]) as xml_file:
                root = ET.fromstring(xml_file.read())
                data_points = root.findall('.//DataPoint')
                
                if 0 < DataPoint_index <= len(data_points):
                    data_point = data_points[DataPoint_index - 1]
                    x_value, y_value = data_point.get('X'), data_point.get('Y')
                    if x_value is not None and y_value is not None:
                        return (float(x_value), float(y_value))
    except Exception as e:
        print(f"Error: {e}")
    return None

def extract_reinforcement_line_coordinates(gsz_file_path, analysis_name):
    reinforcement_coords = {}
    try:
        with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            analysis_xml_files = [f for f in file_list if f.endswith('.xml') and f.startswith(f"{analysis_name}/")]
            if not analysis_xml_files:
                analysis_xml_files = [f for f in file_list if f.endswith('.xml') and '/' not in f]

            with zip_ref.open(analysis_xml_files[0]) as xml_file:
                root = ET.fromstring(xml_file.read())
                reinforcement_lines = root.findall('.//ReinforcementLine')
                
                datapoint_dict = {}
                for dp in root.findall('.//DataPoint'):
                    number, x_value, y_value = dp.get('Number'), dp.get('X'), dp.get('Y')
                    if number and x_value and y_value:
                        datapoint_dict[number] = (float(x_value), float(y_value))
                
                for reinf_line in reinforcement_lines:
                    line_id, point1_id = reinf_line.get('ID'), reinf_line.get('Point1Id')
                    if line_id and point1_id and point1_id in datapoint_dict:
                        reinforcement_coords[int(line_id)] = datapoint_dict[point1_id]
    except Exception as e:
        print(f"Error: {e}")
    return reinforcement_coords

def extract_fully_specified_slip_coordinates(gsz_file_path, analysis_name):
    slip_surfaces = {}
    try:
        with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
            root_xml_files = [f for f in zip_ref.namelist() if f.endswith('.xml') and '/' not in f]
            if not root_xml_files: return slip_surfaces

            with zip_ref.open(root_xml_files[0]) as xml_file:
                root = ET.fromstring(xml_file.read())

                analysis_id = None
                analyses = root.find('Analyses')
                if analyses is not None:
                    for analysis in analyses.findall('Analysis'):
                        if analysis.find('Name') is not None and analysis.find('Name').text == analysis_name:
                            analysis_id = analysis.find('ID').text
                            break
                
                target_stability_item = None
                stability_items = root.find('.//StabilityItems')
                if stability_items is not None:
                    for item in stability_items.findall('StabilityItem'):
                        item_analysis_id = item.find('AnalysisID')
                        if item_analysis_id is not None and (analysis_id is None or item_analysis_id.text == analysis_id):
                            target_stability_item = item
                            break
                
                if target_stability_item is None: return slip_surfaces
                
                entry = target_stability_item.find('Entry')
                if entry is None or entry.find('SlipSurface') is None or entry.find('SlipSurface').find('FullySpecifiedSlips') is None:
                    return slip_surfaces
                
                datapoint_dict = {}
                if entry.find('DataPoints') is not None:
                    for dp in entry.find('DataPoints').findall('DataPoint'):
                        if dp.get('Number') and dp.get('X') and dp.get('Y'):
                            datapoint_dict[dp.get('Number')] = (float(dp.get('X')), float(dp.get('Y')))
                
                for slip_element in entry.find('SlipSurface').find('FullySpecifiedSlips').findall('FullySpecifiedSlip'):
                    if slip_element.find('ID') is None or slip_element.find('DataPoints') is None: continue
                    
                    slip_id = int(slip_element.find('ID').text)
                    coordinates = []
                    for dp in slip_element.find('DataPoints').findall('DataPoint'):
                        if dp.text in datapoint_dict: coordinates.append(datapoint_dict[dp.text])
                    
                    if coordinates: slip_surfaces[slip_id] = coordinates
    except Exception as e:
        print(f"Error: {e}")
    return slip_surfaces

def extract_reinforcement_names(gsz_file_path, analysis_name):
    name_map = {}
    try:
        with zipfile.ZipFile(gsz_file_path, 'r') as z:
            file_list = z.namelist()
            xml_files = [f for f in file_list if f.endswith('.xml') and f.startswith(f"{analysis_name}/")]
            if not xml_files: xml_files = [f for f in file_list if f.endswith('.xml') and '/' not in f]

            for xml_file in xml_files:
                raw = z.read(xml_file).decode('utf-8', errors='ignore')
                if '<Reinforcements' not in raw: continue
                
                root = ET.fromstring(raw)
                reinf_names = {}
                for idx, reinf in enumerate(root.findall('.//Reinforcements/Reinforcement'), start=1):
                    name_el = reinf.find('Name')
                    reinf_names[idx] = name_el.text.strip() if (name_el is not None and name_el.text) else str(idx)

                for reinf_line in root.findall('.//ReinforcementLine'):
                    if reinf_line.get('ID') and reinf_line.get('Reinforcement'):
                        try:
                            line_id, reinf_idx = int(reinf_line.get('ID')), int(reinf_line.get('Reinforcement'))
                            name_map[line_id] = reinf_names.get(reinf_idx, f'Nail {line_id}')
                        except ValueError: pass
                if name_map: break
    except Exception as e:
        print(f"Error: {e}")
    return name_map

def get_analysis_id_by_name(root: ET.Element, analysis_name: str) -> str:
    analyses = root.find('Analyses')
    if analyses is not None:
        for analysis in analyses.findall('Analysis'):
            if analysis.find('Name') is not None and analysis.find('Name').text == analysis_name:
                return analysis.find('ID').text if analysis.find('ID') is not None else None
    return None

def find_stability_item_for_analysis(root: ET.Element, analysis_id: str) -> ET.Element:
    stability_items = root.find('StabilityItems')
    if stability_items is not None:
        for item in stability_items.findall('StabilityItem'):
            if item.find('AnalysisID') is not None and item.find('AnalysisID').text == analysis_id:
                return item
    return None

def set_entry_exit_in_xml(root: ET.Element, analysis_name: str, entry_left_pt: tuple, entry_right_pt: tuple, exit_left_pt: tuple, exit_right_pt: tuple, left_inc: int = 45, right_inc: int = None) -> bool:
    analysis_id = get_analysis_id_by_name(root, analysis_name)
    if not analysis_id: return False
    
    stability_item = find_stability_item_for_analysis(root, analysis_id)
    if not stability_item: return False
    
    slip_surface = stability_item.find('.//SlipSurface')
    if slip_surface is None:
        entry_elem = stability_item.find('Entry')
        if entry_elem is not None: slip_surface = ET.SubElement(entry_elem, 'SlipSurface')
        else: return False
    
    option = stability_item.find('.//Option')
    if option is None:
        option = ET.Element('Option')
        stability_item.insert(0, option)
    option.text = 'EntryAndExit'
    
    radius = slip_surface.find('Radius')
    if radius is not None and 'UsePoints' in radius.attrib: del radius.attrib['UsePoints']
    
    entry_exit = slip_surface.find('EntryExit')
    if entry_exit is None: entry_exit = ET.SubElement(slip_surface, 'EntryExit')
    
    def update_elem(parent, tag, x, y):
        elem = parent.find(tag)
        if elem is None: elem = ET.SubElement(parent, tag)
        elem.set('X', str(x))
        elem.set('Y', str(y))

    update_elem(entry_exit, 'LeftSideLeftPt', *entry_left_pt)
    update_elem(entry_exit, 'LeftSideRightPt', *entry_right_pt)
    update_elem(entry_exit, 'RightSideLeftPt', *exit_left_pt)
    update_elem(entry_exit, 'RightSideRightPt', *exit_right_pt)
    
    left_inc_elem = entry_exit.find('LeftInc')
    if left_inc_elem is None: left_inc_elem = ET.SubElement(entry_exit, 'LeftInc')
    left_inc_elem.text = str(left_inc)
    
    if right_inc is not None:
        right_inc_elem = entry_exit.find('RightInc')
        if right_inc_elem is None: right_inc_elem = ET.SubElement(entry_exit, 'RightInc')
        right_inc_elem.text = str(right_inc)
    
    return True

def modify_gsz_entry_exit(gsz_file_path: str, entry_exit_settings: dict, output_path: str = None, max_retries: int = 3) -> str:
    import time
    if output_path is None: output_path = gsz_file_path
    temp_dir = tempfile.mkdtemp(prefix='gsz_edit_')
    
    try:
        for attempt in range(max_retries):
            try:
                with zipfile.ZipFile(gsz_file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                break
            except PermissionError as e:
                if attempt < max_retries - 1: time.sleep(1)
                else: raise
        
        xml_files = [f for f in os.listdir(temp_dir) if f.endswith('.xml')]
        root_xml_path = os.path.join(temp_dir, xml_files[0])
        tree = ET.parse(root_xml_path)
        root = tree.getroot()
        
        for analysis_name, settings in entry_exit_settings.items():
            set_entry_exit_in_xml(
                root=root, analysis_name=analysis_name,
                entry_left_pt=settings['entry_left'], entry_right_pt=settings['entry_right'],
                exit_left_pt=settings['exit_left'], exit_right_pt=settings['exit_right'],
                left_inc=settings.get('left_inc', 45), right_inc=settings.get('right_inc', None)
            )
        
        tree.write(root_xml_path, encoding='UTF-8', xml_declaration=True)
        
        for item in os.listdir(temp_dir):
            item_path = os.path.join(temp_dir, item)
            if os.path.isdir(item_path):
                for folder_xml in [f for f in os.listdir(item_path) if f.endswith('.xml')]:
                    shutil.copy2(root_xml_path, os.path.join(item_path, folder_xml))
        
        temp_gsz = gsz_file_path + '.tmp' if output_path == gsz_file_path else output_path
        with zipfile.ZipFile(temp_gsz, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    zipf.write(file_path, os.path.relpath(file_path, temp_dir))
        
        if output_path == gsz_file_path:
            for attempt in range(max_retries):
                try:
                    os.replace(temp_gsz, gsz_file_path)
                    break
                except PermissionError as e:
                    if attempt < max_retries - 1: time.sleep(1)
                    else: raise
        return output_path
    finally:
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)

def set_all_reinforcement_tensile_capacity(gsz_file_path, analysis_name, capacity_value):
    try:
        tree = read_xml_from_gsz(gsz_file_path)
        if tree is None: return False
        
        root = tree.getroot()
        found_count = 0
        for reinf in root.findall('.//Reinforcements/Reinforcement'):
            capacity_elem = reinf.find('TensileCapacity')
            if capacity_elem is not None:
                capacity_elem.text = str(capacity_value)
            else:
                new_cap = ET.Element('TensileCapacity')
                new_cap.text = str(capacity_value)
                reinf.append(new_cap)
            found_count += 1
            
        return write_xml_to_gsz(gsz_file_path, tree)
    except Exception as e:
        print(f"Failed to set tensile capacity: {e}")
        return False