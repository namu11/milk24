import re
import pandas as pd

def apply_fun_key(dic,keyy,fun):
    for key in dic:
        if keyy in key:
            key[keyy] = fun(key[keyy])
    return dic

def ext_numbers(numcha):
    remm = re.findall(r'\d+\.\d+|\d+', str(numcha))
    #아래는 리스트형식([])으로 반환됨
    remm2 = [float(num) if '.' in num else int(num) for num in remm]
    return remm2[0]

def read_stkup_sr(TB, lines):
    if TB == 'TOP':
        srtb = 'SRT'
    else:
        srtb = 'SRB'
            
    sr_stk = {
        'ID': f'CAD_{TB}_SOLDERMASK', 
        'TYPE': srtb, 
        'MATERIAL': '', 
        'THICKNESS': '', 
        'TOLERANCE': '', 
        'VIA FILL MATERIAL': ''}

    isr = 0
    while isr < len(lines):
        line = lines[isr].strip()
        if line.startswith(f'CAD_{TB}_SOLDERMASK') and any(key in line for key in ['MATERIAL', 'THICKNESS', 'TOLERANCE', 'FILL_MATERIAL']):
            if lines[isr+1].strip() == '9':
                srvalue = lines[isr+4].strip()
                match line:
                    case _ if 'SOLDERMASK_MATERIAL' in line:
                        sr_stk['MATERIAL'] = srvalue
                    case _ if 'THICKNESS' in line:
                        sr_stk['THICKNESS'] = srvalue
                    case _ if 'TOLERANCE' in line:
                        sr_stk['TOLERANCE'] = srvalue
                    case _:
                        sr_stk['VIA FILL MATERIAL'] = srvalue
            isr += 5
        else:
            isr += 1
    return sr_stk


def extract_drawing(filename):
    with open(filename, 'r', encoding='utf-8') as f:
            lines = f.readlines()

    #list
    stackup = []
    srt_stk = read_stkup_sr("TOP", lines)
    stackup.append(srt_stk)
                                      

    i = 0
    while i < len(lines) -4:
        line = lines[i].strip()
        if line.startswith('CAD_DIELECTRIC') and any(key in line for key in ['_TYPE', '_MATERIAL','_THICKNESS','TOLERANCE']):
            tag = line
            if lines[i+1].strip() == '9':
                value = lines[i+4].strip()
                match = re.match(r'(CAD_DIELECTRIC\d+_\d+)_(TYPE|MATERIAL|THICKNESS|TOLERANCE)',tag)
                if match:
                    dielectric_id, attribute = match.groups()
                    entry = next((item for item in stackup if item['ID'] == dielectric_id), None)
                    if not entry:
                        entry = {'ID': dielectric_id, 'TYPE': '', 'MATERIAL': '', 'THICKNESS':'','TOLERANCE': '', 'VIA FILL MATERIAL':''}
                        stackup.append(entry)
                        #print(entry)
                        #print(stackup)
                    entry[attribute] = value
                    #print(entry)
                    #print(stackup)
            i += 5
        else:
            i += 1
    
    stackup2 = []
    j = 0
    while j < len(lines) -4:
        line = lines[j].strip()
        if line.startswith('CAD_VIA_FILL') and any(key in line for key in ['MATERIAL']):
            tag2 = line
            if lines[j+1].strip() =='9':
                value2 = lines[j+4].strip()
                match2 = re.match(r'(CAD_VIA_FILL\d+_\d+)_(MATERIAL)',tag2)
                if match2:
                    layer_id, vvalue = match2.groups()
                    entry2 = next((item for item in stackup if item['ID'] == layer_id),None)
                    if not entry2:
                        entry2 = {'ID': layer_id, 'MATERIAL':''}
                        stackup2.append(entry2)
                    entry2[vvalue] = value2
            j += 5
        else:
            j += 1

    for item in stackup2:
        item['VIA FILL MATERIAL'] = item.pop('MATERIAL')

    ila = 0



    #print(stackup2) 

    def extract_number(aa):
        return re.findall(r'\d+',aa)   

    for item_a in stackup:
        num_a = extract_number(item_a['ID'])
        for item_b in stackup2:
            num_b = extract_number(item_b['ID'])
            if num_a == num_b:
                item_a['VIA FILL MATERIAL'] = item_b['VIA FILL MATERIAL']
                break
    
    
    srb_stk = read_stkup_sr("BOTTOM", lines)
    stackup.append(srb_stk)

    #add cu layer

    il = 0
    while il < len(lines) -4:
        linel = lines[il].strip()
        if linel.startswith('CAD_LAYER') and any (key in linel for key in ['_MATERIAL','_THICKNESS','TOLERANCE']):
            tagl = linel
            if lines[il+1].strip() == '9':
                valuel = lines[il+4].strip()
                matchl = re.match(r'(CAD_LAYER_\d+)_(MATERIAL|THICKNESS|TOLERANCE)',tagl)
                if matchl:
                    layer_id, attl = matchl.groups()
                    entryl = next((item for item in stackup if item['ID'] == layer_id), None)
                    if not entryl:
                        entryl = {'ID': layer_id, 'TYPE': '-', 'MATERIAL': '', 'THICKNESS': '', 'TOLERANCE': '', 'VIA FILL MATERIAL':'-'}
                        stackup.append(entryl)
                    entryl[attl] = valuel
            il += 5
        else:
            il += 1

    ifi = 0
    while ifi < len(lines) -4:
        linefi = lines[ifi].strip()
        if linefi.startswith('CAD_SURFACE_FINISH') and any (key in linefi for key in ['BASE','BUMP','BOND_FINGER','SMD','BGA','FIDUCIAL','STRIP']):
            tagfi = linefi
            if lines[ifi+1].strip() =='9':
                valuefi = lines[ifi+4].strip()
                matchfi = re.match('(CAD_SURFACE_FINISH)_(BASE|BUMP|BOND_FINGER|SMD_TOP|SMD_BOTTOM|BGA_TOP|BGA_BOTTOM|FIDUCIAL_TOP|FIDUCIAL_BOTTOM|STRIP)',tagfi)
                if matchfi:
                    finish_area_id, attfi = matchfi.groups()
                    entryfi = {'ID':'finish_'+attfi, 'TYPE': valuefi}
                    stackup.append(entryfi)
            ifi += 5
        else:
            ifi += 1            

    #print(stackup)

    apply_fun_key(stackup,'TOLERANCE',ext_numbers)

    #print(stackup)

    return stackup



# 사용 예시
if __name__ == '__main__':
    data = extract_drawing('in_a.txt')  # 파일명 변경 가능
    df = pd.DataFrame(data)
    df.to_excel('ex_a.xlsx', sheet_name='material', index=False)
    print("Excel 파일이 성공적으로 생성되었습니다: ex_a.xlsx")
