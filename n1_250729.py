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

def extract_drawing_ud(filename):
    with open(filename, 'r', encoding='utf-8') as f:
            lines_ud = f.readlines()
    return lines_ud

def extract_drawing(filename):
    with open(filename, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    file_path1 = "ABF_list.xlsx"
    abflist = pd.read_excel(file_path1, engine = 'openpyxl')
    #abfli = abflist.values.tolist()
    abfdi = abflist.to_dict(orient="records")
    #print(abfli)
    print(abfdi)

    file_path2 = "cup.xlsx"
    cupo = pd.read_excel(file_path2, engine = 'openpyxl')
    cupodi = cupo.to_dict(orient="records")
    print(cupodi)

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

    #extract_layer_count
    layer_count = 0
    ii = 0
    while ii < len(lines) -4:
        linei = lines[ii].strip()
        if linei == 'CAD_LAYER_COUNT':
            if lines[ii+1].strip() == '9':
                layer_count = lines[ii+4].strip()
            ii += 5
        else:
            ii += 1
    ii=0        
    
    print(layer_count)

    #extract_unit_size
    unitxy = 0
    unitx = 0
    unity = 0
    lines_ud = extract_drawing_ud('in_b.txt')

    while ii < len(lines_ud) -4:
        lineu = lines_ud[ii].strip()
        if lineu == 'CAD_BODY_SIZE':
            if lines_ud[ii+1].strip() == '9':
                unitxy = lines_ud[ii+4].strip()
            ii += 5
        else:
            ii += 1
    ii=0

    unitxy2 = re.findall(r'\d+\.\d+|\d+', str(unitxy))
    unitxy3 = [float(num) if '.' in num else int(num) for num in unitxy2]
    unitx = unitxy3[0]
    unity = unitxy3[1]

    print(unitx)
    print(unity)

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

    for dic in stackup:
        for key, value in dic.items():
            if isinstance(value, str):
                if 'gx92' in value.lower():
                    dic[key] = 'GX92'
                elif 'gz41' in value.lower():
                    dic[key] = 'GZ41'
                elif 'gl102' in value.lower():
                    dic[key] = 'GL102'
                elif 'gl107' in value.lower():
                    dic[key] = 'GL107'
                elif 'gxt31' in value.lower():
                    dic[key] = 'GXT31'
                elif 'gxt37' in value.lower():
                    dic[key] = 'GXT37'
                elif 'sr7300' in value.lower():
                    if 'c' in value.lower():
                        dic[key] = 'SR7300GRC'
                    else:
                        dic[key] = 'SR7300GRB'
                elif 'aus703' in value.lower():
                    dic[key] = 'AUS703'
                elif '700' in value.lower():
                    if 'h' in value.lower():
                        dic[key] = 'E700GLH'
                    else:
                        dic[key] = 'E700G'
                elif '705' in value.lower():
                    if 'h' in value.lower():
                        dic[key] = 'E705GLH'
                    else:
                        dic[key] = 'E705G'
                elif '795' in value.lower():
                    if 'h' in value.lower():
                        dic[key] = 'E795GLH'
                    else:
                        dic[key] = 'E795G'

    #자재선정
    multiple = float(100/(unitx*unity))
    lc_set = set()
    lc = int(layer_count)
    ilc = 0
    while ilc < lc:
        lc_set.add(ilc+1)
        ilc += 1
    
    print(lc_set)

    mamap = [{'cu':item['cu']} for item in cupodi if item['layer'] in lc_set]
    print(mamap)

    key_to_multiple = ['cu']
    for item in mamap:
        for key in key_to_multiple:
            if key in item:
                item[key] *= multiple
    print(cupodi)
    print(multiple)
    print(mamap)

    #mamap은 리스트고 저 안에 하나하나가 딕셔너리래

    
    #내부 자재선정 list
    in_table = []
    
    itc = 0
    while itc < lc:
        val2 = mamap[itc]['cu']
        val3 = next((item['THICKNESS'] for item in stackup if item['ID'] == f'CAD_LAYER_{itc+1}'),None)
        val3_f = float(val3) * 1000
        val4 = val3_f*(1-(val2/100))
        if itc == 0:
            val5 = 0
        elif itc == lc-1:
            val5 = 0
        else:
            if (itc+1) <= (lc/2):
                val5 = next((item['THICKNESS'] for item in stackup if item['ID'] == f'CAD_DIELECTRIC{itc}_{itc+1}'),None)
            else:
                val5 = next((item['THICKNESS'] for item in stackup if item['ID'] == f'CAD_DIELECTRIC{itc+1}_{itc+2}'),None)
        val5_f = float(val5) * 1000
        val6 = float(round((val4 + val5_f) / 2.5) * 2.5)
        if itc == 0:
            val7 = 'xxx'
        elif itc == lc-1:
            val7 = 'xxx'
        else:
            if (itc+1) <= (lc/2):
                val7 = next((item['MATERIAL'] for item in stackup if item['ID'] == f'CAD_DIELECTRIC{itc}_{itc+1}'),None)
            else:
                val7 = next((item['MATERIAL'] for item in stackup if item['ID'] == f'CAD_DIELECTRIC{itc+1}_{itc+2}'),None)
        val8 = None
        for dic8 in abfdi:
            if dic8.get('len_n') == 50:
                for key8, value8 in dic8.items():
                    if isinstance(value8, str):
                        if val7 in value8:
                            print('s')
                            print(val6)
                            print((dic8['thick_n']))
                            print('e')
                            if dic8['thick_n'] == val6:
                                val8 = dic8['mtrl_id']
        each_layer = {'layer':itc+1,'portion':val2,'CuT':val3,'eatingT':val4,'target':val5_f,'require':val6,'type':val7,'item_no':val8}
        print(each_layer)
        in_table.append(each_layer)
        itc += 1
    
    print(in_table)



    #upload파일만들기
    upload = []
    




    return stackup, in_table



# 사용 예시
if __name__ == '__main__':
    data1, data2 = extract_drawing('in_a.txt')  # 파일명 변경 가능
    df1 = pd.DataFrame(data1)
    df2 = pd.DataFrame(data2)
    with pd.ExcelWriter('ex_a.xlsx', engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='material', index=False)
        df2.to_excel(writer, sheet_name='material_choice', index=False)
    print("Excel 파일이 성공적으로 생성되었습니다: ex_a.xlsx")
