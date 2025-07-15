import re
import pandas as pd

def read_stkup_sr(TB, lines):
    if TB == 'TOP':
        srtb = 'SRT'
    else:
        srtb = 'SRB'
            
    sr_stk = {'ID': f'CAD_+{TB}+_SOLDERMASK', 'TYPE': srtb, 'MATERIAL': '', 'THICKNESS': '', 'TOLERANCE': '', 'VIA FILL MATERIAL': ''}
    for isr in range(len({lines})):
        line = {lines}[isr].strip()
        if line.startswith(f'CAD_{TB}_SOLDERMASK') and any(key in line for key in ['MATERIAL', 'THICKNESS', 'TOLERANCE', 'FILL_MATERIAL']):
            if {lines}[isr+1].strip() == '9':
                srtvalue = {lines}[isr+4].strip()
                match line:
                    case _ if 'SOLDERMASK_MATERIAL' in line:
                        sr_stk['MATERIAL'] = srtvalue
                    case _ if 'THICKNESS' in line:
                        sr_stk['THICKNESS'] = srtvalue
                    case _ if 'TOLERANCE' in line:
                        sr_stk['TOLERANCE'] = srtvalue
                    case _:
                        sr_stk['VIA FILL MATERIAL'] = srtvalue
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
