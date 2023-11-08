import os
import pandas as pd
from framework import PRAD_compiler

#informações dos .docx dos PRAD
path = 'C:/Users/yamad/OneDrive/Documentos/SPVS/tableReader/docs/word/'
files = os.listdir(path)

PRAD_list = list()
prob = list()
for file in files:
    print(f'[PRAD_compiler] Lendo arquivo: {file}...')
    curr_PRAD_obj = PRAD_compiler.dataCompiler(path, file)
    curr_PRAD_df = curr_PRAD_obj.runCompile()
    PRAD_list.append(curr_PRAD_df)
    if len(curr_PRAD_df) == 0:
        prob.append(file)

PRAD_data = pd.concat(PRAD_list, ignore_index=True)
PRAD_data.to_excel('C:/Users/yamad/OneDrive/Documentos/SPVS/tableReader/output/PRAD_completo.xlsx',
                 header=True)

print("[PRAD_compiler] Todos arquivos processados!")
