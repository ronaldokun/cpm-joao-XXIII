# -*- coding: utf-8 -*-
# ---
# jupyter:
#   jupytext:
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.3'
#       jupytext_version: 0.8.6
#   kernelspec:
#     display_name: Python 3
#     language: python
#     name: python3
# ---

# +
import os 

os.chdir("/home/ronaldo/projects/programming/drive/")

from main import *
# -

# ### Função auxiliar para guardar arquivo em dicionário

# +
def keep_file(file, mimeType=None):
    if mimeType is not None and file['mimeType'] == mimeType:
        file['link'] = "https://drive.google.com/open?id="+file['id']
        return file

def keep_files(files, mimeType=None):
    found = []
    for file in files:
        keep = keep_file(file, mimeType)
        if keep:
            found.append(keep)
    return found
# -

# ## Informe a pasta para varredura

folder = "https://drive.google.com/open?id=1CPILIHf3Bcb8kTzbusAqytmH11GNL70E"
id = folder.split("=")[-1]
t0 = f"'{id}' in parents and mimeType contains 'image'"
t0

files = searchFile(1000, t0)

files[0]

# +
lista = keep_files(files, mimeType='image/jpeg')
                     
lista = sorted(lista, key=lambda x: x['name'])
# -

import pandas as pd
codigos, nomes, links = [],[],[]
for i, aluno in enumerate(lista):
    ident = aluno['name'].split("_", 1)
    if len(ident) < 2:
        break
    codigos.append(ident[0])
    nomes.append(ident[1].replace("_", " ").replace(".jpg", "").replace(".jpeg", ""))
    links.append(aluno['link'])

fotos = pd.DataFrame({'Código': codigos, "Nome": nomes, "Link": links}, index=None)
fotos.head()

fotos.to_csv("fotos_alunos.csv", index=False)

planilhas = []
for l in lista:
    t0 = f"'{l['id']}' in parents"
    files = searchFile(1000, t0)
    for file in files:
        planilhas.append(keep_file(file))

print(file['name'])
drive_service.permissions().list(fileId='1Iphkrw9VEphiGW-Xamlqmzdz2EFSmGPM').execute()

# +
body = {'type': 'domain',
              'role': 'writer',
              'domain': 'cidadaopromundo.org',
              'allowFileDiscovery': True}

for f in planilhas:    
    print(f['name'], f['link'])
    #dict_permissions = drive_service.permissions().list(fileId=f['id']).execute()
    #for perm in dict_permissions['permissions']:
    #    drive_service.permissions().delete(fileId=f['id'], permissionId = perm['id']).execute()
    #break
    drive_service.permissions().create(fileId=f['id'], 
        body=body).execute()
# -

planilhas

# +
extra_types = {}

types = ["Extra", "Grammar", "Project", "Vocabulary", "Writing"]
for t in types:
    extras = {}
    for f in files:
        if t in f['name'] and ".pdf" in f["name"]:
            if f['name'] not in extras:
                extras[f['name']] = "https://drive.google.com/open?id="+f['id']
                
    extra_types[t] = extras
# -

#extras = sorted(tuple(extras.items()), key=lambda x: x[0])
for t,v in extra_types.items():
    print(t)
    tupla = sorted(tuple(v.items()), key=lambda x: x[0])
    for value in tupla:
        print(value[1])


