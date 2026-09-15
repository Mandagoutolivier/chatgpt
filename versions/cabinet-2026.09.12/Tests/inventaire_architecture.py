"""Inventaire reproductible des composants et liaisons visibles dans les sources.

Ce controle ne recense pas les commandes Dragon installees sur un poste ni les
appels construits a l'execution. Les points d'entree publics sont donc conserves.
"""
from pathlib import Path
import hashlib
import json
import re
import sys
from audit_statique import ROOT, sans_litteraux


def inventorier(root=ROOT):
    manifest=json.loads((root/'Build/manifest.json').read_text())
    composants=[];liens=[];dynamiques=[];erreurs=[];declares=set()
    for host in ['word','excel']:
        for role,key in [('production',host),('recette',host+'_recette')]:
            for item in manifest.get(key,[]):
                path=root/item['path'];declares.add(item['path'])
                text=path.read_text(encoding='utf-8-sig')
                clean='\n'.join(sans_litteraux(line) for line in text.splitlines())
                publics=sorted(set(re.findall(r'(?im)^\s*(?:Public )?(?:Sub|Function)\s+(\w+)',clean)))
                composants.append({'hote':host,'role':role,'module':item['name'],'source':item['path'],
                                   'lignes':len(text.splitlines()),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
                                   'entrees_publiques':publics})
                for target,member in sorted(set(re.findall(r'\b(mod\w+)\.(\w+)',clean))):
                    liens.append({'hote':host,'source':item['name'],'cible':target,'membre':member})
                for n,line in enumerate(clean.splitlines(),1):
                    if re.search(r'\b(?:Application\.Run|CallByName|OnAction|KeyBindings|OnTime)\b',line,re.I):
                        dynamiques.append({'source':item['path'],'ligne':n})
    for folder in ['Src','Tests/Vba']:
        for path in (root/folder).rglob('*'):
            if path.suffix in {'.bas','.cls','.vba'} and path.relative_to(root).as_posix() not in declares:
                erreurs.append('Composant non classe : '+path.relative_to(root).as_posix())
    archives=[]
    for path in sorted((root/'Archives/Vba').glob('*.bas')):
        text=path.read_text(encoding='utf-8-sig');name=re.search(r'Attribute VB_Name = "([^"]+)"',text)[1]
        archives.append({'module':name,'source':path.relative_to(root).as_posix(),'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
        for c in composants:
            text=(root/c['source']).read_text(encoding='utf-8-sig')
            if re.search(r'\b'+re.escape(name)+r'\b','\n'.join(sans_litteraux(l) for l in text.splitlines()),re.I):
                erreurs.append('Reference vers archive : '+c['module']+' -> '+name)
    return {'composants':composants,'liaisons_qualifiees':liens,'appels_dynamiques_a_qualifier':dynamiques,
            'archives_exclues':archives,'erreurs':erreurs,
            'limite':'Inventaire de sources ; compilation Office et commandes Dragon du poste non verifiees.'}


if __name__=='__main__':
    result=inventorier()
    if '--json' in sys.argv:print(json.dumps(result,ensure_ascii=False,indent=2))
    else:
        for error in result['erreurs']:print('FAIL',error)
        print(len(result['composants']),'composants classes ;',len(result['archives_exclues']),'archives ;',len(result['erreurs']),'erreur(s).')
    raise SystemExit(bool(result['erreurs']))
