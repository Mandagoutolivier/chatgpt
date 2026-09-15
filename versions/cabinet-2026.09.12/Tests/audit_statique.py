"""Contrats statiques du paquet. Ne remplace pas le compilateur VBA d'Office."""
from pathlib import Path
import re, json, hashlib, sys, zipfile
ROOT=Path(__file__).resolve().parents[1]

def sans_litteraux(line):
    out=[]; i=0; quoted=False
    while i<len(line):
        ch=line[i]
        if ch=='"':
            if quoted and i+1<len(line) and line[i+1]=='"':i+=2;continue
            quoted=not quoted;out.append('"')
        elif not quoted:
            if ch=="'":break
            out.append(ch)
        i+=1
    if quoted:raise ValueError('chaine VBA non terminee: '+line[:100])
    return ''.join(out)

def verifier(recette=False):
    errors=[];manifest=json.loads((ROOT/'Build/manifest.json').read_text())
    for name,sha in manifest['models'].items():
        path=ROOT/'ModelesSource'/name
        if hashlib.sha256(path.read_bytes()).hexdigest()!=sha:errors.append('empreinte '+name)
        with zipfile.ZipFile(path) as z:
            if z.testzip():errors.append('archive '+name)
    assets=json.loads((ROOT/'Build/donnees_initiales.json').read_text())
    for path,sha in assets.items():
        if hashlib.sha256((ROOT/'DonneesInitiales'/path).read_bytes()).hexdigest()!=sha:errors.append('ressource '+path)
    inventory=[]
    for host in ['word','excel']:
        names={};public={};symbols={};texts={}
        for item in manifest[host] + (manifest.get(host+'_recette',[]) if recette else []):
            name=item['name'].lower();path=ROOT/item['path']
            if name in names:errors.append(host+' composant double '+name)
            names[name]=item
            if not path.is_file():errors.append('absent '+str(path));continue
            text=path.read_text(encoding='utf-8-sig')
            module=re.search(r'^Attribute VB_Name = "([^"]+)"',text)
            if not module or module[1].lower()!=name:errors.append('nom du composant '+str(path))
            lines=[]
            try:lines=[sans_litteraux(l) for l in text.splitlines() if not l.startswith('Attribute ')]
            except ValueError as e:errors.append(str(path)+': '+str(e))
            code=re.sub(r'_[ \t]*\n[ \t]*',' ', '\n'.join(lines));texts[name]=code
            syms=set(re.findall(r'(?im)^\s*(?:Public |Private |Friend )?(?:Sub|Function|Property (?:Get|Let|Set))\s+(\w+)',code))
            syms.update(re.findall(r'(?im)^\s*Public (?:Const )?(\w+)\s+As\b',code))
            symbols[name]={x.lower() for x in syms}
            if item['kind']=='module':
                for p in re.findall(r'(?im)^\s*(?:Public )?(?:Sub|Function)\s+(\w+)',code):public.setdefault(p.lower(),[]).append(item['name'])
            starts=len(re.findall(r'(?im)^\s*(?:Public |Private |Friend )?(?:Sub|Function|Property (?:Get|Let|Set))\s+\w+',code))
            ends=len(re.findall(r'(?im)^\s*End (?:Sub|Function|Property)\s*$',code))
            if starts!=ends:errors.append('procedures non equilibrees '+str(path))
            inventory.append({'role':'recette' if item in manifest.get(host+'_recette',[]) else 'production','host':host,'module':item['name'],'source':item['path'],'lines':len(text.splitlines()),'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
        for name,defs in public.items():
            if len(defs)>1:errors.append(host+' symbole public ambigu '+name+' '+str(defs))
        for name,code in texts.items():
            for module,member in re.findall(r'\b(mod[A-Z][A-Za-z0-9_]*)\.(\w+)',code):
                if module.lower() not in names:errors.append(name+' module absent '+module)
                elif member.lower() not in symbols[module.lower()]:errors.append(name+' membre absent '+module+'.'+member)
    # L'injection de sources doit consommer le manifeste complet dans les deux hotes.
    for host,file in [('word','construire_modele_unifie.ps1'),('excel','construire_cabinet_secretariat.ps1')]:
        if '$manifest.'+host not in (ROOT/'Build'/file).read_text():errors.append('manifeste ignore '+host)
    return errors,inventory

if __name__=='__main__':
    errors,inventory=verifier(recette='--recette' in sys.argv)
    if '--inventory' in sys.argv:print(json.dumps(inventory,indent=2,ensure_ascii=False))
    else:
        for error in errors:print('FAIL',error)
        print(f'{len(inventory)} composants declares, {sum(i["lines"] for i in inventory)} lignes, {len(errors)} erreur(s) statique(s).')
    sys.exit(bool(errors))
