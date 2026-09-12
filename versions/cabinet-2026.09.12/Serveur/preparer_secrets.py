"""Preparation locale NAS, refuse tout ecrasement de secret."""
from pathlib import Path
import os
import secrets

def main():
    root=Path(__file__).resolve().parent/'secrets'
    root.mkdir(mode=0o700,exist_ok=True)
    for name in ('admin_password.txt','db_password.txt'):
        path=root/name
        if path.exists():
            print(name+' existe : conserve.');continue
        fd=os.open(path,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
        with os.fdopen(fd,'w') as f:f.write(secrets.token_urlsafe(48)+'\n')
        print(name+' cree. Aucun secret affiche.')
    print('Configurer les droits avant demarrage : admin_password.txt 999:999 mode 600 ; db_password.txt 999:GID_API mode 640. Voir INSTALLATION_NAS.md.')
if __name__=='__main__':main()
