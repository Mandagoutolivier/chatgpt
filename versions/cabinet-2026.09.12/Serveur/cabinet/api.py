from __future__ import annotations
from contextlib import asynccontextmanager
from pathlib import Path
import os
import logging
import uuid
import hashlib
import psycopg
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError
from pydantic import BaseModel, ConfigDict, Field
from .domain import Refus
from .files import Documents
from .service import Service

class Commande(BaseModel):
    model_config = ConfigDict(extra='forbid')
    operation: str = Field(max_length=80)
    params: dict = Field(default_factory=dict)
    request_id: str = Field(default='',max_length=100)


def service_environnement():
    dsn=os.getenv('CABINET_DATABASE_URL')
    if not dsn:
        from psycopg.conninfo import make_conninfo
        password=Path(os.environ['DB_PASSWORD_FILE']).read_text().strip()
        dsn=make_conninfo(host=os.getenv('DB_HOST','db'),dbname='cabinet',user='cabinet',password=password)
    group = os.getenv('CABINET_ARCHIVE_GID', '').strip()
    return Service(dsn,Documents(Path(os.getenv('CABINET_DATA','/data')),os.environ['CABINET_UNC'],
                                int(group) if group else None))


def create_app(service: Service | None=None):
    @asynccontextmanager
    async def lifespan(app):
        if app.state.service is None:app.state.service=service_environnement()
        app.state.service.initialiser()
        yield
    app=FastAPI(title='Cabinet NAS',version='2026.09.12',lifespan=lifespan,
                docs_url=None,redoc_url=None,openapi_url=None)
    app.state.service=service

    @app.exception_handler(RequestValidationError)
    async def invalid_request(request, exc):
        # La reponse FastAPI par defaut recopie les entrees invalides.
        return JSONResponse({'error': 'Parametres invalides ou incomplets.', 'code': 'contrat'}, status_code=422)

    @app.middleware('http')
    async def limiter(request: Request,call_next):
        length=request.headers.get('content-length','0')
        try:
            if int(length)>4_000_000:return JSONResponse({'error':'Requete trop volumineuse.'},status_code=413)
        except ValueError:return JSONResponse({'error':'Taille invalide.'},status_code=400)
        body=bytearray()
        async for chunk in request.stream():
            if len(body)+len(chunk)>4_000_000:return JSONResponse({'error':'Requete trop volumineuse.'},status_code=413)
            body.extend(chunk)
        # Starlette reutilise ce corps borne lors du passage au routeur.
        request._body=bytes(body)
        response=await call_next(request)
        response.headers['Cache-Control']='no-store'
        response.headers['X-Content-Type-Options']='nosniff'
        return response

    @app.get('/health')
    def health():
        try:
            with app.state.service.connexion() as db:db.execute('SELECT 1')
            return {'status':'ok','version':'2026.09.12','protocole':2}
        except psycopg.Error:return JSONResponse({'status':'indisponible'},status_code=503)

    @app.post('/v1/rpc')
    def rpc(command: Commande,request:Request):
        correlation = uuid.uuid4().hex
        command_hash = hashlib.sha256(command.request_id.encode()).hexdigest()[:16]
        stage = 'authentification'
        def failure(message, status, code):
            logging.getLogger('cabinet').warning('rpc correlation=%s commande=%s etape=%s categorie=%s statut=%s',
                correlation, command_hash, stage, code, status)
            return JSONResponse({'error': message, 'code': code, 'correlation': correlation}, status_code=status)
        try:
            bearer=request.headers.get('Authorization','')
            if not bearer.startswith('Bearer '):raise Refus('Authentification requise.',401)
            actor=app.state.service.compte(bearer[7:])
            stage = 'transaction'
            result=app.state.service.executer(actor,command.operation,command.params,command.request_id)
            return {'result':result}
        except Refus as exc:return failure(str(exc), exc.status, exc.code)
        except psycopg.Error as exc:
            # Ne jamais retourner le SQL ni un detail contenant des donnees.
            state = exc.sqlstate or ''
            if state.startswith('23'):
                return failure('Conflit de donnees : rechargez avant de reessayer.', 409, 'contrainte')
            if isinstance(exc, psycopg.OperationalError) or state in {'40001', '40P01', '55P03', '57014'}:
                return failure('Transaction indisponible. Reessayez avec la meme commande.', 503, 'indisponible')
            return failure('Erreur de base. Le traitement n est pas confirme.', 500, 'base_interne')
        except Exception:
            return failure('Erreur interne. Le traitement n est pas confirme.', 500, 'interne')
    return app

app=create_app()
