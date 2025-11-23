from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
import tempfile
import os
import shutil

# Importar nuestros módulos
import sys
sys.path.append('utils')
from excel_processor import ExcelProcessor
from drive_handler import DriveHandler

# CUITMapper está en la raíz
from cuit_mapper import CUITMapper

app = FastAPI(title="Procesador de Libros IVA")

# Configurar CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Servir archivos estáticos
app.mount("/static", StaticFiles(directory="static"), name="static")

# Instancia global del handler de Drive y mapper
drive_handler = DriveHandler()
cuit_mapper = CUITMapper()

# Modelos Pydantic
class ProcessRequest(BaseModel):
    client: str
    tipo: str  # "ventas" o "compras"
    year: int
    month: int
    month_name: str


@app.on_event("startup")
async def startup_event():
    """Inicializar la conexión con Drive al arrancar"""
    try:
        print("\n🔄 Verificando autenticación con Google Drive...")
        drive_handler.authenticate()
        print("✅ Conexión con Google Drive establecida\n")
    except Exception as e:
        print(f"⚠️  Advertencia: Error al conectar con Drive: {str(e)}")
        print("💡 Se solicitará autenticación cuando sea necesario\n")


@app.get("/")
async def root():
    """Servir el frontend"""
    return FileResponse("static/index.html")


@app.get("/api/clients")
async def get_clients():
    """Obtener lista de clientes desde el mapeo CUIT"""
    try:
        # Obtener todos los clientes del mapeo
        all_clients = cuit_mapper.get_all_clients()
        
        # Convertir a formato esperado por el frontend
        clients = [
            {
                'name': name,
                'id': cuit,
                'enabled': True  # Todos habilitados
            }
            for cuit, name in all_clients.items()
        ]
        
        # Ordenar alfabéticamente
        clients.sort(key=lambda x: x['name'].lower())
        
        return {"success": True, "clients": clients}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/detect-month")
async def detect_month(file: UploadFile = File(...)):
    """Detecta el mes y año del archivo Excel"""
    try:
        # Guardar archivo temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(await file.read())
            tmp_path = tmp.name
        
        # Procesar
        processor = ExcelProcessor(tmp_path)
        processor.read_excel()
        month, year = processor.detect_month()
        month_name = processor.get_month_name(month)
        
        # Limpiar
        os.unlink(tmp_path)
        
        return {
            "success": True,
            "month": month,
            "year": year,
            "month_name": month_name
        }
    except Exception as e:
        if 'tmp_path' in locals():
            os.unlink(tmp_path)
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/process")
async def process_file(
    file: UploadFile = File(...),
    client: str = Form(...),
    tipo: str = Form(...),
    year: int = Form(...),
    month: int = Form(...),
    month_name: str = Form(...),
    create_if_not_exists: bool = Form(False)
):
    """Procesa el archivo y lo sube a Drive"""
    
    tmp_input = None
    tmp_output = None
    downloaded_file = None
    
    try:
        print(f"\n{'='*60}")
        print(f"INICIANDO PROCESAMIENTO")
        print(f"Cliente: {client}, Tipo: {tipo}, Año: {year}, Mes: {month_name}")
        print(f"{'='*60}\n")
        
        # 1. Guardar archivo subido
        print("📁 Guardando archivo temporal...")
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(await file.read())
            tmp_input = tmp.name
        print(f"   ✓ Archivo guardado: {tmp_input}")
        
        # 2. Procesar Excel
        print("\n🔄 Procesando Excel...")
        processor = ExcelProcessor(tmp_input)
        processor.read_excel()
        df_clean = processor.clean_data()
        print(f"   ✓ Procesado: {len(df_clean)-1} filas, {len(df_clean.columns)} columnas")
        
        # 3. Verificar si existe el archivo del año en Drive
        print(f"\n🔍 Verificando archivo en Drive...")
        file_check = drive_handler.check_year_file_exists(client, tipo, year)
        
        if not file_check['exists']:
            print(f"   ⚠ Archivo '{file_check['filename']}' NO existe")
            if not create_if_not_exists:
                return {
                    "success": False,
                    "needs_confirmation": True,
                    "message": f"El archivo '{file_check['filename']}' no existe. ¿Desea crearlo?"
                }
            else:
                # Crear el archivo del año
                print(f"   🆕 Creando archivo '{file_check['filename']}'...")
                file_id = drive_handler.create_year_file(client, tipo, year)
                file_check['file_id'] = file_id
                file_check['exists'] = True
                print(f"   ✓ Archivo creado: {file_id}")
        else:
            print(f"   ✓ Archivo encontrado: {file_check['file_id']}")
        
        # 4. Descargar el archivo del año desde Drive
        print(f"\n⬇️  Descargando '{file_check['filename']}' desde Drive...")
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            downloaded_file = tmp.name
        
        drive_handler.download_file(file_check['file_id'], downloaded_file)
        print(f"   ✓ Descargado a: {downloaded_file}")
        
        # 5. Agregar pestaña al libro del año
        print(f"\n📝 Agregando pestaña '{month_name}'...")
        result = processor.add_sheet_to_workbook(downloaded_file, month_name, df_clean)
        
        if result == "exists":
            raise HTTPException(
                status_code=400,
                detail=f"La pestaña '{month_name}' ya existe en el archivo '{file_check['filename']}'"
            )
        print(f"   ✓ Pestaña agregada")
        
        # 6. Subir el archivo actualizado a Drive
        print(f"\n⬆️  Subiendo archivo actualizado a Drive...")
        drive_handler.update_file(file_check['file_id'], downloaded_file)
        print(f"   ✓ Archivo actualizado en Drive")
        
        # 7. Limpiar archivos temporales
        print(f"\n🧹 Limpiando archivos temporales...")
        
        # Cerrar cualquier referencia al archivo antes de eliminar
        import gc
        gc.collect()
        
        # Intentar eliminar con retry para Windows
        import time
        for attempt in range(3):
            try:
                if tmp_input and os.path.exists(tmp_input):
                    os.unlink(tmp_input)
                if downloaded_file and os.path.exists(downloaded_file):
                    os.unlink(downloaded_file)
                break
            except PermissionError:
                if attempt < 2:
                    time.sleep(0.5)  # Esperar un poco
                else:
                    print(f"   ⚠️  No se pudieron eliminar archivos temporales (se eliminarán después)")
        
        print(f"   ✓ Limpieza completada")
        
        print(f"\n{'='*60}")
        print(f"✅ PROCESAMIENTO COMPLETADO EXITOSAMENTE")
        print(f"{'='*60}\n")
        
        return {
            "success": True,
            "message": f"Pestaña '{month_name}' agregada exitosamente al archivo '{file_check['filename']}'",
            "rows_processed": len(df_clean) - 1,
            "client": client,
            "tipo": tipo,
            "year": year,
            "month": month_name
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        
        # Limpiar en caso de error
        import time
        import gc
        gc.collect()
        
        for attempt in range(2):
            try:
                if tmp_input and os.path.exists(tmp_input):
                    os.unlink(tmp_input)
                if downloaded_file and os.path.exists(downloaded_file):
                    os.unlink(downloaded_file)
                break
            except PermissionError:
                if attempt < 1:
                    time.sleep(0.5)
        
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/preview")
async def preview_file(file: UploadFile = File(...)):
    """Genera una vista previa del archivo procesado"""
    tmp_path = None
    try:
        # Guardar archivo temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(await file.read())
            tmp_path = tmp.name
        
        # Procesar
        processor = ExcelProcessor(tmp_path)
        processor.read_excel()
        df_clean = processor.clean_data()
        
        # Convertir a dict para JSON (primeras 10 filas + última si es totales)
        if len(df_clean) > 10:
            # Mostrar primeras 10 + la fila de totales
            preview_rows = list(range(10)) + [len(df_clean) - 1]
            preview_df = df_clean.iloc[preview_rows]
        else:
            preview_df = df_clean
        
        # Convertir NaN a None para JSON
        preview_data = preview_df.fillna('').to_dict('records')
        columns = df_clean.columns.tolist()
        total_rows = len(df_clean) - 1  # -1 por la fila de totales
        
        # Limpiar
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        
        return {
            "success": True,
            "preview": preview_data,
            "columns": columns,
            "total_rows": total_rows,
            "columns_kept": len(columns)
        }
    except Exception as e:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/create-client")
async def create_client(
    client_name: str = Form(...),
    cuit: str = Form(...)
):
    """Crea un nuevo cliente en Drive"""
    try:
        print(f"\n{'='*60}")
        print(f"CREANDO NUEVO CLIENTE")
        print(f"Nombre: {client_name}, CUIT: {cuit}")
        print(f"{'='*60}\n")
        
        # Verificar si el CUIT ya existe
        print(f"📋 Verificando si CUIT ya existe...")
        if cuit_mapper.client_exists(cuit):
            existing_name = cuit_mapper.get_client_by_cuit(cuit)
            print(f"   ⚠️  CUIT ya existe para: {existing_name}")
            raise HTTPException(
                status_code=400,
                detail=f"El CUIT {cuit} ya está registrado para el cliente '{existing_name}'"
            )
        
        print(f"   ✓ CUIT disponible")
        
        # Crear cliente en Drive
        print(f"\n📁 Creando estructura en Drive...")
        result = drive_handler.create_client(client_name, cuit)
        
        if not result['success']:
            # Si el cliente ya existe en Drive, preguntamos si solo quiere agregarlo al mapeo
            if "ya existe" in result['message'].lower():
                print(f"   ⚠️  El cliente ya existe en Drive")
                print(f"   💡 Agregando solo al mapeo local...")
                
                # Verificar que tenga la estructura correcta
                try:
                    structure = drive_handler.get_client_structure(client_name)
                    if structure['ventas_id'] and structure['compras_id']:
                        # Todo OK, agregar al mapeo
                        cuit_mapper.add_client(cuit, client_name)
                        print(f"   ✓ Cliente agregado al mapeo")
                        
                        return {
                            "success": True,
                            "message": f"Cliente '{client_name}' agregado al sistema (ya existía en Drive)",
                            "client_name": client_name,
                            "cuit": cuit
                        }
                except Exception as e:
                    print(f"   ❌ Error verificando estructura: {str(e)}")
            
            print(f"   ❌ Error: {result['message']}")
            raise HTTPException(status_code=400, detail=result['message'])
        
        print(f"   ✓ Estructura creada en Drive")
        
        # Guardar en el mapeo
        print(f"\n💾 Guardando en mapeo CUIT...")
        cuit_mapper.add_client(cuit, client_name)
        print(f"   ✓ Guardado en cuit_mapping.json")
        
        # Verificar que se guardó
        saved_client = cuit_mapper.get_client_by_cuit(cuit)
        print(f"   ✓ Verificación: {saved_client}")
        
        print(f"\n{'='*60}")
        print(f"✅ CLIENTE CREADO EXITOSAMENTE")
        print(f"{'='*60}\n")
        
        return {
            "success": True,
            "message": result['message'],
            "client_name": client_name,
            "cuit": cuit
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/edit-client")
async def edit_client(
    old_cuit: str = Form(...),
    new_name: str = Form(...),
    new_cuit: str = Form(...)
):
    """Edita un cliente existente"""
    try:
        print(f"\n{'='*60}")
        print(f"EDITANDO CLIENTE")
        print(f"CUIT anterior: {old_cuit}")
        print(f"Nuevo nombre: {new_name}, Nuevo CUIT: {new_cuit}")
        print(f"{'='*60}\n")
        
        # Verificar que el cliente existe
        old_name = cuit_mapper.get_client_by_cuit(old_cuit)
        if not old_name:
            raise HTTPException(status_code=404, detail="Cliente no encontrado")
        
        print(f"   📋 Cliente actual: {old_name}")
        
        # Si cambió el CUIT, verificar que el nuevo no exista
        if old_cuit != new_cuit:
            if cuit_mapper.client_exists(new_cuit):
                existing = cuit_mapper.get_client_by_cuit(new_cuit)
                raise HTTPException(
                    status_code=400,
                    detail=f"El CUIT {new_cuit} ya está registrado para '{existing}'"
                )
        
        # Eliminar el viejo
        all_clients = cuit_mapper.get_all_clients()
        if old_cuit in all_clients:
            del all_clients[old_cuit]
        
        # Agregar el nuevo
        all_clients[new_cuit] = new_name
        
        # Guardar
        cuit_mapper.mapping = all_clients
        cuit_mapper.save_mapping()
        
        print(f"   ✓ Cliente actualizado")
        print(f"\n{'='*60}")
        print(f"✅ EDICIÓN COMPLETADA")
        print(f"{'='*60}\n")
        
        return {
            "success": True,
            "message": f"Cliente actualizado exitosamente",
            "client_name": new_name,
            "cuit": new_cuit
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/auto-detect-all")
async def auto_detect_all(file: UploadFile = File(...)):
    """Detecta automáticamente: cliente (por CUIT), tipo, mes y año"""
    tmp_path = None
    try:
        print(f"\n{'='*60}")
        print(f"AUTO-DETECTAR TODO")
        print(f"{'='*60}\n")
        
        # Guardar archivo temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(await file.read())
            tmp_path = tmp.name
        
        print(f"📁 Archivo guardado: {tmp_path}")
        
        # Procesar
        processor = ExcelProcessor(tmp_path)
        
        # Detectar info del header (CUIT y tipo)
        print(f"\n🔍 Detectando CUIT y tipo...")
        header_info = processor.detect_info_from_header()
        
        # Buscar cliente por CUIT
        print(f"\n👤 Buscando cliente con CUIT {header_info['cuit']}...")
        client_name = cuit_mapper.get_client_by_cuit(header_info['cuit'])
        
        if client_name:
            print(f"   ✓ Cliente encontrado: {client_name}")
        else:
            print(f"   ⚠️  Cliente no encontrado en mapeo")
        
        # Detectar mes y año
        print(f"\n📅 Detectando mes y año...")
        processor.read_excel()
        month, year = processor.detect_month()
        month_name = processor.get_month_name(month)
        
        print(f"   ✓ Detectado: {month_name} {year}")
        
        # Limpiar
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        
        print(f"\n{'='*60}")
        print(f"✅ DETECCIÓN COMPLETADA")
        print(f"{'='*60}\n")
        
        return {
            "success": True,
            "cuit": header_info['cuit'],
            "client": client_name,
            "client_found": client_name is not None,
            "tipo": header_info['tipo'],
            "month": month,
            "year": year,
            "month_name": month_name
        }
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/health")
async def health_check():
    """Verifica el estado de la API y la conexión con Drive"""
    try:
        drive_handler.authenticate()
        clients_count = len(cuit_mapper.get_all_clients())
        return {
            "status": "healthy",
            "drive_connected": True,
            "clients_count": clients_count
        }
    except Exception as e:
        return {
            "status": "degraded",
            "drive_connected": False,
            "error": str(e)
        }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
