from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Optional
import io
from excel_generator import generate_excel

app = FastAPI(title="Flujo de Caja API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class InversionItem(BaseModel):
    nombre: str
    monto: float
    vida_util: int

class CostoFijo(BaseModel):
    tipo: str
    costo_mensual: float
    observacion: Optional[str] = ""

class CostoVariable(BaseModel):
    tipo: str
    costo: float
    observacion: Optional[str] = ""

class ModeloNegocio(BaseModel):
    socios_estrategicos: Optional[str] = ""
    recursos_clave: Optional[str] = ""
    actividades_clave: Optional[str] = ""
    propuesta_valor: Optional[str] = ""
    relacion_cliente: Optional[str] = ""
    canales: Optional[str] = ""
    segmento_clientes: Optional[str] = ""
    estructura_costos: Optional[str] = ""
    flujo_ingresos: Optional[str] = ""
    barreras_entrada: Optional[str] = ""

class FlujoCajaData(BaseModel):
    nombre_proyecto: Optional[str] = "Proyecto"
    inversiones: List[InversionItem]
    costos_fijos: List[CostoFijo]
    costos_variables: List[CostoVariable]
    precio_venta: float
    cantidades_venta: List[float]  # años 1-5
    tasa_interes: float = 0.20
    tasa_cb: float = 0.15
    modelo_negocio: Optional[ModeloNegocio] = None

@app.post("/generar-excel")
def generar_excel(data: FlujoCajaData):
    output = generate_excel(data)
    filename = f"Flujo_de_Caja_{data.nombre_proyecto.replace(' ', '_')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/health")
def health():
    return {"status": "ok"}
