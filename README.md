# 🗺️ Control Catastral – Líder SIG

Dashboard de control del proceso de actualización catastral para el Líder SIG.
Desarrollado para GEO Proyecciones · Soledad, Atlántico.

## Cómo usar

1. Descarga el Excel del aplicativo LíderSIG
2. Súbelo en el panel izquierdo del dashboard
3. El cuadro de control se actualiza automáticamente

## Despliegue en Streamlit Community Cloud

1. Sube esta carpeta a un repositorio de GitHub (puede ser privado)
2. Ve a https://share.streamlit.io
3. Conecta tu cuenta de GitHub
4. Selecciona el repositorio y el archivo `app.py`
5. Haz clic en Deploy

## Estructura

```
lider_sig_dashboard/
├── app.py            # Aplicación principal
├── requirements.txt  # Dependencias
└── README.md
```

## Estados del proceso

| Estado | Descripción |
|--------|-------------|
| Pendiente Líder SIG | Sin asignar a coordinador |
| En proceso Editor SIG | Editor capturando en campo |
| Terminado Editor SIG | Editor entregó al coordinador |
| Aprobado Coordinador | Coordinador aprobó |
| Rechazado Coordinador | Coordinador devolvió al editor |
| Devuelto a reconocedor | Gestión manual del Líder SIG |
