# Correos automáticos
-En construcción

## Características/funciones


## Estructura del proyecto

```plaintext
├── README.md                      # Descripción y guía de uso
│
├── requirements.txt               # Dependencias del proyecto
│
├── modules/                       # Módulo con las clases y funciones auxiliares
│   ├── outlook_manager.py          # Manejo y envío de correos electrónicos Outlook
│   ├── sharepoint_manager.py       # Funciones de integración con SharePoint
│   └── file_manager.py             # Gestión de archivos (renombrar, clasificar)
│
├── scripts/                       # Módulo con las funciones principales
│   ├── main.py                     # Manejo y envío de correos electrónicos
│   └── email_to_dataframe.py       # Convierte correos filtrados a excel
│
├── templates/                     # Plantillas de correo
│   └── sharepoint_success.html     # Plantilla para correos de confirmación
│
├── images/                        # Images / misc
│
├── logs/                          # Carpeta para guardar logs del proceso
│
└── descargas/                     # Carpeta temporal para descargas



