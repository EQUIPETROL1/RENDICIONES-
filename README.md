# Sistema de Rendición de Gastos

Aplicación web donde cada trabajador sube sus facturas y la IA extrae los datos automáticamente. El administrador puede descargar un Excel consolidado por trabajador.

## Estructura de carpetas generada automáticamente

```
rendiciones/
├── app.py                  ← Backend Flask
├── requirements.txt        ← Dependencias Python
├── templates/
│   ├── index.html          ← Vista del trabajador
│   ├── admin.html          ← Panel del administrador
│   └── admin_login.html    ← Login del admin
├── facturas/               ← Se crea sola al subir facturas
│   ├── <worker_id>/
│   │   ├── invoices.json   ← Datos extraídos de ese trabajador
│   │   └── *.pdf / *.jpg   ← Archivos originales
│   └── workers.json        ← Registro de trabajadores
└── rendiciones_excel/      ← Excels generados
```

## Instalación

### 1. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 2. Configurar variables de entorno
Crea un archivo `.env` o exporta estas variables:

```bash
export ANTHROPIC_API_KEY="sk-ant-..."       # Tu API key de Anthropic
export ADMIN_PASSWORD="tu_clave_segura"     # Contraseña del panel admin (default: admin123)
export SECRET_KEY="clave_aleatoria_larga"   # Clave secreta Flask
```

### 3. Ejecutar localmente
```bash
python app.py
```
Abre http://localhost:5000

## Despliegue en servidor (producción)

### Con Gunicorn (recomendado)
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

### Con Render.com (gratis)
1. Sube el proyecto a GitHub
2. En Render: New → Web Service → selecciona el repo
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`
5. Agrega las variables de entorno en el panel de Render

### Con Railway.app
```bash
railway init
railway up
```

### Con servidor propio (Ubuntu/Debian)
```bash
# Instalar nginx
sudo apt install nginx python3-pip

# Instalar dependencias
pip3 install -r requirements.txt gunicorn

# Correr como servicio
gunicorn -w 4 -b 127.0.0.1:8000 app:app &

# Configurar nginx como proxy
# /etc/nginx/sites-available/rendiciones
server {
    listen 80;
    server_name tu-dominio.com;
    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
    }
}
```

## Uso

### Trabajadores
1. Entrar a la URL de la app
2. Ingresar nombre completo (la próxima vez se recuperan sus facturas)
3. Subir facturas — la IA extrae todos los datos automáticamente
4. Corregir cualquier campo haciendo doble clic
5. Descargar su Excel de rendición individual

### Administrador
1. Ir a `/admin`
2. Ingresar contraseña (default: `admin123`)
3. Ver resumen de todos los trabajadores
4. Descargar Excel individual por trabajador o Excel completo con todos

## Formatos soportados
- PDF (facturas electrónicas, escaneadas)
- JPG / PNG (fotos de facturas)

## Datos extraídos automáticamente
- N° de factura
- Fecha de emisión y vencimiento
- Proveedor y RUC proveedor
- Cliente y RUC cliente
- Descripción del servicio/producto
- Base gravada, IGV y total
- Moneda y forma de pago
