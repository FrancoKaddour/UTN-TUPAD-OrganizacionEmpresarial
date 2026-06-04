# ChatBot-Proveedores-LB

Bot de WhatsApp para la gestion interna de proveedores de **La Basica Pasteleria**.  
Desarrollado como Trabajo Practico Integrador de Organizacion Empresarial — UTN TUPaD.

---

## Funcionalidades

| Opcion | Funcion |
|--------|---------|
| 1 | Ver proveedores (todos o por categoria) |
| 2 | Dar de alta un proveedor |
| 3 | Actualizar datos de un proveedor |
| 4 | Dar de baja un proveedor |

---

## Requisitos previos

- Python 3.10 o superior
- Cuenta en [Twilio](https://www.twilio.com) (plan gratuito)
- [ngrok](https://ngrok.com) instalado

---

## Instalacion

```bash
# 1. Clonar el repositorio
git clone https://github.com/tuusuario/ChatBot-Proveedores-LB.git
cd ChatBot-Proveedores-LB

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Generar la base de datos Excel con proveedores de muestra
python crear_excel.py
```

---

## Configuracion de Twilio Sandbox

1. Ir a [console.twilio.com](https://console.twilio.com) → Messaging → Try it out → Send a WhatsApp message
2. Escanear el QR o enviar el codigo de activacion desde tu WhatsApp
3. Copiar el **Account SID** y el **Auth Token** desde el dashboard (no son necesarios en el codigo actual, se usan para enviar mensajes proactivos)

---

## Ejecucion

### Terminal 1 — Levantar el servidor Flask

```bash
python bot.py
```

El servidor corre en `http://localhost:5000`

### Terminal 2 — Exponer con ngrok

```bash
ngrok http 5000
```

Copiá la URL `https://XXXX.ngrok.io` que aparece.

### Configurar webhook en Twilio

1. Ir a Twilio → Sandbox Settings
2. En **"When a message comes in"** pegar:  
   `https://XXXX.ngrok.io/webhook`
3. Metodo: **HTTP POST**
4. Guardar

---

## Uso

Desde WhatsApp, enviá un mensaje al numero de Twilio Sandbox.  
Escribi **hola** para iniciar el menu.

---

## Estructura del proyecto

```
ChatBot-Proveedores-LB/
├── bot.py              # Logica del chatbot + servidor Flask
├── crear_excel.py      # Script de inicializacion de la BD Excel
├── proveedores.xlsx    # Base de datos (se genera con crear_excel.py)
├── requirements.txt    # Dependencias Python
└── README.md
```

---

## Estructura del Excel (`proveedores.xlsx`)

| Columna | Descripcion |
|---------|-------------|
| ID | Identificador unico (P001, P002...) |
| Nombre | Nombre del proveedor |
| Categoria | Panaderia / Reposteria / Bebidas / Fiambres y Sandwicheria / Insumos de Cocina |
| Productos | Lista de productos que provee |
| Telefono | Numero de contacto |
| Dias_Entrega | Dias habituales de entrega |
| Forma_Pago | Efectivo / Transferencia / Credito 30 dias |
| Estado | Activo / Inactivo |
| Fecha_Alta | Fecha de registro |
| Ultima_Modificacion | Ultima actualizacion |

---

## Integrantes

- Franco Kaddour — Desarrollo
- Gonzalo Isaias — Repositorio y Jira

**UTN TUPaD — Organizacion Empresarial — Comision 11 — 2026**
