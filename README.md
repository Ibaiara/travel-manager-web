# Travel Manager — Web Demo

Web demo de una aplicación de gestión de viajes corporativos, adaptada desde una aplicación de escritorio desarrollada originalmente con Electron.

Este proyecto forma parte de mi portfolio personal y tiene como objetivo mostrar la **lógica, arquitectura y flujos** de una herramienta real utilizada en un entorno empresarial.

---

## 🧭 Contexto del proyecto

En muchas empresas, la gestión de viajes de operarios o comerciales se realiza mediante procesos poco estructurados:
- hojas de cálculo dispersas
- emails manuales
- falta de trazabilidad
- dificultad para visualizar desplazamientos y costes

Esta aplicación se desarrolló para **centralizar y simplificar la gestión de viajes**, mejorando la organización interna y reduciendo errores administrativos.

La versión original se implementó como **aplicación de escritorio con Electron**.  
Para el portfolio, se ha creado esta **versión web de demostración**, reutilizando la lógica principal del backend.

---

## ✅ Qué muestra esta demo

La demo web permite probar los flujos principales de la aplicación:

- Listado de viajes
- Creación de nuevos viajes
- Edición y eliminación de viajes
- Gestión de ubicaciones
- Visualización básica de información de desplazamientos

⚠️ Algunas funcionalidades avanzadas (importación masiva desde Excel, permisos, usuarios reales) no están expuestas en la demo pública por motivos de seguridad y simplicidad.

---

## 🧱 Arquitectura

La aplicación está dividida en dos partes claras:

### Frontend
- HTML, CSS y JavaScript
- Interfaz sencilla orientada a demostrar flujos
- Comunicación con backend mediante API REST

### Backend
- Node.js + Express
- Persistencia local en archivos JSON
- Sistema de backups automáticos
- Control básico de concurrencia mediante locks
- API REST clara y modular

La lógica de negocio es compartida con la versión Electron original, lo que permite:
- reutilización de código
- consistencia de comportamiento
- adaptación a distintos entornos (desktop / web)

---

## 🌐 Demo online

La demo web está desplegada en la nube y es accesible sin necesidad de instalar nada:

👉 **Demo:**  
https://travel-manager-web.onrender.com

> Nota: al estar desplegada en un plan gratuito, la primera carga puede tardar unos segundos si el servicio está inactivo.

---

## 💻 Ejecutar en local

Si quieres ejecutar la demo en tu propio equipo:

```bash
npm install
npm start
El servidor se iniciará en:

javascript
http://localhost:3001
🛠️ Tecnologías utilizadas
Node.js
Express
JavaScript
Persistencia en JSON
HTML / CSS
Despliegue en Render
📌 Sobre el uso de Electron
La aplicación original fue desarrollada como aplicación de escritorio con Electron para un entorno corporativo real.

Esta versión web:

no reemplaza a la aplicación de escritorio
existe únicamente como demo accesible
reutiliza la lógica central del proyecto original
Esto permite mostrar el proyecto sin necesidad de descargar ejecutables.

📬 Contacto
Si quieres más información sobre este proyecto o su versión de escritorio:

GitHub: https://github.com/Ibaiara
LinkedIn: https://www.linkedin.com/in/ibai-araña-b2832027a